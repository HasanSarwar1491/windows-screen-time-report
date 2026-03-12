import datetime
import win32evtlog

REQUIRED_HOURS = 8
IDLE_GAP_MINUTES = 5


def fmt(td):
    total_minutes = int(td.total_seconds() // 60)
    h, m = divmod(total_minutes, 60)
    return f"{h}h {m}m"


def pct(actual_td, required_td):
    if required_td.total_seconds() == 0:
        return "  --"
    return f"{actual_td.total_seconds() / required_td.total_seconds() * 100:5.1f}%"


def is_weekday(d):
    return d.weekday() < 5


def required_for_day(d):
    return datetime.timedelta(hours=REQUIRED_HOURS) if is_weekday(d) else datetime.timedelta()


def sum_range(daily, start, end):
    total = datetime.timedelta()
    required = datetime.timedelta()
    d = start
    while d <= end:
        has_activity = d in daily and daily[d].total_seconds() > 0
        if has_activity:
            total += daily[d]
            required += required_for_day(d)
        d += datetime.timedelta(days=1)
    return total, required


def get_current_cycle(today):
    if today.day >= 16:
        start = today.replace(day=16)
        next_month = (today.month % 12) + 1
        year = today.year + (1 if next_month == 1 else 0)
        end = datetime.date(year, next_month, 15)
    else:
        prev_month = ((today.month - 2) % 12) + 1
        year = today.year - (1 if prev_month == 12 else 0)
        start = datetime.date(year, prev_month, 16)
        end = today.replace(day=15)
    return start, end


def get_previous_cycle(current_start):
    end = current_start - datetime.timedelta(days=1)
    prev_month = ((current_start.month - 2) % 12) + 1
    year = current_start.year - (1 if prev_month == 12 else 0)
    start = datetime.date(year, prev_month, 16)
    return start, end


def to_naive(t):
    if hasattr(t, "tzinfo") and t.tzinfo is not None:
        t = t.replace(tzinfo=None)
    return datetime.datetime(t.year, t.month, t.day, t.hour, t.minute, t.second)


def _read_log(logname, cutoff, classifier):
    events = []
    try:
        hand = win32evtlog.OpenEventLog(None, logname)
    except Exception:
        return events

    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    try:
        while True:
            records = win32evtlog.ReadEventLog(hand, flags, 0)
            if not records:
                break
            hit_cutoff = False
            for rec in records:
                t = to_naive(rec.TimeGenerated)
                if t < cutoff:
                    hit_cutoff = True
                    break
                result = classifier(rec, t)
                if result:
                    events.append(result)
            if hit_cutoff:
                break
    finally:
        win32evtlog.CloseEventLog(hand)
    return events


def _classify_system(rec, t):
    eid = rec.EventID & 0xFFFF
    src = rec.SourceName.lower()

    if "kernel-power" in src:
        if eid == 507: return ("START", t, "DisplayOn")
        if eid == 506: return ("STOP",  t, "DisplayOff")
        if eid == 42:  return ("STOP",  t, "Sleep")
        if eid == 109: return ("STOP",  t, "Shutdown(KP)")

    if "kernel-general" in src:
        if eid == 12: return ("START", t, "Boot")
        if eid == 13: return ("STOP",  t, "Shutdown")

    if "power-troubleshooter" in src:
        if eid == 1: return ("START", t, "Wake")

    if src == "eventlog":
        if eid == 6005: return ("START", t, "Boot(Evt)")
        if eid == 6006: return ("STOP",  t, "Off(Evt)")

    return None


def _classify_security(rec, t):
    eid = rec.EventID & 0xFFFF
    if eid == 4801: return ("START", t, "Unlock")
    if eid == 4800: return ("STOP",  t, "Lock")
    return None


def collect_events(days):
    cutoff = datetime.datetime.now() - datetime.timedelta(days=days)
    raw = _read_log("System", cutoff, _classify_system)
    sec = _read_log("Security", cutoff, _classify_security)
    raw += sec
    raw.sort(key=lambda x: x[1])
    return raw, len(sec) > 0


def build_sessions(events):
    """
    Strict state machine:
      OFF  + START → record start, go ON
      ON   + STOP  → record session, go OFF
      ON   + START → IGNORE (keep earliest start in this run)
      OFF  + STOP  → ignore
    """
    sessions = []
    state = "OFF"
    session_start = None

    for kind, t, _label in events:
        if state == "OFF":
            if kind == "START":
                session_start = t
                state = "ON"
        else:
            if kind == "STOP":
                sessions.append((session_start, t))
                session_start = None
                state = "OFF"

    if state == "ON" and session_start:
        sessions.append((session_start, datetime.datetime.now()))

    return sessions


def merge_short_gaps(sessions, min_gap_minutes=IDLE_GAP_MINUTES):
    if len(sessions) < 2:
        return sessions
    merged = [sessions[0]]
    for start, end in sessions[1:]:
        prev_start, prev_end = merged[-1]
        gap_min = (start - prev_end).total_seconds() / 60
        if gap_min <= min_gap_minutes:
            merged[-1] = (prev_start, end)
        else:
            merged.append((start, end))
    return merged


def distribute_to_days(sessions):
    daily = {}
    day_segments = {}
    for start, end in sessions:
        cursor = start
        while cursor.date() < end.date():
            next_midnight = datetime.datetime.combine(
                cursor.date() + datetime.timedelta(days=1), datetime.time.min
            )
            d = cursor.date()
            daily[d] = daily.get(d, datetime.timedelta()) + (next_midnight - cursor)
            day_segments.setdefault(d, []).append((cursor, next_midnight))
            cursor = next_midnight
        d = cursor.date()
        daily[d] = daily.get(d, datetime.timedelta()) + (end - cursor)
        day_segments.setdefault(d, []).append((cursor, end))
    return daily, day_segments


def main():
    today = datetime.date.today()
    cur_start, cur_end = get_current_cycle(today)
    prev_start, prev_end = get_previous_cycle(cur_start)

    lookback = (today - prev_start).days + 5
    (events, has_security) = collect_events(days=lookback)
    raw_sessions = build_sessions(events)
    sessions = merge_short_gaps(raw_sessions)
    daily, day_segments = distribute_to_days(sessions)

    w = 62
    src_note = "Power + Lock/Unlock" if has_security else "Power + Display (run as Admin for lock events)"
    print("=" * w)
    print("  DAILY SCREEN-ON TIME (Last 30 Days)")
    print(f"  Weekdays: Mon-Fri | Required: 8h/day")
    print(f"  Source: {src_note}")
    print("=" * w)

    for i in range(29, -1, -1):
        day = today - datetime.timedelta(days=i)
        td = daily.get(day, datetime.timedelta())
        segs = day_segments.get(day, [])
        has_activity = day in daily and td.total_seconds() > 0

        if not has_activity and not is_weekday(day):
            continue
        if not has_activity and day != today:
            continue

        label = day.strftime("%a %d %b %Y")

        if not is_weekday(day):
            tag = "  [  w/e ]"
        elif has_activity:
            tag = f"  [{pct(td, required_for_day(day))}]"
        else:
            tag = "  [  --  ]"

        span = ""
        if segs:
            fs = segs[0][0].strftime("%H:%M")
            ls = segs[-1][1].strftime("%H:%M")
            span = f"  ({fs} - {ls})"

        marker = "  <-- Today" if day == today else ""
        print(f"  {label}  :  {fmt(td):>7}{tag}{span}{marker}")

        if len(segs) > 1:
            parts = [f"{s.strftime('%H:%M')}-{e.strftime('%H:%M')}" for s, e in segs]
            print(f"                        sessions: {', '.join(parts)}")

    print("\n" + "=" * w)
    print("  SUMMARY")
    print("=" * w)

    today_td = daily.get(today, datetime.timedelta())
    today_pct = pct(today_td, required_for_day(today)) if is_weekday(today) else "  w/e"
    print(f"  Today              :  {fmt(today_td):>7}   [{today_pct}]")

    t7, r7 = sum_range(daily, today - datetime.timedelta(days=6), today)
    print(f"  Last  7 Days       :  {fmt(t7):>7}   [{pct(t7, r7)}]")

    t30, r30 = sum_range(daily, today - datetime.timedelta(days=29), today)
    print(f"  Last 30 Days       :  {fmt(t30):>7}   [{pct(t30, r30)}]")

    print("-" * w)

    tc, rc = sum_range(daily, cur_start, today)
    cycle_lbl = f"{cur_start.strftime('%d %b')} - {cur_end.strftime('%d %b %Y')}"
    print(f"  This Cycle         :  {fmt(tc):>7}   [{pct(tc, rc)}]   ({cycle_lbl})")

    tp, rp = sum_range(daily, prev_start, prev_end)
    prev_lbl = f"{prev_start.strftime('%d %b')} - {prev_end.strftime('%d %b %Y')}"
    print(f"  Previous Cycle     :  {fmt(tp):>7}   [{pct(tp, rp)}]   ({prev_lbl})")

    print("=" * w)

    print(f"\n  Raw sessions       :  {len(raw_sessions)}")
    print(f"  After merge (<{IDLE_GAP_MINUTES}m) :  {len(sessions)}")
    print(f"  Events collected   :  {len(events)}")
    print(f"  Security log       :  {'Yes' if has_security else 'No (need Admin)'}")

    print(f"\n{'=' * w}")
    print(f"  TODAY'S RAW EVENTS ({today.strftime('%d %b %Y')})")
    print(f"{'=' * w}")
    today_events = [(k, t, lbl) for k, t, lbl in events if t.date() == today]
    if today_events:
        for kind, t, lbl in today_events:
            arrow = ">>" if kind == "START" else "<<"
            print(f"  {arrow} {t.strftime('%H:%M:%S')}  {kind:<5}  {lbl}")
    else:
        print("  (no events)")
    print(f"{'=' * w}")


if __name__ == "__main__":
    main()
