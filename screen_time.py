import datetime
import win32evtlog

REQUIRED_HOURS = 8


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
        td = daily.get(d, datetime.timedelta())
        if td.total_seconds() <= 0:
            # Zero-activity days are intentionally excluded from both total and required.
            d += datetime.timedelta(days=1)
            continue

        total += td
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
        if eid in (42, 506):
            return {"time": t, "id": eid, "type": "SLEEP", "label": "KernelPower"}
        if eid in (107, 507):
            return {"time": t, "id": eid, "type": "WAKE", "label": "KernelPower"}

    if src == "eventlog":
        if eid == 6005:
            return {"time": t, "id": eid, "type": "WAKE", "label": "EventLogBoot"}
        if eid == 6006:
            return {"time": t, "id": eid, "type": "SLEEP", "label": "EventLogShutdown"}

    return None


def collect_power_events(start_dt, end_dt):
    events = _read_log("System", start_dt, _classify_system)
    events = [e for e in events if start_dt <= e["time"] <= end_dt]
    events.sort(key=lambda x: x["time"])
    return events


def build_slots_for_day(day, all_events, now_dt):
    day_start = datetime.datetime.combine(day, datetime.time.min)
    next_midnight = day_start + datetime.timedelta(days=1)
    day_end = min(next_midnight, now_dt) if day == now_dt.date() else next_midnight

    if day_end <= day_start:
        return []

    today_events = [e for e in all_events if day_start <= e["time"] <= day_end]

    lookback_start = day_start - datetime.timedelta(hours=2)
    lookback_events = [e for e in all_events if lookback_start <= e["time"] < day_start]
    last_pre_event = lookback_events[-1] if lookback_events else None
    system_on_at_midnight = bool(last_pre_event and last_pre_event["type"] == "WAKE")

    processed = []
    for idx, event in enumerate(today_events):
        is_shutdown_artifact_wake = (
            event["type"] == "WAKE"
            and event["id"] == 107
            and idx > 0
            and today_events[idx - 1]["id"] == 42
            and (event["time"] - today_events[idx - 1]["time"]).total_seconds() <= 60
        )

        # Ignore same-second DisplayOn after DisplayOff blips.
        is_display_blip_wake = (
            event["type"] == "WAKE"
            and event["id"] == 507
            and idx > 0
            and today_events[idx - 1]["id"] == 506
            and (event["time"] - today_events[idx - 1]["time"]).total_seconds() == 0
        )

        if not is_shutdown_artifact_wake and not is_display_blip_wake:
            processed.append(event)

    slots = []
    wake_time = day_start if system_on_at_midnight else None

    for event in processed:
        if event["type"] == "WAKE" and wake_time is None:
            wake_time = event["time"]
        elif event["type"] == "SLEEP" and wake_time is not None:
            duration = (event["time"] - wake_time).total_seconds()
            if duration > 5:
                slots.append((wake_time, event["time"]))
            wake_time = None

    if wake_time is not None:
        duration = (day_end - wake_time).total_seconds()
        if duration > 5:
            slots.append((wake_time, day_end))

    return slots


def build_daily_summary(start_day, end_day, all_events, now_dt):
    daily = {}
    day_segments = {}
    all_slots = []

    day = start_day
    while day <= end_day:
        slots = build_slots_for_day(day, all_events, now_dt)
        if slots:
            total_seconds = sum((end - start).total_seconds() for start, end in slots)
            daily[day] = datetime.timedelta(seconds=total_seconds)
            day_segments[day] = slots
            all_slots.extend(slots)
        day += datetime.timedelta(days=1)

    return daily, day_segments, all_slots


def main():
    today = datetime.date.today()
    cur_start, cur_end = get_current_cycle(today)
    prev_start, prev_end = get_previous_cycle(cur_start)

    now_dt = datetime.datetime.now()
    first_needed_day = prev_start
    system_read_start = datetime.datetime.combine(first_needed_day, datetime.time.min) - datetime.timedelta(hours=2)
    system_read_end = now_dt
    events = collect_power_events(system_read_start, system_read_end)

    daily, day_segments, sessions = build_daily_summary(prev_start, today, events, now_dt)

    w = 62
    src_note = "Kernel-Power + EventLog (slot-based event processing)"
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
    today_required = required_for_day(today)
    if today_td.total_seconds() <= 0:
        today_pct = "  --"
    elif (not is_weekday(today)) and today_required.total_seconds() == 0:
        today_pct = "  w/e"
    else:
        today_pct = pct(today_td, today_required)
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

    print(f"\n  Slots built        :  {len(sessions)}")
    print(f"  Events collected   :  {len(events)}")
    print("  Security log       :  Not used")

    sep = "=" * 66
    print("")
    print(f"  {sep}")
    print("         LAPTOP WORKING HOURS REPORT")
    print(f"  {sep}")
    print(f"   Date   : {today.strftime('%A, %B %d, %Y')}")
    print(f"   Report : Up to {now_dt.strftime('%I:%M:%S %p')}  (session may still be ongoing)")
    print(f"  {sep}")
    print("")
    print("  Collecting power events...")

    today_events = [e for e in events if e["time"].date() == today]
    print(f"  Found {len(today_events)} power event(s) for this day.")

    slots_today = day_segments.get(today, [])

    if slots_today:
        print("")
        print("   #    Turned ON / Woke Up    Turned OFF / Slept     Duration")
        print("  " + ("-" * 78))

        total_today_seconds = 0
        for idx, (start, end) in enumerate(slots_today, 1):
            slot_seconds = int((end - start).total_seconds())
            total_today_seconds += slot_seconds

            h = slot_seconds // 3600
            m = (slot_seconds % 3600) // 60
            s = slot_seconds % 60

            wake_str = start.strftime('%I:%M:%S %p')
            is_open_slot = (today == now_dt.date()) and abs((end - now_dt).total_seconds()) < 2
            if is_open_slot:
                sleep_str = f"NOW {end.strftime('%I:%M:%S %p')} *"
            else:
                sleep_str = end.strftime('%I:%M:%S %p')

            dur_str = f"{h}h {m}m {s}s"
            print(f"   {idx:<4} {wake_str:<23} {sleep_str:<23} {dur_str}")

        th = total_today_seconds // 3600
        tm = (total_today_seconds % 3600) // 60
        ts = total_today_seconds % 60

        print("  " + ("-" * 78))
        print("")
        print(f"   TOTAL WORKING TIME   :   {th} hrs  {tm} min  {ts} sec")
    else:
        print("")
        print(f"  No active working slots found for {today.strftime('%A, %B %d, %Y')}.")

    print(f"  {sep}")


if __name__ == "__main__":
    main()
