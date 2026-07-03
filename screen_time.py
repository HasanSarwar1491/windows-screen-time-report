import datetime
import win32evtlog
import sys
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from rich import box

# Ensure UTF-8 output for Windows Console to avoid 'charmap' errors with blocks
if sys.stdout.encoding.lower() != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REQUIRED_HOURS = 8
console = Console()

def fmt(td):
    total_minutes = int(td.total_seconds() // 60)
    h, m = divmod(total_minutes, 60)
    return f"{h}h {m}m"

def pct(actual_td, required_td):
    if required_td.total_seconds() == 0:
        return 0.0
    return (actual_td.total_seconds() / required_td.total_seconds()) * 100

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

def get_status_color(percentage):
    if percentage >= 100: return "green"
    if percentage >= 80: return "yellow"
    if percentage > 0: return "red"
    return "white"

def main():
    today = datetime.date.today()
    cur_start, cur_end = get_current_cycle(today)
    prev_start, prev_end = get_previous_cycle(cur_start)
    now_dt = datetime.datetime.now()
    
    with console.status("[bold blue]Analyzing Windows Event Logs...", spinner="dots"):
        first_needed_day = prev_start
        system_read_start = datetime.datetime.combine(first_needed_day, datetime.time.min) - datetime.timedelta(hours=2)
        system_read_end = now_dt
        events = collect_power_events(system_read_start, system_read_end)
        daily, day_segments, sessions = build_daily_summary(prev_start, today, events, now_dt)

    # 1. Header
    console.print(Panel(
        Text.assemble(
            ("WINDOWS SCREEN TIME REPORT\n", "bold cyan"),
            (f"Target: {REQUIRED_HOURS}h/day (Weekdays) | Generated: {now_dt.strftime('%Y-%m-%d %H:%M:%S')}", "dim")
        ),
        box=box.DOUBLE,
        border_style="blue"
    ))

    # 2. Daily Log Table
    table = Table(title="Activity History (Last 30 Days)", box=box.ROUNDED, expand=True)
    table.add_column("Date", style="cyan", no_wrap=True)
    table.add_column("Time Active", justify="right")
    table.add_column("Progress Bar", ratio=1)
    table.add_column("Goal %", justify="right")
    table.add_column("Sessions", style="dim")

    for i in range(29, -1, -1):
        day = today - datetime.timedelta(days=i)
        td = daily.get(day, datetime.timedelta())
        segs = day_segments.get(day, [])
        has_activity = day in daily and td.total_seconds() > 0

        if not has_activity and not is_weekday(day): continue
        if not has_activity and day != today: continue

        label = day.strftime("%a %d %b")
        if day == today: label = f"[bold yellow]{label} (Today)[/bold yellow]"
        
        req = required_for_day(day)
        p = pct(td, req) if is_weekday(day) else (100.0 if has_activity else 0.0)
        color = get_status_color(p) if is_weekday(day) else "blue"
        
        # Simple ASCII Bar for legacy compatibility if UTF blocks fail, 
        # but here we use the characters that usually work in modern terminals
        progress_val = min(p, 100)
        bar_count = int(progress_val / 5)
        bar = "#" * bar_count + "." * (20 - bar_count)
        bar_display = f"[{color}][{bar}][/{color}]"

        session_str = f"{segs[0][0].strftime('%H:%M')}-{segs[-1][1].strftime('%H:%M')}" if segs else "-"
        
        table.add_row(
            label,
            fmt(td),
            bar_display,
            f"[{color}]{p:5.1f}%[/{color}]" if is_weekday(day) else "[blue]w/e[/blue]",
            session_str
        )

    console.print(table)

    # 3. Summary Cards
    summary_table = Table.grid(expand=True, padding=1)
    summary_table.add_column(ratio=1)
    summary_table.add_column(ratio=1)
    summary_table.add_column(ratio=1)

    t7, r7 = sum_range(daily, today - datetime.timedelta(days=6), today)
    t30, r30 = sum_range(daily, today - datetime.timedelta(days=29), today)
    tc, rc = sum_range(daily, cur_start, today)

    summary_table.add_row(
        Panel(f"[bold]{fmt(t7)}[/bold]\n[dim]Last 7 Days[/dim]", border_style="cyan", title="7D Avg"),
        Panel(f"[bold]{fmt(t30)}[/bold]\n[dim]Last 30 Days[/dim]", border_style="magenta", title="30D Total"),
        Panel(f"[bold]{fmt(tc)}[/bold]\n[dim]{pct(tc, rc):.1f}% of target[/dim]", border_style="green", title="Current Cycle")
    )
    console.print(summary_table)

    # 4. Today's Detail
    slots_today = day_segments.get(today, [])
    if slots_today:
        detail_table = Table(title=f"Today's Session Breakdown ({today.strftime('%A')})", box=box.SIMPLE_HEAD)
        detail_table.add_column("#", style="dim")
        detail_table.add_column("Start Time", style="green")
        detail_table.add_column("End Time", style="red")
        detail_table.add_column("Duration", justify="right")

        for idx, (start, end) in enumerate(slots_today, 1):
            slot_seconds = int((end - start).total_seconds())
            h, m, s = slot_seconds // 3600, (slot_seconds % 3600) // 60, slot_seconds % 60
            
            is_open = (today == now_dt.date()) and abs((end - now_dt).total_seconds()) < 5
            end_str = f"[bold]ACTIVE[/bold]" if is_open else end.strftime('%I:%M:%S %p')
            
            detail_table.add_row(str(idx), start.strftime('%I:%M:%S %p'), end_str, f"{h}h {m}m {s}s")
        
        console.print(Panel(detail_table, border_style="yellow"))
    
    # 5. Footer
    console.print(f"[dim]Stats: {len(events)} events processed across {len(sessions)} power slots. System uses Modern Standby (ID 506/507).[/dim]", justify="right")

if __name__ == "__main__":
    main()
