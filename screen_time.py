import datetime
import win32evtlog
import sys
import os
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from rich import box
from collections import Counter

# Ensure UTF-8 output
if sys.stdout.encoding.lower() != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REQUIRED_HOURS = 8
console = Console()
LOG_FILE = os.path.expanduser("~/.screen_time/activity.log")

def fmt(td):
    total_minutes = int(td.total_seconds() // 60)
    h, m = divmod(total_minutes, 60)
    return f"{h}h {m}m"

def pct(actual_td, required_td):
    if required_td.total_seconds() == 0: return 0.0
    return (actual_td.total_seconds() / required_td.total_seconds()) * 100

def is_weekday(d): return d.weekday() < 5

def required_for_day(d):
    return datetime.timedelta(hours=REQUIRED_HOURS) if is_weekday(d) else datetime.timedelta()

def sum_range(daily, start, end):
    total, required = datetime.timedelta(), datetime.timedelta()
    d = start
    while d <= end:
        td = daily.get(d, {}).get("system", datetime.timedelta())
        if td.total_seconds() > 0:
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
    if hasattr(t, "tzinfo") and t.tzinfo is not None: t = t.replace(tzinfo=None)
    return datetime.datetime(t.year, t.month, t.day, t.hour, t.minute, t.second)

def _read_log(logname, cutoff, classifier):
    events = []
    try: hand = win32evtlog.OpenEventLog(None, logname)
    except: return events
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    try:
        while True:
            records = win32evtlog.ReadEventLog(hand, flags, 0)
            if not records: break
            hit_cutoff = False
            for rec in records:
                t = to_naive(rec.TimeGenerated)
                if t < cutoff:
                    hit_cutoff = True
                    break
                result = classifier(rec, t)
                if result: events.append(result)
            if hit_cutoff: break
    finally: win32evtlog.CloseEventLog(hand)
    return events

def _classify_system(rec, t):
    eid = rec.EventID & 0xFFFF
    src = rec.SourceName.lower()
    if "kernel-power" in src:
        if eid in (42, 506): return {"time": t, "id": eid, "type": "SLEEP"}
        if eid in (107, 507): return {"time": t, "id": eid, "type": "WAKE"}
    if src == "eventlog":
        if eid == 6005: return {"time": t, "id": eid, "type": "WAKE"}
        if eid == 6006: return {"time": t, "id": eid, "type": "SLEEP"}
    return None

def load_activity_logs():
    activity = {} # date -> {minute_timestamp -> (state, app)}
    if not os.path.exists(LOG_FILE): return activity
    try:
        with open(LOG_FILE, "r") as f:
            for line in f:
                parts = line.strip().split(",")
                if len(parts) < 2: continue
                dt = datetime.datetime.fromisoformat(parts[0])
                d = dt.date()
                if d not in activity: activity[d] = {}
                state = parts[1]
                app = parts[2] if len(parts) > 2 else "Unknown"
                activity[d][dt.replace(second=0, microsecond=0)] = (state, app)
    except: pass
    return activity

def build_slots_for_day(day, all_events, activity_data, now_dt):
    day_start = datetime.datetime.combine(day, datetime.time.min)
    next_midnight = day_start + datetime.timedelta(days=1)
    day_end = min(next_midnight, now_dt) if day == now_dt.date() else next_midnight

    lookback_events = [e for e in all_events if (day_start - datetime.timedelta(hours=2)) <= e["time"] < day_start]
    system_on_at_midnight = bool(lookback_events and lookback_events[-1]["type"] == "WAKE")
    
    today_events = [e for e in all_events if day_start <= e["time"] <= day_end]
    
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
        if event["type"] == "WAKE" and wake_time is None: wake_time = event["time"]
        elif event["type"] == "SLEEP" and wake_time is not None:
            if (event["time"] - wake_time).total_seconds() > 5:
                slots.append((wake_time, event["time"]))
            wake_time = None
    if wake_time is not None: slots.append((wake_time, day_end))

    system_seconds = sum((e-s).total_seconds() for s, e in slots)
    
    active_seconds = 0
    hourly_activity = [0] * 24
    apps = []
    if day in activity_data:
        day_activity = activity_data[day]
        for start, end in slots:
            curr = start.replace(second=0, microsecond=0)
            while curr <= end:
                entry = day_activity.get(curr)
                if entry and entry[0] == "active": 
                    active_seconds += 60
                    hourly_activity[curr.hour] += 1
                    apps.append(entry[1])
                curr += datetime.timedelta(minutes=1)
    
    active_seconds = min(active_seconds, system_seconds)
    top_apps = Counter(apps).most_common(5)
    return slots, datetime.timedelta(seconds=system_seconds), datetime.timedelta(seconds=active_seconds), hourly_activity, top_apps

def main():
    today = datetime.date.today()
    cur_start, cur_end = get_current_cycle(today)
    prev_start, prev_end = get_previous_cycle(cur_start)
    now_dt = datetime.datetime.now()
    
    with console.status("[bold blue]Generating High-Fidelity Teramind Report...", spinner="dots"):
        events = _read_log("System", datetime.datetime.combine(prev_start, datetime.time.min), _classify_system)
        events.sort(key=lambda x: x["time"])
        activity_data = load_activity_logs()
        
        daily = {}
        day = prev_start
        while day <= today:
            slots, sys_td, act_td, hourly, top_apps = build_slots_for_day(day, events, activity_data, now_dt)
            if sys_td.total_seconds() > 0:
                daily[day] = {"system": sys_td, "active": act_td, "slots": slots, "hourly": hourly, "apps": top_apps}
            day += datetime.timedelta(days=1)

    # 1. Header
    console.print(Panel(Text.assemble(
        ("TERAMIND HIGH-FIDELITY DASHBOARD\n", "bold green"),
        (f"Focused Window Tracking & Productivity Score | {now_dt.strftime('%H:%M:%S')}", "dim")
    ), box=box.DOUBLE, border_style="green"))

    # 2. Key Stats Panel
    today_data = daily.get(today, {"system": datetime.timedelta(), "active": datetime.timedelta(), "hourly": [0]*24, "apps": []})
    today_sys, today_act = today_data["system"], today_data["active"]
    score = (today_act.total_seconds() / today_sys.total_seconds() * 100) if today_sys.total_seconds() > 0 else 0
    score_color = "green" if score > 75 else "yellow" if score > 50 else "red"

    stats_row = Table.grid(expand=True, padding=1)
    stats_row.add_column(ratio=1); stats_row.add_column(ratio=1); stats_row.add_column(ratio=1)
    stats_row.add_row(
        Panel(f"[bold]{fmt(today_act)}[/bold]\n[dim]Productive Time[/dim]", border_style="green", title="Focused Active"),
        Panel(f"[bold]{score:.1f}%[/bold]\n[dim]Intensity Score[/dim]", border_style=score_color, title="Accountability"),
        Panel(f"[bold]{fmt(today_sys - today_act)}[/bold]\n[dim]Idle/Background[/dim]", border_style="red", title="Unaccounted")
    )
    console.print(stats_row)

    # 3. Top Focused Applications (Teramind Feature)
    if today_data["apps"]:
        app_table = Table(title="Top Focused Applications (Today)", box=box.SIMPLE, expand=True)
        app_table.add_column("Application Name", style="cyan")
        app_table.add_column("Active Minutes", justify="right", style="green")
        app_table.add_column("Share", justify="right")
        
        total_m = sum(count for _, count in today_data["apps"])
        for app, count in today_data["apps"]:
            app_table.add_row(app, f"{count}m", f"{(count/total_m*100):.1f}%")
        console.print(Panel(app_table, border_style="blue"))

    # 4. Hourly Intensity Map
    heatmap = Text("Work Intensity Map: ", style="bold")
    for h in range(24):
        intensity = today_data["hourly"][h]
        char = "█" if intensity > 45 else "▓" if intensity > 30 else "▒" if intensity > 10 else "░" if intensity > 0 else " "
        color = "green" if intensity > 30 else "yellow" if intensity > 10 else "dim"
        heatmap.append(f" {h:02d}", style="dim")
        heatmap.append(char, style=color)
    console.print(Panel(heatmap, title="Hourly Activity Timeline", border_style="dim"))

    # 5. Historical Log
    table = Table(title="Historical Accountability (Last 30 Days)", box=box.ROUNDED, expand=True)
    table.add_column("Date", style="cyan")
    table.add_column("System", justify="right")
    table.add_column("Focused", style="green", justify="right")
    table.add_column("Score", justify="right")
    table.add_column("Intensity Bar", ratio=1)

    for i in range(29, -1, -1):
        day = today - datetime.timedelta(days=i)
        if day not in daily and not is_weekday(day): continue
        data = daily.get(day, {"system": datetime.timedelta(), "active": datetime.timedelta(), "slots": []})
        sys_td, act_td = data["system"], data["active"]
        if sys_td.total_seconds() == 0 and day != today: continue

        eff = (act_td.total_seconds() / sys_td.total_seconds() * 100) if sys_td.total_seconds() > 0 else 0
        eff_color = "green" if eff > 75 else "yellow" if eff > 40 else "dim"
        
        bar_width = 15
        act_bars = int((act_td.total_seconds() / (REQUIRED_HOURS*3600)) * bar_width) if is_weekday(day) else int(bar_width/2)
        sys_bars = int((sys_td.total_seconds() / (REQUIRED_HOURS*3600)) * bar_width) if is_weekday(day) else int(bar_width/2)
        bar = "[green]" + "█" * act_bars + "[/green]" + "░" * max(0, sys_bars - act_bars) + " " * max(0, bar_width - sys_bars)

        table.add_row(day.strftime("%a %d %b"), fmt(sys_td), fmt(act_td) if act_td.total_seconds() > 0 else "-", f"[{eff_color}]{eff:.1f}%[/{eff_color}]", bar)
    console.print(table)

if __name__ == "__main__":
    main()
