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
from rich.progress import ProgressBar
from rich.columns import Columns

# Ensure UTF-8 output
if sys.stdout.encoding.lower() != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REQUIRED_HOURS = 8
MONTHLY_TARGET = 176
SLOT_MINUTES = 10 
SLOT_THRESHOLD_MINUTES = 2 
console = Console()
LOG_FILE = os.path.expanduser("~/.screen_time/activity.log")

def fmt(td):
    if isinstance(td, (int, float)): h, m = divmod(int(td), 60)
    else:
        total_minutes = int(td.total_seconds() // 60)
        h, m = divmod(total_minutes, 60)
    return f"{h}h {m}m"

def is_weekday(d): return d.weekday() < 5

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
                if t < cutoff: hit_cutoff = True; break
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
    activity = {}
    if not os.path.exists(LOG_FILE): return activity
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split(",", 3)
                if len(parts) < 2: continue
                try: dt = datetime.datetime.fromisoformat(parts[0])
                except: continue
                d = dt.date()
                if d not in activity: activity[d] = {}
                state, app, title = parts[1], parts[2] if len(parts) > 2 else "Unknown", parts[3] if len(parts) > 3 else ""
                activity[d][dt.replace(second=0, microsecond=0)] = (state, app, title)
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
        is_shutdown_artifact_wake = (event["type"] == "WAKE" and event["id"] == 107 and idx > 0 and today_events[idx - 1]["id"] == 42 and (event["time"] - today_events[idx - 1]["time"]).total_seconds() <= 60)
        is_display_blip_wake = (event["type"] == "WAKE" and event["id"] == 507 and idx > 0 and today_events[idx - 1]["id"] == 506 and (event["time"] - today_events[idx - 1]["time"]).total_seconds() == 0)
        if not is_shutdown_artifact_wake and not is_display_blip_wake: processed.append(event)

    slots = []
    wake_time = day_start if system_on_at_midnight else None
    for event in processed:
        if event["type"] == "WAKE" and wake_time is None: wake_time = event["time"]
        elif event["type"] == "SLEEP" and wake_time is not None:
            if (event["time"] - wake_time).total_seconds() > 5: slots.append((wake_time, event["time"]))
            wake_time = None
    if wake_time is not None: slots.append((wake_time, day_end))

    day_activity = activity_data.get(day, {})
    if day_activity:
        known_on_minutes = set()
        for s, e in slots:
            curr = s.replace(second=0, microsecond=0)
            while curr <= e: known_on_minutes.add(curr); curr += datetime.timedelta(minutes=1)
        for am in sorted(day_activity.keys()):
            if am not in known_on_minutes: slots.append((am, am + datetime.timedelta(minutes=1)))
        slots.sort(); merged = []
        if slots:
            curr_s, curr_e = slots[0]
            for next_s, next_e in slots[1:]:
                if next_s <= curr_e: curr_e = max(curr_e, next_e)
                else: merged.append((curr_s, curr_e)); curr_s, curr_e = next_s, next_e
            merged.append((curr_s, curr_e))
        slots = merged

    system_seconds = sum((e-s).total_seconds() for s, e in slots)
    active_seconds = 0
    hourly_activity = [0] * 24
    apps = []
    for start, end in slots:
        curr = start.replace(second=0, microsecond=0)
        while curr <= end:
            entry = day_activity.get(curr)
            if entry and entry[0] == "active": 
                active_seconds += 60
                hourly_activity[curr.hour] += 1
                apps.append(entry[1])
            curr += datetime.timedelta(minutes=1)
    
    # --- Teramind Discrete Logic ---
    EXPECTED_SAMPLES = (SLOT_MINUTES * 60) // 30
    buckets = {}
    for ts, data in day_activity.items():
        bucket_ts = ts.replace(minute=(ts.minute // SLOT_MINUTES) * SLOT_MINUTES)
        if bucket_ts not in buckets: buckets[bucket_ts] = []
        buckets[bucket_ts].append(data)
    
    active_slots_count = 0
    total_intensity = 0
    discrete_slots = []
    for b_ts in sorted(buckets.keys()):
        samples = buckets[b_ts]
        active_samples = [s for s in samples if s[0] == 'active']
        active_minutes = len(active_samples) / 2 if len(active_samples) > 0 else 0
        intensity = (len(active_samples) / EXPECTED_SAMPLES) * 100
        if active_minutes >= SLOT_THRESHOLD_MINUTES:
            active_slots_count += 1
            total_intensity += intensity
            discrete_slots.append({"start": b_ts, "end": b_ts + datetime.timedelta(minutes=SLOT_MINUTES), "intensity": intensity})
    
    discrete_min = active_slots_count * SLOT_MINUTES
    avg_intensity = (total_intensity / active_slots_count) if active_slots_count > 0 else 0
    ultra_min = active_seconds / 60

    return {
        "slots": slots,
        "system_td": datetime.timedelta(seconds=system_seconds),
        "active_td": datetime.timedelta(seconds=active_seconds),
        "hourly": hourly_activity,
        "apps": Counter(apps).most_common(5),
        "discrete_min": discrete_min,
        "avg_intensity": avg_intensity,
        "discrete_slots": discrete_slots,
        "ultra_min": ultra_min
    }

def get_calendar_month_cycle(today):
    start = today.replace(day=1)
    if today.month == 12: end = today.replace(year=today.year + 1, month=1, day=1) - datetime.timedelta(days=1)
    else: end = today.replace(month=today.month + 1, day=1) - datetime.timedelta(days=1)
    return start, end

def main():
    today = datetime.date.today()
    cal_start, cal_end = get_calendar_month_cycle(today)
    now_dt = datetime.datetime.now()
    
    with console.status("[bold blue]Generating Hybrid High-Fidelity Report...", spinner="dots"):
        events = _read_log("System", datetime.datetime.combine(today - datetime.timedelta(days=30), datetime.time.min), _classify_system)
        events.sort(key=lambda x: x["time"])
        activity_data = load_activity_logs()
        daily = {}
        day = today - datetime.timedelta(days=29)
        while day <= today:
            res = build_slots_for_day(day, events, activity_data, now_dt)
            if res["system_td"].total_seconds() > 0 or res["discrete_min"] > 0: daily[day] = res
            day += datetime.timedelta(days=1)

    # 1. Header
    console.print(Panel(Text.assemble(("HYBRID TERAMIND & SYSTEM DASHBOARD\n", "bold green"), (f"Continuous & Discrete Analysis | {now_dt.strftime('%H:%M:%S')}", "dim")), box=box.DOUBLE, border_style="green"))

    # 2. Key Stats
    today_res = daily.get(today, {"system_td": datetime.timedelta(), "active_td": datetime.timedelta(), "hourly": [0]*24, "apps": [], "discrete_min": 0, "avg_intensity": 0, "slots": [], "discrete_slots": [], "ultra_min": 0})
    
    score_color = "bright_green" if today_res["avg_intensity"] > 70 else "yellow" if today_res["avg_intensity"] > 40 else "red"
    stats_row = Table.grid(expand=True, padding=1)
    stats_row.add_column(ratio=1); stats_row.add_column(ratio=1); stats_row.add_column(ratio=1)
    stats_row.add_row(
        Panel(f"[bold cyan]{fmt(today_res['discrete_min'])}[/bold cyan]\n[dim]Enterprise Quantized[/dim]", border_style="cyan", title="Work Time"),
        Panel(f"[bold {score_color}]{today_res['avg_intensity']:.1f}%[/bold {score_color}]\n[dim]Mean Intensity Score[/dim]", border_style=score_color, title="Efficiency"),
        Panel(f"[bold magenta]{fmt(today_res['ultra_min'])}[/bold magenta]\n[dim]True Active Duration[/dim]", border_style="magenta", title="Ultra-Realistic")
    )
    console.print(stats_row)

    # 3. Monthly Goal
    t_month_min = sum(res["discrete_min"] for d, res in daily.items() if cal_start <= d <= today)
    hours_done = t_month_min / 60
    month_pct = min((hours_done / MONTHLY_TARGET) * 100, 100)
    month_progress = ProgressBar(total=MONTHLY_TARGET, completed=hours_done, width=None, finished_style="bright_green")
    console.print(Panel(Columns([f"[bold bright_blue]MTD: {hours_done:.1f}h / {MONTHLY_TARGET}h ({month_pct:.1f}%)[/bold bright_blue]", month_progress, f"[dim]Remaining: {max(0, MONTHLY_TARGET - hours_done):.1f}h[/dim]"], expand=True, align="left"), title="Monthly Target Tracker (176h)", border_style="bright_blue"))

    # 4. Apps
    if today_res["apps"]:
        app_table = Table(box=box.SIMPLE, expand=True)
        app_table.add_column("Application Name", style="cyan"); app_table.add_column("Minutes (Approx)", justify="right", style="green")
        for app, count in today_res["apps"]: app_table.add_row(app, f"{count}m")
        console.print(Panel(app_table, title="Application Focus Distribution", border_style="dim"))

    # 5. Hourly Map
    heatmap = Text("Intensity: ", style="bold")
    for h in range(24):
        intensity = today_res["hourly"][h]
        char = "█" if intensity > 45 else "▆" if intensity > 30 else "▄" if intensity > 10 else "▂" if intensity > 0 else " "
        color = "bright_green" if intensity > 30 else "yellow" if intensity > 10 else "dim"
        heatmap.append(f" {h:02d}", style="dim"); heatmap.append(char, style=color)
    console.print(Panel(heatmap, title="Hourly Work Timeline", border_style="dim"))

    # 6. Session Breakdown
    if today_res["slots"]:
        session_table = Table(box=box.SIMPLE, expand=True)
        session_table.add_column("#", style="dim"); session_table.add_column("Start Time", style="cyan"); session_table.add_column("End Time", style="cyan"); session_table.add_column("Duration", justify="right"); session_table.add_column("Activity %", justify="right")
        for i, (s, e) in enumerate(today_res["slots"]):
            dur = e - s; act_m = 0; curr = s.replace(second=0, microsecond=0)
            while curr <= e:
                entry = activity_data.get(today, {}).get(curr)
                if entry and entry[0] == "active": act_m += 1
                curr += datetime.timedelta(minutes=1)
            act_pct = (act_m / (dur.total_seconds()/60) * 100) if dur.total_seconds() > 0 else 0
            act_color = "bright_green" if act_pct > 75 else "yellow" if act_pct > 40 else "red"
            session_table.add_row(str(i+1), s.strftime("%H:%M:%S"), e.strftime("%H:%M:%S"), fmt(dur), f"[{act_color}]{act_pct:.0f}%[/{act_color}]")
        console.print(Panel(session_table, title="Continuous Session Breakdown (Today)", border_style="dim"))

    # 7. Discrete Slot
    if today_res["discrete_slots"]:
        slot_table = Table(box=box.SIMPLE, expand=True)
        slot_table.add_column("10-Min Slot", style="dim"); slot_table.add_column("Intensity Bar", ratio=1); slot_table.add_column("Score", justify="right")
        for slot in today_res["discrete_slots"][-8:]:
            bar_w = 20; filled = int((slot['intensity'] / 100) * bar_w); color = "bright_green" if slot['intensity'] > 70 else "yellow" if slot['intensity'] > 40 else "red"
            bar = f"[{color}]" + "█" * filled + "[/]" + "░" * (bar_w - filled)
            slot_table.add_row(f"{slot['start'].strftime('%H:%M')} - {slot['end'].strftime('%H:%M')}", bar, f"[{color}]{slot['intensity']:.0f}%[/]")
        console.print(Panel(slot_table, title="Teramind Discrete Slot Analysis (Recent)", border_style="dim"))

    # 8. History
    table = Table(title="Historical Accountability (Last 30 Days)", box=box.ROUNDED, expand=True)
    table.add_column("Date", style="cyan"); table.add_column("System On", justify="right"); table.add_column("Focused", style="bright_green", justify="right"); table.add_column("Quantized", style="bright_blue", justify="right"); table.add_column("Intensity", justify="right"); table.add_column("Activity Intensity", ratio=1)
    for i in range(29, -1, -1):
        day = today - datetime.timedelta(days=i)
        if day not in daily: continue
        res = daily[day]
        s_color = "bright_green" if res["avg_intensity"] > 70 else "yellow" if res["avg_intensity"] > 40 else "dim"
        bar_width = 20; act_bars = int((res["discrete_min"] / (REQUIRED_HOURS*60)) * bar_width)
        intensity_bar = "[bright_green]" + "█" * act_bars + "[/bright_green]" + " " * max(0, bar_width - act_bars)
        table.add_row(day.strftime("%a %d %b"), fmt(res["system_td"]), fmt(res["active_td"]), fmt(res["discrete_min"]), f"[{s_color}]{res['avg_intensity']:.1f}%[/{s_color}]", intensity_bar)
    console.print(table)
    console.print(f"[dim]Hybrid tracking: Windows Power Events (Presence) + 10m Discrete Buckets (Accountability).[/dim]", justify="right")

if __name__ == "__main__": main()
