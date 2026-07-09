import datetime
import io
import os
import sys
from collections import Counter

import win32evtlog
from rich import box
from rich.columns import Columns
from rich.console import Console
from rich.panel import Panel
from rich.progress import ProgressBar
from rich.table import Table
from rich.text import Text

# Ensure UTF-8 output (safe when encoding is None in redirected shells)
_stdout_encoding = (sys.stdout.encoding or "").lower()
if _stdout_encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

REQUIRED_HOURS = 8
CYCLE_TARGET_HOURS = 176
CYCLE_START_DAY = 16
SLOT_MINUTES = 5
SLOT_THRESHOLD_MINUTES = 0.0833

console = Console()
LOG_FILE = os.path.expanduser("~/.screen_time/activity.log")


# -----------------------------
# Formatting & date utilities
# -----------------------------
def fmt(td_or_minutes):
    if isinstance(td_or_minutes, (int, float)):
        h, m = divmod(int(td_or_minutes), 60)
    else:
        total_minutes = int(td_or_minutes.total_seconds() // 60)
        h, m = divmod(total_minutes, 60)
    return f"{h}h {m}m"


def to_naive(t):
    if hasattr(t, "tzinfo") and t.tzinfo is not None:
        t = t.replace(tzinfo=None)
    return datetime.datetime(t.year, t.month, t.day, t.hour, t.minute, t.second)


def shift_month(year, month, delta):
    idx = (year * 12 + (month - 1)) + delta
    new_year = idx // 12
    new_month = (idx % 12) + 1
    return new_year, new_month


def get_cycle_range(today, cycle_start_day=CYCLE_START_DAY):
    """
    Returns a cycle of [16th -> 15th] around 'today'.
    If today >= 16, cycle starts this month 16th, else previous month 16th.
    """
    if today.day >= cycle_start_day:
        start = datetime.date(today.year, today.month, cycle_start_day)
    else:
        py, pm = shift_month(today.year, today.month, -1)
        start = datetime.date(py, pm, cycle_start_day)

    ny, nm = shift_month(start.year, start.month, 1)
    end = datetime.date(ny, nm, cycle_start_day) - datetime.timedelta(days=1)
    return start, end


def summarize_period(daily, start_day, end_day):
    rows = [res for d, res in daily.items() if start_day <= d <= end_day]
    if not rows:
        return {
            "days": 0,
            "system_td": datetime.timedelta(),
            "active_td": datetime.timedelta(),
            "quantized_min": 0,
            "avg_intensity": 0.0,
        }

    system_td = sum((r["system_td"] for r in rows), datetime.timedelta())
    active_td = sum((r["active_td"] for r in rows), datetime.timedelta())
    quantized_min = sum(r["discrete_min"] for r in rows)
    avg_intensity = sum(r["avg_intensity"] for r in rows) / len(rows)

    return {
        "days": len(rows),
        "system_td": system_td,
        "active_td": active_td,
        "quantized_min": quantized_min,
        "avg_intensity": avg_intensity,
    }


def get_day_bounds(slots):
    if not slots:
        return "--:--", "--:--"
    return slots[0][0].strftime("%H:%M"), slots[-1][1].strftime("%H:%M")


# -----------------------------
# Event & activity loading
# -----------------------------
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
                
                type_name, evt_id = classifier(rec, t)
                if type_name:
                    events.append({"time": t, "type": type_name, "id": evt_id})
            
            if hit_cutoff:
                break
    finally:
        win32evtlog.CloseEventLog(hand)
    
    return events


def _classify_system(rec, t):
    # Standard Windows System Log Source: Kernel-General (1, 12, 13) or Power-Troubleshooter (1)
    # 107 = System Resume, 42 = Sleep, 1 = Wake, 12/13 = Startup/Shutdown
    eid = rec.EventID & 0xFFFF
    src = rec.SourceName
    
    if src == "Microsoft-Windows-Kernel-General":
        if eid == 12: return "WAKE", eid   # OS Startup
        if eid == 13: return "SLEEP", eid  # OS Shutdown
    if src == "Microsoft-Windows-Power-Troubleshooter":
        if eid == 1: return "WAKE", eid    # Resume
    if src == "Microsoft-Windows-Kernel-Power":
        if eid == 42: return "SLEEP", eid  # Sleep
        if eid == 107: return "WAKE", eid  # System Resume
        
    return None, None


def load_activity_logs():
    activity = {}
    if not os.path.exists(LOG_FILE):
        return activity

    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split(",", 4)
                if len(parts) < 5:
                    continue

                try:
                    dt = datetime.datetime.fromisoformat(parts[0])
                except Exception:
                    continue

                d = dt.date()
                if d not in activity:
                    activity[d] = {}

                # parts[1]=state, parts[2]=app, parts[3]=title, parts[4]=active_seconds
                activity[d][dt.replace(second=0, microsecond=0)] = (
                    parts[1], parts[2], parts[3], int(parts[4])
                )
    except Exception:
        pass
    return activity


def build_slots_for_day(day, all_events, activity_data, now_dt):
    day_start = datetime.datetime.combine(day, datetime.time.min)
    next_midnight = day_start + datetime.timedelta(days=1)
    day_end = min(next_midnight, now_dt) if day == now_dt.date() else next_midnight

    lookback_events = [
        e for e in all_events
        if (day_start - datetime.timedelta(hours=2)) <= e["time"] < day_start
    ]
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
        if not is_shutdown_artifact_wake:
            processed.append(event)

    slots = []
    current_wake = day_start if system_on_at_midnight else None
    
    for event in processed:
        if event["type"] == "WAKE" and not current_wake:
            current_wake = event["time"]
        elif event["type"] == "SLEEP" and current_wake:
            slots.append((current_wake, event["time"]))
            current_wake = None
            
    if current_wake:
        slots.append((current_wake, day_end))

    day_activity = activity_data.get(day, {})
    
    # NEW: Ensure slots include any activity recorded in the logger even if power logs missed it
    if day_activity:
        known_on_minutes = set()
        for s, e in slots:
            curr = s.replace(second=0, microsecond=0)
            while curr <= e:
                known_on_minutes.add(curr)
                curr += datetime.timedelta(minutes=1)

        for am in sorted(day_activity.keys()):
            if am not in known_on_minutes:
                # Add a 1-minute slot for any recorded activity not covered by power events
                slots.append((am, am + datetime.timedelta(minutes=1)))

        slots.sort()
        # Merge overlapping slots and account for activity even if power log is messy
        merged = []
        if slots:
            curr_s, curr_e = slots[0]
            for next_s, next_e in slots[1:]:
                if next_s <= curr_e:
                    curr_e = max(curr_e, next_e)
                else:
                    merged.append((curr_s, curr_e))
                    curr_s, curr_e = next_s, next_e
            merged.append((curr_s, curr_e))
        slots = merged

    system_seconds = sum((e - s).total_seconds() for s, e in slots)

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

    # High-Fidelity Discrete Logic
    buckets = {}
    for ts, data in day_activity.items():
        # Correctly bucket by stripping seconds to match SLOT_MINUTES alignment
        bucket_ts = ts.replace(minute=(ts.minute // SLOT_MINUTES) * SLOT_MINUTES, second=0, microsecond=0)
        buckets.setdefault(bucket_ts, []).append(data)

    active_slots_count = 0
    total_intensity = 0
    discrete_slots = []

    for b_ts in sorted(buckets.keys()):
        samples = buckets[b_ts]
        # Calculate intensity based on active seconds across all samples in bucket
        # Each sample is POLL_SECONDS (30s)
        total_active_seconds = sum(s[3] for s in samples)
        expected_seconds = (SLOT_MINUTES * 60)
        
        # Scaling for activity saturation threshold. 
        # Most enterprise trackers reach 100% activity if you are active ~50% of the time.
        # We multiply by 2.0 (or divide expected by 0.5) to make it less aggressive.
        intensity = (total_active_seconds / (expected_seconds * 0.5)) * 100
        # Cap intensity at 100%
        intensity = min(100.0, intensity)

        # We count a slot as 'worked' if it has at least some significant activity
        # If intensity > 1.0% or meets a threshold
        if intensity > 1.0: 
            active_slots_count += 1
            total_intensity += intensity
            discrete_slots.append(
                {
                    "start": b_ts,
                    "end": b_ts + datetime.timedelta(minutes=SLOT_MINUTES),
                    "intensity": intensity,
                }
            )

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
        "ultra_min": ultra_min,
    }


# -----------------------------
# Main dashboard
# -----------------------------
def main():
    today = datetime.date.today()
    now_dt = datetime.datetime.now()
    cycle_start, cycle_end = get_cycle_range(today)

    # Ensure enough history for: last 30 days + current cycle
    lookback_start = min(today - datetime.timedelta(days=29), cycle_start)
    cutoff_dt = datetime.datetime.combine(lookback_start - datetime.timedelta(days=2), datetime.time.min)

    with console.status("[bold blue]Generating Hybrid High-Fidelity Report...", spinner="dots"):
        events = _read_log("System", cutoff_dt, _classify_system)
        events.sort(key=lambda x: x["time"])

        activity_data = load_activity_logs()

        daily = {}
        day = lookback_start
        while day <= today:
            res = build_slots_for_day(day, events, activity_data, now_dt)
            if res["system_td"].total_seconds() > 0 or res["discrete_min"] > 0:
                daily[day] = res
            day += datetime.timedelta(days=1)

    # 1) Header
    console.print(
        Panel(
            Text.assemble(
                ("HYBRID ACTIVITY & SYSTEM DASHBOARD\n", "bold green"),
                (f"Continuous & Discrete Analysis | {now_dt.strftime('%H:%M:%S')}", "dim"),
            ),
            box=box.DOUBLE,
            border_style="green",
        )
    )

    # 2) Key stats (today)
    today_res = daily.get(
        today,
        {
            "system_td": datetime.timedelta(),
            "active_td": datetime.timedelta(),
            "hourly": [0] * 24,
            "apps": [],
            "discrete_min": 0,
            "avg_intensity": 0,
            "slots": [],
            "discrete_slots": [],
            "ultra_min": 0,
        },
    )

    score_color = (
        "bright_green" if today_res["avg_intensity"] > 70
        else "yellow" if today_res["avg_intensity"] > 40
        else "red"
    )

    stats_row = Table.grid(expand=True, padding=1)
    stats_row.add_column(ratio=1)
    stats_row.add_column(ratio=1)
    stats_row.add_column(ratio=1)
    stats_row.add_row(
        Panel(
            f"[bold cyan]{fmt(today_res['discrete_min'])}[/bold cyan]\n[dim]Enterprise Quantized[/dim]",
            border_style="cyan",
            title="Work Time",
        ),
        Panel(
            f"[bold {score_color}]{today_res['avg_intensity']:.1f}%[/bold {score_color}]\n[dim]Mean Intensity Score[/dim]",
            border_style=score_color,
            title="Efficiency",
        ),
        Panel(
            f"[bold magenta]{fmt(today_res['ultra_min'])}[/bold magenta]\n[dim]True Active Duration[/dim]",
            border_style="magenta",
            title="Ultra-Realistic",
        ),
    )
    console.print(stats_row)

    # 3) Dynamic cycle tracker (16th -> 15th), cycle-to-date
    cycle_to_date = summarize_period(daily, cycle_start, today)

    cycle_system_hours = cycle_to_date["system_td"].total_seconds() / 3600
    cycle_focused_hours = cycle_to_date["active_td"].total_seconds() / 3600
    cycle_focus_ratio = (cycle_focused_hours / cycle_system_hours * 100) if cycle_system_hours > 0 else 0

    cycle_pct = min((cycle_system_hours / CYCLE_TARGET_HOURS) * 100, 100)
    cycle_progress = ProgressBar(
        total=CYCLE_TARGET_HOURS,
        completed=cycle_system_hours,
        width=None,
        finished_style="bright_green",
    )

    cycle_label = f"{cycle_start.strftime('%d %b')} - {cycle_end.strftime('%d %b')}"
    console.print(
        Panel(
            Columns(
                [
                    f"[bold bright_blue]System MTD: {cycle_system_hours:.1f}h / {CYCLE_TARGET_HOURS}h ({cycle_pct:.1f}%)[/bold bright_blue]",
                    cycle_progress,
                    (
                        f"[bold green]Focused: {cycle_focused_hours:.1f}h[/bold green]\n"
                        f"[dim]Focus Ratio: {cycle_focus_ratio:.1f}% | Avg Intensity: {cycle_to_date['avg_intensity']:.1f}%[/dim]"
                    ),
                    f"[dim]Remaining: {max(0, CYCLE_TARGET_HOURS - cycle_system_hours):.1f}h[/dim]",
                ],
                expand=True,
                align="left",
            ),
            title=f"Cycle Target Tracker ({cycle_label})",
            border_style="bright_blue",
        )
    )

    # 4) Period summaries (from today backward)
    last_7_start = today - datetime.timedelta(days=6)
    last_30_start = today - datetime.timedelta(days=29)

    sum7 = summarize_period(daily, last_7_start, today)
    sum30 = summarize_period(daily, last_30_start, today)
    sum_cycle = cycle_to_date

    summary_table = Table(title="Rolling Summaries", box=box.ROUNDED, expand=True)
    summary_table.add_column("Period", style="cyan")
    summary_table.add_column("Range", style="dim")
    summary_table.add_column("System On", justify="right")
    summary_table.add_column("Focused", justify="right", style="bright_green")
    summary_table.add_column("Quantized", justify="right", style="bright_blue")
    summary_table.add_column("Avg Intensity", justify="right")

    def add_summary_row(name, start_d, end_d, summary):
        s_color = (
            "bright_green" if summary["avg_intensity"] > 70
            else "yellow" if summary["avg_intensity"] > 40
            else "dim"
        )
        summary_table.add_row(
            name,
            f"{start_d.strftime('%d %b')} - {end_d.strftime('%d %b')}",
            fmt(summary["system_td"]),
            fmt(summary["active_td"]),
            fmt(summary["quantized_min"]),
            f"[{s_color}]{summary['avg_intensity']:.1f}%[/{s_color}]",
        )

    add_summary_row("Last 7 Days", last_7_start, today, sum7)
    add_summary_row("Last 30 Days", last_30_start, today, sum30)
    add_summary_row("Current Cycle", cycle_start, today, sum_cycle)
    console.print(summary_table)

    # 5) Apps
    if today_res["apps"]:
        app_table = Table(box=box.SIMPLE, expand=True)
        app_table.add_column("Application Name", style="cyan")
        app_table.add_column("Minutes (Approx)", justify="right", style="green")
        for app, count in today_res["apps"]:
            app_table.add_row(app, f"{count}m")
        console.print(Panel(app_table, title="Application Focus Distribution", border_style="dim"))

    # 6) Hourly map
    heatmap = Text("Intensity: ", style="bold")
    for h in range(24):
        intensity = today_res["hourly"][h]
        char = "█" if intensity > 45 else "▆" if intensity > 30 else "▃" if intensity > 10 else "▂" if intensity > 0 else " "
        color = "bright_green" if intensity > 30 else "yellow" if intensity > 10 else "dim"
        heatmap.append(f" {h:02d}", style="dim")
        heatmap.append(char, style=color)
    console.print(Panel(heatmap, title="Hourly Work Timeline", border_style="dim"))

    # 7) Session breakdown (today)
    if today_res["slots"]:
        session_table = Table(box=box.SIMPLE, expand=True)
        session_table.add_column("#", style="dim")
        session_table.add_column("Session Range", style="cyan")
        session_table.add_column("Duration", justify="right")
        for i, (s, e) in enumerate(today_res["slots"], 1):
            session_table.add_row(str(i), f"{s.strftime('%H:%M:%S')} - {e.strftime('%H:%M:%S')}", fmt(e - s))
        console.print(Panel(session_table, title="Daily System Presence (Sessions)", border_style="dim"))

    # 8) Discrete breakdown (today)
    if today_res["discrete_slots"]:
        slot_table = Table(box=box.SIMPLE, expand=True)
        slot_table.add_column("Time Slot", style="cyan")
        slot_table.add_column("Intensity Score", justify="right")
        slot_table.add_column("Intensity Bar", ratio=1)
        
        # Show last 12 slots
        for s in today_res["discrete_slots"][-12:]:
            score = s["intensity"]
            color = "bright_green" if score > 70 else "yellow" if score > 30 else "red"
            bar_width = int(score / 5)
            slot_table.add_row(
                f"{s['start'].strftime('%H:%M')} - {s['end'].strftime('%H:%M')}",
                f"[{color}]{score:.1f}%[/{color}]",
                f"[{color}]{'█' * bar_width}[/{color}]"
            )

        console.print(Panel(slot_table, title="Discrete High-Fidelity Slot Analysis (Recent)", border_style="dim"))

    # 9) Historical table (Last 30 days) with day bounds
    table = Table(title="Historical Accountability (Last 30 Days)", box=box.ROUNDED, expand=True)
    table.add_column("Date", style="cyan")
    table.add_column("Start", justify="right", style="dim")
    table.add_column("End", justify="right", style="dim")
    table.add_column("Sessions", style="magenta", justify="left", min_width=24, overflow="fold")
    table.add_column("System On", justify="right")
    table.add_column("Focused", justify="right", style="bright_green")
    table.add_column("Quantized", justify="right", style="bright_blue")
    table.add_column("Intensity", justify="right")

    for i in range(29, -1, -1):
        day = today - datetime.timedelta(days=i)
        if day not in daily:
            continue

        res = daily[day]
        start_str, end_str = get_day_bounds(res["slots"])
        sessions_count = len(res["slots"])
        session_ranges = ", ".join(
            f"{s.strftime('%H:%M')}-{e.strftime('%H:%M')}"
            for s, e in res["slots"]
        )
        sessions_cell = f"{sessions_count} ({session_ranges})" if session_ranges else str(sessions_count)

        s_color = (
            "bright_green" if res["avg_intensity"] > 70
            else "yellow" if res["avg_intensity"] > 40
            else "dim"
        )

        table.add_row(
            day.strftime("%a %d %b"),
            start_str,
            end_str,
            sessions_cell,
            fmt(res["system_td"]),
            fmt(res["active_td"]),
            fmt(res["discrete_min"]),
            f"[{s_color}]{res['avg_intensity']:.1f}%[/{s_color}]",
        )

    console.print(table)
    console.print(
        f"[dim]Hybrid tracking: Windows Power Events (Presence) + {SLOT_MINUTES}m Discrete Buckets (Accountability).[/dim]",
        justify="right",
    )


if __name__ == "__main__":
    main()
