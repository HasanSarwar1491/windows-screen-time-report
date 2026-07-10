# Windows Accountability Tracker

Windows screen-time and high-fidelity accountability reporting with:
- **System presence tracking** (wake/sleep based)
- **High-Fidelity Input Density Tracking** (keyboard, mouse move, click, and scroll)
- **Foreground activity logging** (focused active minutes)
- **Rich terminal dashboard**
- **Cycle-based target tracking**

---

## Quick Setup

Run:

```bat
setup.bat
```

This will:
1. Install dependencies (`pywin32`, `rich`, `psutil`, `pynput`)
2. Create Startup launcher for the background logger
3. Start logger silently
4. Generate initial report

To open report anytime:

```powershell
python screen_time.py
```

---

## What’s New (Current Version)

- **Input Density Tracking:** Replaces simple presence with per-second input monitoring (keyboard + mouse).
- **High-Fidelity Intensity:** Intensity is now calculated as the actual percentage of seconds active within a slot.
- **5-minute discrete accountability buckets.**
- **Cycle-based target window:** `16th -> 15th`.
- **Cycle tracker now uses System On hours** for progress.
- **Cycle quality indicators:** Focused hours, Focus Ratio, Avg Intensity.
- **Rolling summaries:** Last 7 days, Last 30 days, Current cycle.
- **Historical Accountability includes:**
  - Start / End times
  - Sessions (count + full session ranges)
  - System On / Focused / Quantized durations
  - Intensity Score
- **Today session breakdown** includes a **TOTAL** row.

---

## Metric Definitions

### 1) System On
Total time Windows was awake/on (presence time).

### 2) Focused
Total duration of time where the computer was on and not idle.

### 3) Enterprise Quantized
Accountability minutes aggregated from **5-minute discrete buckets**. A bucket is counted if it meets the minimum activity threshold (5 seconds of input).

### 4) Mean Intensity Score
The "Work Density" of your active 5-minute slots. It represents what percentage of the time you were actually interacting with your computer (typing/clicking) versus passive viewing.

### 5) Ultra-Realistic (True Active Duration)
The total time the foreground window was tracked while you were present at the desk.

---

## Reporting Sections

### Key Stats (Today)
- **Work Time (Quantized):** "Billable" style work time based on active buckets.
- **Efficiency (Mean Intensity Score):** Your overall activity density for the day.
- **Ultra-Realistic:** Every minute you were actively using the machine.

### Cycle Target Tracker
- Current cycle range (16th -> 15th)
- System MTD vs 176h target
- Remaining hours
- Focused hours + Focus Ratio + Avg Intensity

### Rolling Summaries
- Last 7 Days
- Last 30 Days
- Current Cycle

### Historical Accountability (Last 30 Days)
Per day analysis of sessions, system time, focused time, and intensity scores.

---

## Files

- `activity_logger.py` - Background logger with global input hooks.
- `screen_time.py` - Dashboard/report generator with high-fidelity logic.
- `setup.bat` - One-click setup/start flow.

---

## Requirements

- Windows 10/11
- Python 3.8+
- Packages: `pywin32`, `rich`, `psutil`, `pynput`

---

## Notes

- Logs are stored in `~/.screen_time/activity.log`
- Logger uses a lock file to prevent duplicate background instances.
- Data is stored in a 5-column format: `Timestamp, State, App, WindowTitle, ActiveSeconds`.
