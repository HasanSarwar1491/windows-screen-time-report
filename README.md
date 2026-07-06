# Windows Accountability Tracker

Windows screen-time and accountability reporting with:
- **System presence tracking** (wake/sleep based)
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
1. Install dependencies (`pywin32`, `rich`, `psutil`)
2. Create Startup launcher for the background logger
3. Start logger silently
4. Generate initial report

To open report anytime:

```powershell
python screen_time.py
```

---

## What’s New (Current Version)

- **5-minute discrete accountability buckets** (updated from 10-minute)
- **Cycle-based target window:** `16th -> 15th`
- **Cycle tracker now uses System On hours** for progress (not quantized-only)
- **Cycle quality indicators:** Focused hours, Focus Ratio, Avg Intensity
- **Rolling summaries:** Last 7 days, Last 30 days, Current cycle
- **Historical Accountability includes:**
  - Start
  - End
  - Sessions (count + full session ranges)
  - System On
  - Focused
  - Quantized
  - Intensity
- **Today session breakdown** includes a **TOTAL** row

---

## Metric Definitions

### 1) System On
Total time Windows was awake/on (presence time).

### 2) Focused
Minutes with active input while a foreground app/window was tracked.

### 3) Quantized
Accountability minutes aggregated from **5-minute discrete buckets**.

### 4) Intensity
Average activity intensity from discrete bucket sampling.

### 5) Focus Ratio
`Focused / System On * 100` for the selected period.

---

## Reporting Sections

### Key Stats (Today)
- Work Time (Quantized)
- Efficiency (Mean Intensity Score)
- Ultra-Realistic (True Active Duration)

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
Per day:
- Date, Start, End
- Sessions (`N (HH:MM-HH:MM, ...)`)
- System On, Focused, Quantized
- Intensity

---

## Files

- `activity_logger.py` - background logger (active/idle + foreground app)
- `screen_time.py` - dashboard/report generator
- `setup.bat` - one-click setup/start flow
- `install_task.py` - optional scheduled-task installer

---

## Requirements

- Windows 10/11
- Python 3.8+
- Packages: `pywin32`, `rich`, `psutil`

---

## Notes

- Logs are stored in `~/.screen_time/activity.log`
- Logger lock file prevents duplicate background logger instances
- If scheduled-task creation is denied, use Startup-based setup (`setup.bat`) instead
