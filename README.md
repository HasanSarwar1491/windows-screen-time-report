# Windows Accountability Tracker

Windows screen-time tracking and Teramind-style accountability monitoring.

## 🚀 Quick Setup

If you are a new user, simply run the setup script:

```powershell
./setup.bat
```

This will:
1. Install all dependencies (`pywin32`, `rich`, `psutil`).
2. Configure the **Background Activity Logger** to start automatically with Windows.
3. Launch the logger silently.
4. Generate your first report.

---

## 📊 How to Use

### 1. View Dashboard
To see your productivity dashboard anytime, run:
```bash
python screen_time.py
```

### 2. Best Practices for Accountability
* **Hibernate over Sleep:** When finishing work, use **Hibernate**. This ensures the "System On" clock stops exactly when you do.
* **Focused Window Rule:** The tracker logs the application you are actually interacting with. Use this to prove your focus even if background apps (like YouTube or Spotify) are running.

## 🛠️ Components

- `screen_time.py` - Main dashboard (Accountability scores, Heatmap, Session lists).
- `activity_logger.py` - Background agent that monitors active windows and input.
- `setup.bat` - One-click installer and configuration tool.

---

## 📊 Understanding the Metrics

To help you maintain accountability, the dashboard provides several key metrics inspired by enterprise monitoring tools:

### 1. System On (Total Time)
This is the total duration your computer was "awake" or the screen was active. It is derived from Windows System Event Logs. 
*   **Note:** This may include "ghost wakes" (Windows Update, maintenance) where the PC is on but you aren't using it.

### 2. Focused Active Time
This is the most critical metric. It counts the minutes where the background agent detected **actual keyboard or mouse input** AND a **specific application was in focus**. 
*   If you are away for 5+ minutes, the clock stops, even if the screen is on.

### 3. Accountability Score (Efficiency)
Calculated as: `(Focused Active Time / System On Time) * 100`.
*   **75% - 100% (Green):** High intensity. You were actively working for most of the time the PC was on.
*   **40% - 74% (Yellow):** Moderate activity. Indicates frequent breaks or background tasks.
*   **0% - 39% (Red):** Low activity. Likely indicates the PC was left on while you were away, or "ghost" system wakes occurred.

### 4. Work Intensity Map (Heatmap)
A 24-hour visual timeline of your day.
*   **Dark Blocks (█):** Constant activity (45+ minutes of input in that hour).
*   **Light Blocks (░):** Low/Spasmodic activity.
*   **Empty:** System was asleep or no activity was detected.

### 5. Top Focused Applications
Lists the applications that were "on top" and actively clicked during your productive hours. Use this to verify that your time was spent in work-related tools (IDE, Browser, Slack) rather than background noise.

---

## 📝 Requirements

- Windows 10/11
- Python 3.8+
- Admin rights (for initial dependency installation via pip)
