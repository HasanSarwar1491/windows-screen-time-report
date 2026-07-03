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

## 📊 Understanding the Metrics

### 1. System On (Total Time)
Total duration your computer was "awake" or the screen was active. (Includes Windows maintenance wakes).

### 2. Focused Active Time
Minutes where **keyboard/mouse input** was detected AND a **specific application was in focus**. 

### 3. Accountability Score (Efficiency)
Calculated as: `(Focused Active Time / System On Time) * 100`.
*   **75%+ (Green):** High intensity.
*   **40%-74% (Yellow):** Moderate activity / frequent breaks.
*   **<40% (Red):** Low activity / PC left idle.

### 4. Work Intensity Map (Heatmap)
A 24-hour visual timeline. Dark blocks (█) mean constant activity.

### 5. Top Focused Applications
Lists apps that were "on top" and actively used. Useful for proving work focus.

### 6. Monthly Goal (176h Target)
Specifically for monthly hour requirements. Shows **Month-to-Date** totals and **Remaining** hours left in the month.

---

## 📅 Best Practices for 176h/Month
*   **Hibernate over Sleep:** Use **Hibernate** at the end of the day. It stops all timers instantly and prevents "ghost" wakes at night.
*   **The 5-Minute Rule:** If you are idle for 5 minutes, the active clock stops. Stay active in your focused window to keep your score high.
*   **Retention:** Logs are kept for **400 days** by default to ensure you can track your history across years.

---

## 📝 Requirements
- Windows 10/11 | Python 3.8+ | Admin rights (for pip)
