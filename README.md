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

## 📝 Requirements

- Windows 10/11
- Python 3.8+
- Admin rights (for initial dependency installation via pip)
