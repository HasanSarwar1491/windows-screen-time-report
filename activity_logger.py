import ctypes
from ctypes import wintypes
import datetime
import os
import signal
import time
import sys
import msvcrt
import psutil

# Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

LOG_DIR = os.path.expanduser("~/.screen_time")
LOG_FILE = os.path.join(LOG_DIR, "activity.log")
LOCK_FILE = os.path.join(LOG_DIR, "activity_logger.lock")

INPUT_IDLE_MINUTES = 5
POLL_SECONDS = 30
RETENTION_DAYS = 400

class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

def _ensure_log_dir():
    os.makedirs(LOG_DIR, exist_ok=True)

def _current_tick_ms():
    if hasattr(kernel32, "GetTickCount64"):
        return int(kernel32.GetTickCount64())
    return int(kernel32.GetTickCount())

def _seconds_since_last_input():
    info = LASTINPUTINFO()
    info.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if not user32.GetLastInputInfo(ctypes.byref(info)):
        return None
    elapsed_ms = _current_tick_ms() - int(info.dwTime)
    if elapsed_ms < 0: elapsed_ms = 0
    return elapsed_ms / 1000.0

def get_foreground_info():
    try:
        hwnd = user32.GetForegroundWindow()
        if not hwnd: return "Idle/Desktop", "Desktop"
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        app_name, window_title = "Unknown", "Unknown"
        try:
            process = psutil.Process(pid.value)
            app_name = process.name()
        except: pass
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            window_title = buff.value.replace(",", ";")
        if app_name == "Unknown" and window_title != "Unknown": app_name = "System/Elevated"
        return app_name, window_title
    except:
        return "Unknown", "Unknown"

def _get_state():
    app, title = get_foreground_info()
    seconds = _seconds_since_last_input()
    
    if seconds is None: return "active", app, title
    state = "idle" if seconds >= INPUT_IDLE_MINUTES * 60 else "active"
    if state == "idle": app, title = "None", "None"
    return state, app, title

def _append_state(state, app, title):
    now = datetime.datetime.now().replace(microsecond=0)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{now.isoformat()},{state},{app},{title}\n")
    except: pass

def _rotate_log():
    if not os.path.exists(LOG_FILE): return
    cutoff = datetime.datetime.now() - datetime.timedelta(days=RETENTION_DAYS)
    kept = []
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.split(",", 1)
                if len(parts) < 2: continue
                try:
                    if datetime.datetime.fromisoformat(parts[0]) >= cutoff: kept.append(line.strip())
                except: continue
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            for line in kept: f.write(line + "\n")
    except: pass

def main():
    _ensure_log_dir()
    lock_file_handle = open(LOCK_FILE, 'w')
    try: msvcrt.locking(lock_file_handle.fileno(), msvcrt.LK_NBLCK, 1)
    except: sys.exit(0)

    should_run = {"value": True}
    def _stop_handler(_s, _f): should_run["value"] = False
    signal.signal(signal.SIGINT, _stop_handler)
    
    _rotate_log()
    
    while should_run["value"]:
        state, app, title = _get_state()
        _append_state(state, app, title)
        now = time.time()
        time.sleep(POLL_SECONDS - (now % POLL_SECONDS))
    
    lock_file_handle.close()
    try: os.remove(LOCK_FILE)
    except: pass

if __name__ == "__main__":
    main()
