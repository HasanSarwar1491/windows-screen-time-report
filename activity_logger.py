import ctypes
from ctypes import wintypes
import datetime
import os
import signal
import time
import sys
import msvcrt
import psutil

# Windows API for getting foreground window
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

INPUT_IDLE_MINUTES = 5
POLL_SECONDS = 30
RETENTION_DAYS = 400
LOG_DIR = os.path.expanduser("~/.screen_time")
LOG_FILE = os.path.join(LOG_DIR, "activity.log")
LOCK_FILE = os.path.join(LOG_DIR, "activity_logger.lock")

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
    if elapsed_ms < 0:
        elapsed_ms = 0
    return elapsed_ms / 1000.0

def get_foreground_info():
    """Returns (app_name, window_title)"""
    try:
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return "Idle/Desktop", "Desktop"
            
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        
        app_name = "Unknown"
        window_title = "Unknown"

        # Get Process Name
        try:
            process = psutil.Process(pid.value)
            app_name = process.name()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
            
        # Get Window Title
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            window_title = buff.value.replace(",", ";") # Avoid CSV issues
            
        if app_name == "Unknown" and window_title != "Unknown":
            app_name = "System/Elevated"
            
        return app_name, window_title
    except Exception:
        return "Unknown", "Unknown"

def _get_state():
    seconds = _seconds_since_last_input()
    if seconds is None:
        return "active", "Unknown", "Unknown"
    state = "idle" if seconds >= INPUT_IDLE_MINUTES * 60 else "active"
    if state == "active":
        app, title = get_foreground_info()
    else:
        app, title = "None", "None"
    return state, app, title

def _rotate_log():
    if not os.path.exists(LOG_FILE):
        return
    now = datetime.datetime.now()
    cutoff = now - datetime.timedelta(days=RETENTION_DAYS)
    kept = []
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line: continue
                parts = line.split(",", 1)
                if len(parts) < 2: continue
                try:
                    ts = datetime.datetime.fromisoformat(parts[0])
                except ValueError: continue
                if ts >= cutoff:
                    kept.append(line)
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            for line in kept:
                f.write(line + "\n")
    except Exception:
        pass

def _append_state(state, app, title):
    now = datetime.datetime.now().replace(microsecond=0)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{now.isoformat()},{state},{app},{title}\n")
    except Exception:
        pass

def _sleep_to_next_tick():
    now = time.time()
    remainder = now % POLL_SECONDS
    sleep_for = POLL_SECONDS - remainder
    time.sleep(sleep_for)

def main():
    _ensure_log_dir()
    
    lock_file_handle = open(LOCK_FILE, 'w')
    try:
        msvcrt.locking(lock_file_handle.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        sys.exit(0)

    should_run = {"value": True}

    def _stop_handler(_signum, _frame):
        should_run["value"] = False

    signal.signal(signal.SIGINT, _stop_handler)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, _stop_handler)

    _rotate_log()

    while should_run["value"]:
        state, app, title = _get_state()
        _append_state(state, app, title)
        _sleep_to_next_tick()
    
    lock_file_handle.close()
    try:
        os.remove(LOCK_FILE)
    except:
        pass

if __name__ == "__main__":
    main()
