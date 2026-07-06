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
POLL_SECONDS = 60
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

def get_foreground_app_name():
    """Improved app name detection including fallback for elevated windows."""
    try:
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return "Idle/Desktop"
            
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        
        # Method 1: psutil (fastest, covers most user apps)
        try:
            process = psutil.Process(pid.value)
            return process.name()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
            
        # Method 2: Internal Windows Window Title (fallback for restricted PIDs)
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buff = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buff, length + 1)
            # Try to extract app name from title (usually after last ' - ')
            title = buff.value
            if " - " in title:
                return title.split(" - ")[-1]
            return title[:30] # Limit length
            
        return "System/Elevated"
    except Exception:
        return "Unknown"

def _get_state():
    seconds = _seconds_since_last_input()
    if seconds is None:
        return "active", "Unknown"
    state = "idle" if seconds >= INPUT_IDLE_MINUTES * 60 else "active"
    app = get_foreground_app_name() if state == "active" else "None"
    return state, app

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
                parts = line.split(",", 2)
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

def _append_state(state, app):
    now = datetime.datetime.now().replace(second=0, microsecond=0)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{now.isoformat()},{state},{app}\n")
    except Exception:
        pass

def _sleep_to_next_tick():
    now = time.time()
    remainder = int(now) % POLL_SECONDS
    sleep_for = POLL_SECONDS - remainder
    if sleep_for <= 0:
        sleep_for = POLL_SECONDS
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
        state, app = _get_state()
        _append_state(state, app)
        _sleep_to_next_tick()
    
    lock_file_handle.close()
    try:
        os.remove(LOCK_FILE)
    except:
        pass

if __name__ == "__main__":
    main()
