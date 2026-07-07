import ctypes
from ctypes import wintypes
import datetime
import os
import signal
import time
import sys
import msvcrt
import psutil
from pynput import mouse, keyboard
import threading

# Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

LOG_DIR = os.path.expanduser("~/.screen_time")
LOG_FILE = os.path.join(LOG_DIR, "activity.log")
LOCK_FILE = os.path.join(LOG_DIR, "activity_logger.lock")

INPUT_IDLE_MINUTES = 5
POLL_SECONDS = 30
RETENTION_DAYS = 400

# Counters for Teramind-like activity tracking
# We track how many seconds in the POLL_SECONDS window had at least one event
activity_lock = threading.Lock()
active_seconds_set = set() # Stores the second offsets within the current window that had activity

def on_input():
    global active_seconds_set
    with activity_lock:
        # Map current time to a discrete second to avoid double counting multiple events in same second
        active_seconds_set.add(int(time.time()))

def on_click(x, y, button, pressed):
    if pressed:
        on_input()

def on_press(key):
    on_input()

def on_scroll(x, y, dx, dy):
    on_input()

def on_move(x, y):
    # Optional: Significant movement only?
    # For now, any movement counts as activity per Teramind logic
    on_input()

class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

def _ensure_log_dir():
    os.makedirs(LOG_DIR, exist_ok=True)

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

def _append_state(state, app, title, active_seconds):
    now = datetime.datetime.now().replace(microsecond=0)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            # Format: timestamp,state,app,title,active_seconds_in_period
            f.write(f"{now.isoformat()},{state},{app},{title},{active_seconds}\n")
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
    global active_seconds_set
    _ensure_log_dir()
    lock_file_handle = open(LOCK_FILE, 'w')
    try: msvcrt.locking(lock_file_handle.fileno(), msvcrt.LK_NBLCK, 1)
    except: sys.exit(0)

    # Start listeners
    mouse_listener = mouse.Listener(on_click=on_click, on_scroll=on_scroll, on_move=on_move)
    key_listener = keyboard.Listener(on_press=on_press)
    mouse_listener.start()
    key_listener.start()

    should_run = {"value": True}
    def _stop_handler(_s, _f): should_run["value"] = False
    signal.signal(signal.SIGINT, _stop_handler)
    
    _rotate_log()
    
    try:
        while should_run["value"]:
            # Wait for the next poll interval
            now = time.time()
            sleep_time = POLL_SECONDS - (now % POLL_SECONDS)
            time.sleep(sleep_time)
            
            # Capture current window and calculate activity for the period
            app, title = get_foreground_info()
            
            with activity_lock:
                # Count how many unique seconds had input
                active_count = len(active_seconds_set)
                active_seconds_set.clear()

            # State is "active" if there was at least 1 second of input
            state = "active" if active_count > 0 else "idle"
            
            _append_state(state, app, title, active_count)
    finally:
        mouse_listener.stop()
        key_listener.stop()
        lock_file_handle.close()
        try: os.remove(LOCK_FILE)
        except: pass

if __name__ == "__main__":
    main()
