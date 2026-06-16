import ctypes
import datetime
import os
import signal
import time


INPUT_IDLE_MINUTES = 5
POLL_SECONDS = 60
RETENTION_DAYS = 60
LOG_DIR = os.path.expanduser("~/.screen_time")
LOG_FILE = os.path.join(LOG_DIR, "activity.log")


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32


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


def _get_state():
    seconds = _seconds_since_last_input()
    if seconds is None:
        return "active"
    return "idle" if seconds >= INPUT_IDLE_MINUTES * 60 else "active"


def _rotate_log():
    if not os.path.exists(LOG_FILE):
        return
    cutoff = datetime.datetime.now() - datetime.timedelta(days=RETENTION_DAYS)
    kept = []
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split(",", 1)
            if len(parts) != 2:
                continue
            try:
                ts = datetime.datetime.fromisoformat(parts[0])
            except ValueError:
                continue
            if ts >= cutoff:
                kept.append(line)
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        for line in kept:
            f.write(line + "\n")


def _append_state(state):
    now = datetime.datetime.now().replace(second=0, microsecond=0)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{now.isoformat()},{state}\n")


def _sleep_to_next_tick():
    now = time.time()
    remainder = int(now) % POLL_SECONDS
    sleep_for = POLL_SECONDS - remainder
    if sleep_for <= 0:
        sleep_for = POLL_SECONDS
    time.sleep(sleep_for)


def main():
    should_run = {"value": True}

    def _stop_handler(_signum, _frame):
        should_run["value"] = False

    signal.signal(signal.SIGINT, _stop_handler)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, _stop_handler)

    _ensure_log_dir()
    _rotate_log()

    while should_run["value"]:
        _append_state(_get_state())
        _sleep_to_next_tick()


if __name__ == "__main__":
    main()
