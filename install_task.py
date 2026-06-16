import argparse
import os
import subprocess
import sys


TASK_NAME = "ScreenTimeActivityLogger"


def _pythonw_path():
    exe = sys.executable
    if exe.lower().endswith("pythonw.exe"):
        return exe
    candidate = os.path.join(os.path.dirname(exe), "pythonw.exe")
    if os.path.exists(candidate):
        return candidate
    return exe


def _logger_script_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "activity_logger.py")


def install_task():
    logger_path = _logger_script_path()
    if not os.path.exists(logger_path):
        raise FileNotFoundError(f"Logger not found: {logger_path}")

    pyw = _pythonw_path()
    task_cmd = f'"{pyw}" "{logger_path}"'

    cmd = [
        "schtasks",
        "/Create",
        "/TN",
        TASK_NAME,
        "/SC",
        "ONLOGON",
        "/TR",
        task_cmd,
        "/F",
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "Failed to create task.")
    print(f"Installed task: {TASK_NAME}")


def uninstall_task():
    cmd = ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        err = result.stderr.strip() or result.stdout.strip() or "Failed to delete task."
        if "cannot find the file specified" in err.lower():
            print(f"Task not found: {TASK_NAME}")
            return
        raise RuntimeError(err)
    print(f"Removed task: {TASK_NAME}")


def main():
    parser = argparse.ArgumentParser(description="Install/uninstall Screen Time activity logger task.")
    parser.add_argument("--uninstall", action="store_true", help="Remove scheduled task.")
    args = parser.parse_args()

    if args.uninstall:
        uninstall_task()
    else:
        install_task()


if __name__ == "__main__":
    main()
