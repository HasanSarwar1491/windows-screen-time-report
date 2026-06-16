# screen-time-report

Windows screen-time reporting using System Event Log power events.

The report includes:
- Daily screen-on totals for the last 30 days
- Session segments per day
- 7-day, 30-day, and cycle summaries against an 8-hour weekday target
- A detailed daily working-hours table for today

## Files

- `screen_time.py` - main report script
- `activity_logger.py` - optional local input-state logger utility
- `install_task.py` - optional scheduler helper for `activity_logger.py`

## Requirements

- Windows 10/11
- Python 3.8+
- `pywin32`

Install dependency:

```bash
pip install pywin32
```

## Run

```bash
python screen_time.py
```

## Optional utilities

If you want to run the activity logger in the background at logon:

```bash
python install_task.py
```

To remove that scheduled task:

```bash
python install_task.py --uninstall
```

## Notes

- Core reporting is based on Event Log power events.
- The script applies event cleanup rules to avoid counting transient wake artifacts.
- Billing cycle is 16th to 15th.
