# screen-time-report

Windows screen-time tracker that calculates daily active hours from Event Log power and display events.

## Requirements

- Windows 10/11
- Python 3.8+
- `pywin32` package

```bash
pip install pywin32
```

## Usage

```bash
python screen_time.py
```

For lock/unlock idle detection (optional, improves accuracy):

```bash
# Run PowerShell as Administrator
python screen_time.py
```

## How It Works

The script reads the Windows **System** event log for power and display state changes:

| Event | Source | ID | Meaning |
|-------|--------|----|---------|
| Display On | Kernel-Power | 507 | User returned / screen woke |
| Display Off | Kernel-Power | 506 | Idle timeout / screen sleep |
| Sleep | Kernel-Power | 42 | System entering sleep/hibernate |
| Shutdown | Kernel-Power | 109 | System shutting down |
| Boot | Kernel-General | 12 | OS started |
| Shutdown | Kernel-General | 13 | OS stopped |
| Wake | Power-Troubleshooter | 1 | Resumed from sleep |

If run as Admin, it also reads **Security** log lock/unlock events (4800/4801) for finer idle detection.

A strict state machine pairs START→STOP events into sessions. Gaps under 5 minutes are merged (configurable via `IDLE_GAP_MINUTES`). Sessions spanning midnight are split across days.

## Output

- **Daily breakdown** — active hours, percentage vs 8h target, first-on/last-off times, and individual session segments
- **Summaries** — Today, Last 7 Days, Last 30 Days
- **Billing cycles** — Current and previous 16th-to-15th monthly cycles with hours and percentages
- Days with zero activity are excluded from the report and percentage calculations

## Configuration

Constants at the top of `screen_time.py`:

| Variable | Default | Description |
|----------|---------|-------------|
| `REQUIRED_HOURS` | `8` | Target hours per weekday |
| `IDLE_GAP_MINUTES` | `5` | Gaps shorter than this are merged as active time |
