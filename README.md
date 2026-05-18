# Finish Status Report Automation

Phase-1 automation for DepotNet:
- Login
- Open `Reporting` -> `Finish Status Report`
- Disable filter if enabled
- Click Export (Excel)

## Project Structure

- `src/main.py` - Entry point
- `src/core/job_manger.py` - Job orchestration
- `src/portal/login_page.py` - Login flow
- `src/portal/finish_status_report_page.py` - Report navigation and export
- `src/utils/logger.py` - Logging setup
- `src/utils/retry_helper.py` - Retry decorator with exponential backoff
- `logs/` - Runtime logs

## Setup

```bash
pip install -r requirements.txt
playwright install
```

## Environment Variables

Set in `.env`:

- `DEPOTNET_URL`
- `DEPOTNET_USERNAME`
- `DEPOTNET_PASSWORD`

## Run

```bash
python -m src.main
```
