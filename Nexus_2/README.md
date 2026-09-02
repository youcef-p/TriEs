# Nexus

Nexus is a Flask-based workspace for file listing and comparison, PDF text
extraction, and spreadsheet reception-data processing. It also includes a
separate Flask service for generating monthly DOCX reports.

## Components

- **Nexus Hub** — `app.py`, served on port 5000 locally.
- **Monthly Reports Service** — `monthly_reports_service/app.py`, served on
  port 5001 locally.

## Run in Replit

The Replit Run button starts the Nexus Hub from the `Nexus` directory and
binds it to Replit's assigned `PORT`. This is the primary browser-facing app.

To run the monthly report service separately while developing:

```bash
cd Nexus/monthly_reports_service
python3 -m flask --app app run --host 0.0.0.0 --port 5001
```

The original Windows launchers are kept unchanged for local Windows use:
`run.bat` and `monthly_reports_service/run_dev.bat`.

## Dependencies

The shared Replit environment uses the root `pyproject.toml` and `uv.lock`.
The complete imported-project dependency list is also documented in
`Nexus/requirements.txt`.