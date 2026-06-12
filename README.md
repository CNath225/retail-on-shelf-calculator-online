# Retail On-shelf Rate Calculator Online by CodeNATHAN

Streamlit online version of the on-shelf rate calculator.

## Local Run

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

1. Put this folder in a GitHub repository.
2. In Streamlit Community Cloud, create a new app from that repository.
3. Set the main file path to `app.py`.
4. The app will install packages from `requirements.txt`.

Chinese deployment notes are in `DEPLOYMENT_CN.md`.

## Inputs

- Repsly raw export workbook with a `Submissions` sheet.
- Range workbook with `TTL Store#`, `Range#`, and `Range%` columns.
- Report template workbook, usually the same workbook as the range file.

Uploaded files and generated reports are temporary runtime files.
