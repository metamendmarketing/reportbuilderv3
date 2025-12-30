# Monthly SEO Email Builder (Streamlit)

Generates a client-ready **monthly SEO update email** as:
- **HTML** (for copy/paste)
- **.eml** (recommended for Outlook; preserves inline screenshots)

## Run
```bash
pip install streamlit openai
streamlit run monthly_report_builder_app.py
```

## Secrets
Put your key in Streamlit secrets or env var:
- `.streamlit/secrets.toml` (do not commit):
```toml
OPENAI_API_KEY="sk-..."
```
