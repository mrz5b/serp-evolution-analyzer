# SERP Evolution Analyzer

Five Blocks internal tool for analyzing how search results changed before and after a specific event.

Upload an IMPACT CSV, set a pivot date, and get back:
- **Branded PPTX deck** — 6 slides: title, executive summary, stats cards, persistent URL butterfly chart, new URL bar chart, source type shift chart
- **URL Analysis Excel** — URL-level pre/post day counts with New/Persistent/Dropped classification
- **Source Type Analysis Excel** — Slot-day aggregation by source category with domain-level detail

## Authentication

Production: Google OAuth restricted to `@fiveblocks.com` accounts. Users see a "Sign in with Google" button — only Five Blocks email addresses get through.

Local dev: Auth gate is automatically skipped when Google OAuth secrets aren't configured.

## Configuration

### Anthropic API Key (optional)
Enter in the sidebar for:
- AI-powered domain categorization (unknown domains classified by Claude)
- AI-generated executive summary (4 concise data-driven points)

Without an API key, the app runs in **fallback mode**:
- Domain categorization uses the built-in mapping (unmapped domains → "Other")
- Executive summary uses data-driven templates

### Owned Domains
Enter comma-separated domains that belong to the client (e.g., `brookfield.com, privatewealth.brookfield.com`). These are automatically categorized as "Corporate / Owned."

## Deployment (Streamlit Cloud)

1. Push to GitHub
2. Connect repo to [Streamlit Cloud](https://streamlit.io/cloud)
3. Add secrets in the Streamlit Cloud dashboard:

**Required for authentication:**
- `GOOGLE_CLIENT_ID` — from Google Cloud Console OAuth 2.0 credentials
- `GOOGLE_CLIENT_SECRET` — from the same credential
- `REDIRECT_URI` — your Streamlit Cloud app URL (e.g., `https://serp-evolution.streamlit.app/`)
- `COOKIE_KEY` — any random string for cookie encryption

**Optional:**
- `ANTHROPIC_API_KEY` — enables AI-powered exec summaries + domain classification

4. In Google Cloud Console:
   - APIs & Services → Credentials → Create OAuth 2.0 Client ID
   - Application type: Web application
   - Authorized redirect URI: your Streamlit Cloud URL
   - You can reuse the existing fb-search-evolution-case-study credentials if the redirect URI is added

5. Deploy

### Local Development (no auth)

When running locally without Google OAuth secrets configured, the auth gate is skipped entirely — the app runs without login. This is the expected dev workflow.

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Pipeline

1. **Data processing** — Parse IMPACT CSV, filter to Standard results, split at pivot date
2. **Domain categorization** — Built-in mapping + optional Claude API for unknowns
3. **Source type aggregation** — Slot-days per category per period
4. **Executive summary** — Claude API or fallback template generation
5. **PPTX generation** — python-pptx with hand-drawn shape charts matching Five Blocks branding
6. **Excel generation** — openpyxl with formatted headers and conditional fills

## Design System

- **Colors**: Navy (#1A365D), Teal (#0D7C7D), Gray (#9CA3AF), Dark (#1F2937)
- **Fonts**: Calibri Light for titles, Calibri for body
- **Charts**: Hand-drawn rectangle shapes (not native PPTX chart objects)
- **Branding**: Five Blocks logo in footer of every slide
