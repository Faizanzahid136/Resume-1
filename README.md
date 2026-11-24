# Resume Tailor (Streamlit)

Upload a CV once, add a job post, get a CV that fits the job. Only facts from the person's own CV are used.

## Put it online (free, about 10 minutes)

1. Create a free GitHub account and a new **public** repository (for example `resume-tailor`).
2. Upload these files to the repository: `app.py`, `rules.txt`, `cv_rules.txt`, `requirements.txt`, and the folder `.streamlit/` (with `config.toml`).
3. Go to https://share.streamlit.io and sign in with GitHub.
4. Click **Create app**, pick your repository, set the main file to `app.py`, and click **Deploy**.
5. Open **Settings > Secrets** for the app and add your Anthropic API key:

       ANTHROPIC_API_KEY = "sk-ant-..."

   Get a key at https://console.anthropic.com (it is billed separately from a Claude chat subscription).
   Without a key in Secrets, each visitor can paste their own key in the sidebar instead.

## Run it on your own computer

    pip install -r requirements.txt
    streamlit run app.py

## Notes
- Nothing is stored on a server. A CV only lives while the page is open, so visitors download their CV before leaving.
- If you set your own key in Secrets, every visitor's use is billed to you (a few cents per CV). Consider asking visitors to bring their own key.
- Optional: set `MODEL` in Secrets to change the Claude model.
