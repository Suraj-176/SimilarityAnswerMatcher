# Semantic Answer Matcher

This project is a Streamlit app that compares answers from two Excel files using state-of-the-art open-source models. It computes semantic similarity, highlights differences, and provides explanations for each answer pair.

## Features
- Upload two Excel files (questions in column 1, answers in column 2)
- Choose from multiple local language models
- Computes semantic similarity, fuzzy match, and cross-encoder scores
- Highlights differences between answers
- Download results as Excel
- All computation is local (no data leaves your machine)

## Requirements
- Python 3.8+
- See `requirements.txt` for dependencies

## Usage

Below are clear, copyable commands for Windows PowerShell (adjust paths if your workspace is in a different location).

1) Activate the virtual environment (recommended)

```powershell
& 'C:/SemanticAnswerMatcher/.venv/Scripts/Activate.ps1'
```

2) Install dependencies (only needed once or when requirements change)

```powershell
pip install -r requirements.txt
```

3) Run the Streamlit app (after activating the venv)

```powershell
streamlit run app.py
```

Alternative (run without activating the venv):

```powershell
C:/SemanticAnswerMatcher/.venv/Scripts/python.exe -m streamlit run app.py
```

Change the port (if 8501 is in use):

```powershell
C:/SemanticAnswerMatcher/.venv/Scripts/python.exe -m streamlit run app.py --server.port 8502
```

4) Open the local URL printed by Streamlit (usually http://localhost:8501). If you changed the port, use that port instead.

Sharing the app with colleagues using LocalTunnel (quick & temporary)
-----------------------------------------------------------------

LocalTunnel (https://localtunnel.github.io/www/) exposes a local server to the internet via a simple command-line tool called `lt`. Below are the steps for Windows PowerShell.

Prerequisite: have Node.js and npm installed so you can install the `localtunnel` package globally (one-time):

```powershell
npm install -g localtunnel
```

Start your Streamlit app locally (see step 3). Then, in a separate PowerShell window run:

```powershell
lt --port 8501 --subdomain yourname-demo
```

Notes:
- Replace `8501` with the port your Streamlit server is running on.
- Replace `yourname-demo` with any available subdomain (it may be claimed already). If the requested subdomain is unavailable, `lt` will return an error and you can try a different subdomain or omit the `--subdomain` flag to get a random public URL.
- The `lt` process keeps running and prints a publicly accessible URL (e.g. https://yourname-demo.loca.lt) — share that URL with your colleague. When you stop `lt` (Ctrl+C) the URL will stop serving.

Security and privacy reminders:
- LocalTunnel makes your local server accessible to the public. Only share with trusted colleagues and avoid exposing sensitive data.
- The streamlit app may load local models which can be resource intensive; ensure your machine has enough memory/CPU for collaborators' use.

Common troubleshooting
----------------------
- If `streamlit run` fails, first ensure your virtual environment is activated and the `streamlit` package is installed in that environment.
- If `lt` can't bind to the requested subdomain, try a different subdomain or omit `--subdomain` to get an ephemeral URL.
- If the app is reachable locally but not via LocalTunnel, check firewall rules and corporate network restrictions.

RECOMMENDED: Use LocalTunnel (or similar) to expose your app

Here’s a short recap you can copy-paste (matches the quick steps above):

1) Install Node.js (skip if already installed):

	- Download and install the LTS build from https://nodejs.org/

2) Install LocalTunnel globally:

```powershell
npm install -g localtunnel
```

3) Run your local Streamlit app (default port 8501):

```powershell
streamlit run app.py
```

4) Expose port 8501 with LocalTunnel (simple):

```powershell
lt --port 8501
```

Output example:

```
your url is: https://magical-bear.loca.lt
```

Optional: request a custom subdomain (must be unique):

```powershell
lt --port 8501 --subdomain yourname-demo
```

If the subdomain is available you will get `https://yourname-demo.loca.lt`. If it's already taken, choose another name or omit the `--subdomain` flag.

Security reminder: LocalTunnel makes your local server publicly reachable. Only share the URL with trusted users and stop the `lt` process when you're done.

Optional: Add a simple authentication gate
----------------------------------------

The app supports a lightweight, environment-variable controlled username/password gate. This is intended for short demos shared via LocalTunnel and is NOT a production-grade authentication system. For production, use a proper auth provider or reverse proxy.

Set these environment variables (PowerShell examples):

```powershell
$env:AUTH_ENABLED = "true"
$env:AUTH_USERNAME = "alice"
$env:AUTH_PASSWORD = "s3cret"
```

Then start the app as usual. When `AUTH_ENABLED` is true, users will need to sign in with the configured username and password before they can access the app UI. Share the credentials separately (secure channel) with colleagues when you give them the LocalTunnel URL.

Example sharing workflow with credentials
----------------------------------------
1. Start the app locally:

```powershell
streamlit run app.py
```

2. In another PowerShell, run LocalTunnel and request a subdomain:

```powershell
lt --port 8501 --subdomain yourname-demo
```

3. Copy the public URL printed by `lt` and send it to your colleague together with the username/password you configured via environment variables.

4. Stop `lt` when finished (Ctrl+C).


## Notes
- For best results, use the recommended models in the dropdown.
- All processing is done locally for privacy.
