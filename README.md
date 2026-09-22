# JobBot

Tailored resume + cover letter web app (Flask).

## Local run

The site is open to anyone with the URL when `APP_PASSWORD` is unset or empty. Set a non-empty `APP_PASSWORD` only if you want to require a password.

```powershell
pip install -r requirements.txt
python app.py
```

Optional password gate:

```powershell
$env:APP_PASSWORD = "your-password"
python app.py
```

Open http://127.0.0.1:8000

## Azure

Deploy with `deploy.ps1` from the deploy package, or zip-deploy to App Service `tchamna-jobbot` (plan `ASP-ragaifoundationsdemorg-87d6`, Canada Central).

Startup: `gunicorn --bind=0.0.0.0:8000 --timeout 120 --workers 2 app:app`
