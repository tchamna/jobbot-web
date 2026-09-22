# JobBot

Tailored resume + cover letter web app (Flask).

## Local run

```powershell
pip install -r requirements.txt
$env:APP_PASSWORD = "your-password"
python app.py
```

Open http://127.0.0.1:8000

## Azure

Deploy with `deploy.ps1` from the deploy package, or zip-deploy to App Service `tchamna-jobbot` (plan `ASP-ragaifoundationsdemorg-87d6`, Canada Central).

Startup: `gunicorn --bind=0.0.0.0:8000 --timeout 120 --workers 2 app:app`
