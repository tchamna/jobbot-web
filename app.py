"""jobbot web — tailored resume + cover letter generator (Flask)."""
from __future__ import annotations
import csv, datetime as dt, io, json, os, re, secrets, shutil, zipfile
from functools import wraps
from pathlib import Path

import yaml
from flask import (Flask, abort, flash, redirect, render_template, request,
                   send_file, session, url_for)
from markupsafe import Markup, escape

import jobbot as jb

APP_DIR = Path(__file__).resolve().parent
DATA = Path(os.environ.get("DATA_DIR", APP_DIR / "data"))      # /home/data on Azure = persistent
JOBS = DATA / "jobs"
PROFILE = DATA / "profile.yaml"
OVR = DATA / "overrides"
TRACK = DATA / "tracker.csv"
FIELDS = ["id", "date", "company", "title", "location", "salary", "url", "arch", "status", "notes"]
STATUSES = ["prepared", "applied", "interview", "offer", "rejected", "withdrawn"]

for d in (JOBS, OVR):
    d.mkdir(parents=True, exist_ok=True)
if not PROFILE.exists():
    shutil.copy(APP_DIR / "profile.yaml", PROFILE)
for f in (APP_DIR / "overrides").glob("*.yaml"):
    if not (OVR / f.name).exists():
        shutil.copy(f, OVR / f.name)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  SESSION_COOKIE_SECURE=bool(os.environ.get("WEBSITE_SITE_NAME")),
                  PERMANENT_SESSION_LIFETIME=dt.timedelta(days=14))
# Empty or unset APP_PASSWORD leaves the site public. A non-empty value
# restores the optional password gate.
PASSWORD = os.environ.get("APP_PASSWORD", "").strip()

# ------------------------------------------------------------------ helpers
def login_required(f):
    @wraps(f)
    def w(*a, **k):
        if not PASSWORD or session.get("ok"):
            return f(*a, **k)
        return redirect(url_for("login", next=request.path))
    return w

def load_profile():
    return yaml.safe_load(PROFILE.read_text(encoding="utf-8"))

def rows():
    if not TRACK.exists():
        return []
    return list(csv.DictReader(TRACK.open(encoding="utf-8")))

def save_rows(rs):
    with TRACK.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
        w.writeheader(); w.writerows(rs)

def job_dir(jid): 
    p = JOBS / str(int(jid))
    if not p.exists():
        abort(404)
    return p

def generate(job: dict, extra: dict, use_llm: bool) -> str:
    profile = load_profile()
    sel = jb.select(profile, job)
    of = OVR / f"{jb.slug(job['company'])}.yaml"
    ovr = (yaml.safe_load(of.read_text(encoding="utf-8")) if of.exists() else {}) or {}
    ovr.update({k: v for k, v in extra.items() if v})
    if use_llm:
        ovr = jb.llm_polish(profile, job, sel, ovr)
    sel["headline"], sel["summary"] = ovr.get("headline"), ovr.get("summary")
    rs = rows()
    jid = str(max([int(r["id"]) for r in rs] or [0]) + 1)
    d = JOBS / jid
    d.mkdir(parents=True, exist_ok=True)
    last = profile["name"].split(",")[0].split()[-1]
    s = jb.slug(job["company"])
    jb.build_resume(profile, sel, d / f"Resume_{last}_{s}.docx")
    blocks = jb.build_cover(profile, sel, job, ovr, d / f"CoverLetter_{last}_{s}.docx")
    jb.write_report(job, sel, d / "fit_report.md")
    a = profile["archetypes"][sel["arch"]]
    view = {
        "profile": profile, "job": job, "arch": sel["arch"],
        "headline": sel["headline"] or a["headline"], "summary": sel["summary"] or a["summary"],
        "skills": sel["skills"][:5], "experience": sel["experience"], "projects": sel["projects"],
        "matched": sel["matched"], "gaps": sel["gaps"], "cover": blocks,
    }
    (d / "view.json").write_text(json.dumps(view, ensure_ascii=False), encoding="utf-8")
    (d / "job.txt").write_text(job["text"], encoding="utf-8")
    rs.append({"id": jid, "date": dt.date.today().isoformat(), "company": job["company"],
               "title": job["title"], "location": job.get("location", ""), "salary": job.get("salary", ""),
               "url": job.get("url", ""), "arch": sel["arch"], "status": "prepared", "notes": ""})
    save_rows(rs)
    return jid

# ------------------------------------------------------------------ routes
@app.route("/login", methods=["GET", "POST"])
def login():
    if not PASSWORD:
        return redirect(url_for("home"))
    if request.method == "POST":
        if secrets.compare_digest(request.form.get("password", ""), PASSWORD):
            session.permanent = True; session["ok"] = True
            nxt = request.args.get("next") or "/"
            return redirect(nxt if nxt.startswith("/") and not nxt.startswith("//") else "/")
        flash("Wrong password.")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home") if not PASSWORD else url_for("login"))

@app.route("/", methods=["GET", "POST"])
@login_required
def home():
    if request.method == "POST":
        url = request.form.get("url", "").strip()
        text = request.form.get("text", "").strip()
        try:
            if text:
                meta = dict(re.findall(r"^(URL|TITLE|COMPANY|LOCATION|SALARY):\s*(.+)$", text, re.M))
                job = {"url": url or meta.get("URL", ""), "title": meta.get("TITLE", ""),
                       "company": meta.get("COMPANY", ""), "location": meta.get("LOCATION", ""),
                       "salary": meta.get("SALARY", ""), "text": text}
            elif url:
                job = jb.fetch_job(url)
            else:
                flash("Paste a job link or the job description."); return redirect("/")
        except Exception as e:
            flash(f"{e}".split("\n")[0] + " Paste the job description text instead.")
            return render_template("home.html", form=request.form, recent=rows()[-5:][::-1])
        for k in ("title", "company", "location", "salary"):
            if request.form.get(k, "").strip():
                job[k] = request.form[k].strip()
        if not job["title"] or not job["company"]:
            flash("Please fill in the job title and company.")
            return render_template("home.html", form=request.form, recent=rows()[-5:][::-1])
        extra = {"why_company": request.form.get("why_company", "").strip(),
                 "company_name": re.sub(r",?\s*(Inc|LLC|Ltd|Corp)\.?$", "", job["company"]).strip()}
        jid = generate(job, extra, bool(request.form.get("llm")))
        return redirect(url_for("result", jid=jid))
    return render_template("home.html", form={}, recent=rows()[-5:][::-1],
                           llm=bool(os.environ.get("ANTHROPIC_API_KEY")))

@app.route("/job/<jid>")
@login_required
def result(jid):
    d = job_dir(jid)
    view = json.loads((d / "view.json").read_text(encoding="utf-8"))
    row = next((r for r in rows() if r["id"] == str(jid)), {})
    files = sorted(p.name for p in d.glob("*.docx"))
    report = (d / "fit_report.md").read_text(encoding="utf-8")
    return render_template("result.html", v=view, jid=jid, row=row, files=files,
                           report=report, statuses=STATUSES)

@app.route("/job/<jid>/print/<kind>")
@login_required
def printable(jid, kind):
    d = job_dir(jid)
    view = json.loads((d / "view.json").read_text(encoding="utf-8"))
    if kind not in ("resume", "cover"):
        abort(404)
    return render_template(f"print_{kind}.html", v=view)

@app.route("/job/<jid>/file/<name>")
@login_required
def download(jid, name):
    d = job_dir(jid)
    p = (d / name).resolve()
    if p.parent != d.resolve() or not p.exists() or p.suffix not in (".docx", ".md"):
        abort(404)
    return send_file(p, as_attachment=True)

@app.route("/job/<jid>/zip")
@login_required
def zipall(jid):
    d = job_dir(jid)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in d.iterdir():
            if p.suffix in (".docx", ".md"):
                z.write(p, p.name)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"application_{jid}.zip")

@app.route("/job/<jid>/status", methods=["POST"])
@login_required
def set_status(jid):
    rs = rows()
    for r in rs:
        if r["id"] == str(jid):
            if request.form.get("status") in STATUSES:
                r["status"] = request.form["status"]
            if "notes" in request.form:
                r["notes"] = request.form["notes"][:500]
    save_rows(rs)
    return redirect(request.referrer or url_for("tracker"))

@app.route("/tracker")
@login_required
def tracker():
    rs = rows()[::-1]
    counts = {s: sum(r["status"] == s for r in rs) for s in STATUSES}
    return render_template("tracker.html", rows=rs, statuses=STATUSES, counts=counts)

@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    if request.method == "POST":
        txt = request.form.get("yaml", "")
        try:
            data = yaml.safe_load(txt)
            assert isinstance(data, dict) and "experience" in data and "archetypes" in data
        except Exception as e:
            flash(f"Not saved: the profile has a formatting error ({str(e)[:160]}).")
            return render_template("profile.html", text=txt)
        shutil.copy(PROFILE, DATA / f"profile.backup.{dt.datetime.now():%Y%m%d%H%M%S}.yaml")
        PROFILE.write_text(txt, encoding="utf-8")
        flash("Profile saved. New applications will use it.")
        return redirect(url_for("profile"))
    return render_template("profile.html", text=PROFILE.read_text(encoding="utf-8"))

@app.route("/healthz")
def health():
    return "ok"

if __name__ == "__main__":
    app.run(debug=True, port=8000)
