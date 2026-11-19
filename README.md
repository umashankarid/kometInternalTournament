# kometInternalTournament

Web app for internal Komet badminton tournaments.

- **Admin**: create/edit/delete tournaments (name, dates, registration deadline, cancellation deadline), add
  categories (Singles, Doubles, Mixed, or any custom name with a type), view registrations as a table and
  export to Excel (per tournament, or everything incl. a Users sheet).
- **Players**: sign up (name, email, phone, gender, date of birth), browse tournaments, register for
  categories. For Doubles/Mixed the partner name is optional - leave it empty to be "open to any partner".
  Players can cancel until the cancellation deadline. Registration closes after the registration deadline.
- The Excel export contains DOB and **age at tournament start** for later age separation.

Stack: Flask + SQLite (single container, no Node build), gunicorn. Same shape as bmktimereport.

## Deploy on Coolify

1. Push this folder to a git repo and create a new Coolify resource from it.
   Build pack: **Dockerfile** (or **Docker Compose** using `docker-compose.yml`). Port: **5000**.
2. Domain: set to `https://internal.bmkkomet.se` in the resource's *Domains* field, and create a DNS
   record for that host pointing at the Coolify server (A/CNAME). Coolify's proxy issues the TLS certificate.
3. Environment variables (see `.env.example`): `ADMIN_EMAIL`, `ADMIN_PASSWORD` (required), `ADMIN_NAME`,
   `SECRET_KEY` (optional), `COOKIE_SECURE=1`, `TZ`.
4. Persistent storage: mount a volume at **`/app/data`** (the compose file does this). It holds the SQLite
   database (`tournaments.db`) and the generated secret key. Without it, data is lost on redeploy.
5. Health check path: `/api/health`.

Log in with `ADMIN_EMAIL` / `ADMIN_PASSWORD` to reach the **Admin** menu. If a normal user already
registered with `ADMIN_EMAIL`, that user is promoted to admin on startup (their password is kept).

## Run locally

```bash
python3 -m venv venv && . venv/bin/activate
pip install -r requirements.txt
ADMIN_EMAIL=admin@example.com ADMIN_PASSWORD=secret123 COOKIE_SECURE=0 python app.py
# http://localhost:5000
```

Or with Docker:

```bash
cp .env.example .env   # edit values, set COOKIE_SECURE=0 for http
docker compose up --build
```

## Tests

```bash
pip install pytest && python -m pytest -q
```

## Notes

- "Registration date" and "cancellation date" are implemented as deadlines: players can register up to and
  including the registration deadline, and cancel up to and including the cancellation deadline (admin can
  remove registrations any time).
- Backup: copy `tournaments.db` from the `/app/data` volume.
