# PV-Blink Inverter & Banga Solar Pvt Ltd — DailyAnalysis

Streamlit analytics for sales, returns, field activity and service operations.

Select **PV-Blink Inverter** (orange) or **Banga Solar Pvt Ltd** (dark green)
from the company selector. Both use the same analytics and Excel processing.
Switching companies clears uploaded files and filters; upload the selected
company's reports afterward. Sales, returns, field and service inputs support
multiple files. Avoid overlapping reports to prevent double counting.

This repository contains the application source, not a hosted dashboard.

| Setting | Value |
| --- | --- |
| Main entry file | `pv_blink_dashboard.py` |
| Framework | Streamlit 1.63.0 |
| Python | 3.11 |
| Container port | 8501 |
| Login helper | `PV_Blink_Login_Form.py` |
| Required asset | `assets/pvblink_logo.png` |

Docker automatically starts the main Streamlit file. **No manual Python command
is needed after the container starts.** The deployment administrator configures
the final browser URL/hostname and its mapping to the container.

## Docker deployment

Install Docker Engine on the server, or use Docker Desktop in Linux-container
mode for local testing. Verify `docker version` and `docker info` work. Copy or
clone the repository, then run these commands from the folder with `Dockerfile`:

```sh
docker build -t dailyanalysis:1.0 .
docker run -d --name dailyanalysis --restart unless-stopped -p 8501:8501 dailyanalysis:1.0
```

For a local test, open `http://localhost:8501`. On the server, use the address
provided by the deployment administrator. The server needs outbound DNS/HTTPS
access to the login API; the build needs access to the image and package registries.

The image contains only the requirements, the two application Python files and
the logo. It runs as a non-root user. Local virtual environments, reports, Git
metadata, tests and `.env` files are excluded from the build context.

Exact automatic startup command:

```sh
python -m streamlit run pv_blink_dashboard.py --server.address=0.0.0.0 --server.port=8501 --server.headless=true
```

## Environment configuration

No environment variable is mandatory with the existing default API endpoint.

| Variable | Default / meaning |
| --- | --- |
| `PVBLINK_LOGIN_API_URL` | Optional override. Unset or blank uses `https://solar.e-khata.com/api/acct-crm/user/login`. |
| `STREAMLIT_SERVER_MAX_UPLOAD_SIZE` | Optional upload limit in MB; Streamlit defaults to 200. |

Copy `.env.example` to `.env` and edit it if needed. Pass it explicitly at container
creation; neither Python nor Docker automatically reads this file:

```sh
docker run -d --name dailyanalysis --restart unless-stopped --env-file .env -p 8501:8501 dailyanalysis:1.0
```

Use this **instead of** the first run command, not as a second container with the
same name. Keep credentials out of source control. Users log in through the app;
no account password or token needs to be configured in the image.

## Health, logs and lifecycle

```sh
docker ps --filter name=dailyanalysis
docker logs --tail 100 dailyanalysis
docker inspect --format='{{.State.Health.Status}}' dailyanalysis
docker restart dailyanalysis
docker stop dailyanalysis
docker start dailyanalysis
```

The Python-based healthcheck requests `http://127.0.0.1:8501/_stcore/health` every
30 seconds. It checks Streamlit availability, not login API credentials or
spreadsheet processing. Confirm browser login and upload a report after deployment.
The restart policy handles process exits and Docker restarts unless manually
stopped; an unhealthy status alone does not restart a container.

To update, copy/pull the reviewed changes, build first, then replace the container:

```sh
docker build -t dailyanalysis:1.1 .
docker stop dailyanalysis
docker rm dailyanalysis
docker run -d --name dailyanalysis --restart unless-stopped -p 8501:8501 dailyanalysis:1.1
```

Reuse your environment options and administrator-approved port binding. Keep the
previous image tag for rollback using the same replacement sequence.

## Local development and validation

Create a Python 3.11 environment (virtual environments are not included):

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run pv_blink_dashboard.py
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

`run_dashboard.bat` uses `.venv\Scripts\python.exe`, with `venv` as a fallback.
No PowerShell environment activation is required.

Linux equivalents:

```sh
python3.11 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m streamlit run pv_blink_dashboard.py
```

Validation performed on 2026-09-23: fresh isolated dependency installation,
`pip check`, runtime imports, syntax, regression tests, login-page rendering,
authenticated navigation, sales/return test fixtures, the supplied service workbook,
and a local Streamlit HTTP health check. Test authentication responses are mocked;
no real credentials are included or submitted by the test suite.

**Docker validation completed on 2026-09-24:** image `dailyanalysis:1.0` built
successfully on Docker Desktop's Linux engine. Container `dailyanalysis` started
automatically as `appuser` (UID 10001), reported `healthy` with zero restarts,
and returned HTTP 200 for `/` and `/_stcore/health` (body `ok`). Startup logs
contained no application traceback, and `pip check` found no broken requirements.
All 10 regression tests also passed inside the Linux image with network access
disabled and the test folder mounted read-only.

The demonstration container is bound to `127.0.0.1:8501` on the development PC;
open `http://localhost:8501` on that PC. This is a local Docker deployment, not a
company-hosted URL. Live login with an approved account and browser interaction
with custom sidebar/loading JavaScript remain manual checks. The administrator
must still deploy and verify the image on the target server.

## Data and repository notes

- Uploaded reports and session state are in memory. Restarting or replacing the
  container resets sessions; users must log in and upload again. Retain original
  workbooks separately. No database or persistent report storage is configured.
- Existing sales/returns calculations, representative attribution, product lists,
  service grouping and date-filter rules are preserved. Missing optional Rate or
  representative columns now load safely; XLS field reports use the detected
  Excel engine. Distinct ticket records with different statuses are not silently
  merged because that requires a business decision.
- Virtual environments, bytecode, credentials and uploaded Excel reports are
  excluded from the GitHub source upload and Docker image.
- If the page is unreachable, check container logs, port mapping and the
  administrator's network configuration. If port 8501 is occupied locally, use
  `-p 8502:8501` and open `http://localhost:8502`.
- If login fails, check the configured API endpoint and outbound connectivity.
  If an upload fails, read the page's processing error and check the report format.

Server-specific DNS, hostname, TLS, reverse proxy and firewall settings belong to
the deployment administrator and are not configured by this repository.

References: [Streamlit Docker deployment](https://docs.streamlit.io/deploy/tutorials/docker)
and [Docker restart policies](https://docs.docker.com/engine/containers/start-containers-automatically/).
