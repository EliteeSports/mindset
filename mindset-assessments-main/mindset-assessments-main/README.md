# Lighthouse Mindset Assessments

This service automates the Lighthouse "Mindset Advantage" assessment workflow. It runs as an HTTP-triggered Google Cloud Function, identifies candidates who need assessments, sends assessment notifications, tracks report files across BigQuery, Cloud Storage, SharePoint, and Salesforce, updates candidate stages when all four reports are complete, and rebuilds a consolidated Excel tracking report.

---

## 1) How it works and what it generates

**High-level flow**
1. The HTTP function entry point `mindset_assessments()` is invoked by Cloud Scheduler or a manual HTTP call.
2. The service loads:
   - Salesforce credentials from Secret Manager secret `leathwaite-ai-sf-prod`
   - Microsoft Graph credentials from Secret Manager secret `leathwaite`
3. It queries Salesforce for candidates who:
   - were sent for assessment
   - have not completed the assessment
   - belong to roles with Growth Mindset criteria
4. For new candidates, it:
   - resolves search lead users
   - emails the candidate when a personal email exists
   - inserts four report-tracking rows into BigQuery
   - notifies the search lead
   - reverts the Salesforce flag when the candidate cannot be contacted
5. It scans GCS bucket `lw-ms` for completed report PDFs and uploads valid files to SharePoint.
6. It updates BigQuery with SharePoint metadata for uploaded reports.
7. It downloads completed reports from SharePoint and uploads them into Salesforce as `ContentVersion` records.
8. It notifies search leads when all four reports for a candidate are available.
9. It moves candidates to the "assessment complete" stage in Salesforce and writes an invoice ID back to BigQuery.
10. It rebuilds `report_results.xlsx` and uploads it to both SharePoint and GCS.
11. It emails a pass/fail log attachment at the end of the run.

**What it generates**
- Candidate assessment invitation and notification emails
- Four tracked report records per candidate in BigQuery
- Uploaded PDFs in SharePoint and Salesforce
- Updated Salesforce assessment completion state
- Consolidated Excel report `report_results.xlsx`
- End-of-run log email with `log.txt`

---

## 2) Trigger and workflow behavior

The deployed service is an HTTP-triggered Cloud Function, but the handler does not inspect the request body in any meaningful way.

**Important behavior**
- Any caller that can invoke the function can trigger a full assessment run.
- The function returns:
  - `("OK", 200)` on success
  - `("ERROR", 500)` on an unhandled exception
- The candidate selection logic currently includes a hard-coded Salesforce date threshold inside `get_candidate_details_from_lh()`.

**Report expectations**
- The workflow expects exactly these report names:
  - `Mindset Advantage Consultant Report`
  - `Mindset Advantage Client Report_Portrait`
  - `Mindset Advantage Client Report_Landscape`
  - `Mindset Advantage Personal Report`
- The filename pattern is:

```text
<search_id>–<contact_id>–<Report Name>.pdf
```

- The separator is an en dash (`–`), not a normal hyphen.

---

## 3) Deploying to GCP with `gcp_build.sh`

Script summary (from `gcp_build.sh`):
1. Builds the local Docker image for a quick container check.
2. Deploys a Gen 2 HTTP Cloud Function in `europe-west2`.
3. Uses runtime `python311`.
4. Uses explicit entry point `mindset_assessments`.
5. Disables unauthenticated access.
6. Sets memory to `1Gi`.
7. Sets timeout to `3600` seconds.
8. Routes Cloud Run traffic to the latest revision.
9. Deletes and recreates the 2-hour Cloud Scheduler trigger.

**Command**
```bash
./gcp_build.sh
```

**Notes**
- The script expects `gcloud` to be installed and authenticated.
- It targets:
  - Project: `focal-rig-396713`
  - Region: `europe-west2`
- The configured scheduler job is:
  - every 2 hours
  - on the hour
  - London time
- The caller identity is:
  - `640977351943-compute@developer.gserviceaccount.com`

---

## 4) Local testing with `local_build_run.sh`

**Build + run locally**
```bash
./local_build_run.sh
```

This will:
- resolve Google credentials from `GOOGLE_APPLICATION_CREDENTIALS`, or fall back to `~/.config/gcloud/application_default_credentials.json`
- mount the credentials file into Docker as `/tmp/gcp-adc.json`
- set `GOOGLE_APPLICATION_CREDENTIALS=/tmp/gcp-adc.json` inside the container
- build the Docker image from `Dockerfile`
- run the container
- mount the repository into `/app` so generated outputs such as logs and reports remain visible locally

If the credentials file is missing, run:
```bash
gcloud auth application-default login
```

**Run directly with Python**
```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

**Notes**
- Local execution still requires live access to:
  - Google Secret Manager
  - BigQuery
  - BigQuery Storage API
  - Cloud Storage
  - Salesforce
  - Microsoft Graph / SharePoint
- Running `python main.py` executes the full workflow immediately; it does not start a persistent web server.

---

## 5) Prerequisites on a brand-new server

**System prerequisites**
- Python 3.11
- Docker, if you want to use the container workflow
- `gcloud` CLI, if you want to deploy or manage the scheduler from that machine

**GCP prerequisites**
- Access to project `focal-rig-396713`
- Permission to deploy Gen 2 Cloud Functions in `europe-west2`
- Permission to manage Cloud Scheduler jobs in `europe-west2`
- Permission to access Google Secret Manager secrets in project `640977351943`
- Permission to read and update:
  - BigQuery dataset `CandidateAssessment`
  - Cloud Storage bucket `lw-ms`

**Secrets required in GCP Secret Manager**
- `leathwaite-ai-sf-prod`
  - Expected keys: `username`, `password`, `token`, `domain`
- `leathwaite`
  - Expected keys: `tenant_id`, `client_id`, `client_secret`

**Other external requirements**
- Outbound HTTPS access to:
  - Salesforce
  - `graph.microsoft.com`
  - `login.microsoftonline.com`
  - Google Cloud APIs
- Microsoft Graph permissions that allow:
  - SharePoint file access under `CandidateAssessment`
  - sending workflow emails

**Data stores used by the workflow**
- BigQuery table: `focal-rig-396713.CandidateAssessment.candidate_assessments`
- GCS bucket: `lw-ms`
- SharePoint folders under `CandidateAssessment/`

---

## File map

- `main.py`: Full assessment workflow, BigQuery updates, GCS processing, SharePoint integration, Salesforce updates, and notification emails.
- `gcp_build.sh`: Builds the image, deploys the Gen 2 Cloud Function with `1Gi` memory, routes latest traffic, and recreates the 2-hour Cloud Scheduler trigger.
- `local_build_run.sh`: Builds and runs the workflow locally in Docker with Application Default Credentials mounted into the container.
- `Dockerfile`: Python 3.11 container image for local execution.
- `requirements.txt`: Python dependencies.
