#!/bin/bash
set -euo pipefail

function_name="mindset-assessments"
service_name="mindset-assessments"
job_name="mindset-assessments"
image_name="mindset-assessments"
entry_point="mindset_assessments"
region="europe-west2"
project="focal-rig-396713"
timezone="Europe/London"
schedule="0 */2 * * *"
scheduler_service_account="640977351943-compute@developer.gserviceaccount.com"
function_url="https://${region}-${project}.cloudfunctions.net/${function_name}"

printf 'Building Docker image: %s\n' "$image_name"
docker build -t "$image_name" .

function_exists=false
service_exists=false

if gcloud functions describe "$function_name" --gen2 --region="$region" --project="$project" >/dev/null 2>&1; then
  function_exists=true
fi

if gcloud run services describe "$service_name" --region="$region" --project="$project" >/dev/null 2>&1; then
  service_exists=true
fi

if [ "$function_exists" = true ]; then
  printf 'Updating existing function: %s\n' "$function_name"
else
  if [ "$service_exists" = true ]; then
    printf 'Deleting stale Cloud Run service before creating function: %s\n' "$service_name"
    gcloud run services delete "$service_name" \
      --region="$region" \
      --project="$project" \
      --quiet
  fi
  printf 'Creating new function: %s\n' "$function_name"
fi

gcloud functions deploy "$function_name" \
  --gen2 \
  --region="$region" \
  --project="$project" \
  --runtime=python311 \
  --source=. \
  --entry-point="$entry_point" \
  --trigger-http \
  --no-allow-unauthenticated \
  --memory=1Gi \
  --timeout=3600

printf 'Routing all traffic to the latest revision for service: %s\n' "$service_name"
gcloud run services update-traffic "$service_name" \
  --region="$region" \
  --project="$project" \
  --to-latest

printf 'Rebuilding Cloud Scheduler job: %s\n' "$job_name"
gcloud scheduler jobs delete "$job_name" \
  --location="$region" \
  --project="$project" \
  --quiet >/dev/null 2>&1 || true

gcloud scheduler jobs create http "$job_name" \
  --location="$region" \
  --project="$project" \
  --schedule="$schedule" \
  --time-zone="$timezone" \
  --uri="$function_url" \
  --http-method=POST \
  --message-body='{"name":"Scheduler"}' \
  --oidc-service-account-email="$scheduler_service_account"

gcloud scheduler jobs describe "$job_name" \
  --location="$region" \
  --project="$project"
