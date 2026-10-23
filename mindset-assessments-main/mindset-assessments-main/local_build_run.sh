#!/bin/bash

current_directory=$(pwd)  # Get the current working directory
current_directory_name=$(basename "$current_directory")
credentials_file="${GOOGLE_APPLICATION_CREDENTIALS:-$HOME/.config/gcloud/application_default_credentials.json}"

echo "The application name is: $current_directory_name"

if [ ! -f "$credentials_file" ]; then
  echo "Missing Google credentials file: $credentials_file" >&2
  exit 1
fi

# Build the Docker image
docker build -t "$current_directory_name" .

# Run the container, and mount current dir to get the profile output
docker run --rm \
  -v "$credentials_file:/tmp/gcp-adc.json:ro" \
  -e GOOGLE_APPLICATION_CREDENTIALS=/tmp/gcp-adc.json \
  -v "$current_directory:/app" \
  "$current_directory_name"
