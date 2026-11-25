set shell := ["bash", "-euo", "pipefail", "-c"]

venv := ".venv"
python := venv + "/bin/python"
pip := venv + "/bin/pip"
gcloud := "./google-cloud-sdk/bin/gcloud"
project := "project-20839851-36ff-4357-ae9"

# List available recipes
default:
    @just --list

# Create the Python virtual environment and install dependencies
setup:
    python3 -m venv {{venv}}
    {{pip}} install --upgrade pip
    {{pip}} install openpyxl google-cloud-translate

# Authenticate Application Default Credentials in browser
gcloud-auth:
    {{gcloud}} auth application-default login

# Configure gcloud project, ADC quota project, and enable Cloud Translation API
gcloud-configure:
    {{gcloud}} config set project {{project}}
    {{gcloud}} auth application-default set-quota-project {{project}}
    {{gcloud}} services enable translate.googleapis.com --project={{project}}

# Check current gcloud account, project, ADC status, and Translation API
gcloud-status:
    {{gcloud}} auth list
    {{gcloud}} config list
    {{gcloud}} auth application-default print-access-token >/dev/null && echo "ADC: OK"
    {{gcloud}} services list --enabled --project={{project}} --filter="name:translate.googleapis.com"

# Run a local dry translation without calling an external API
dry-run:
    {{python}} viessmann_translate.py --provider none

# Run translation using Google Cloud Translation v3
translate:
    {{python}} viessmann_translate.py --provider google-v3 --project {{project}}
