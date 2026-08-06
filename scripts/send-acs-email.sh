#!/usr/bin/env bash

set -Eeuo pipefail

readonly ACS_EMAIL_API_VERSION='2023-03-31'
readonly ACS_TOKEN_RESOURCE='https://communication.azure.com'
readonly POLL_TIMEOUT_SECONDS=300

body_file="${1:-}"
endpoint="${AZURE_COMMUNICATION_ENDPOINT:-}"
sender="${AZURE_EMAIL_SENDER:-}"
recipient="${AZURE_EMAIL_RECIPIENT:-}"
subject="${AZURE_EMAIL_SUBJECT:-}"

for command in az curl jq awk mktemp; do
  command -v "$command" >/dev/null 2>&1 || {
    echo "::error::Required command is unavailable: $command"
    exit 1
  }
done

[[ -n "$body_file" && -f "$body_file" ]] || {
  echo "::error::Pass a readable email body file to send-acs-email.sh."
  exit 1
}
[[ "$endpoint" =~ ^https://[^/[:space:]]+/?$ ]] || {
  echo "::error::AZURE_COMMUNICATION_ENDPOINT must be an HTTPS origin."
  exit 1
}
[[ "$sender" =~ ^[^[:space:]@]+@[^[:space:]@]+$ ]] || {
  echo "::error::AZURE_EMAIL_SENDER must be a valid email address."
  exit 1
}
[[ "$recipient" =~ ^[^[:space:]@]+@[^[:space:]@]+$ ]] || {
  echo "::error::AZURE_EMAIL_RECIPIENT must be a valid email address."
  exit 1
}
[[ -n "$subject" ]] || {
  echo "::error::AZURE_EMAIL_SUBJECT is required."
  exit 1
}

endpoint="${endpoint%/}"
payload_file="$(mktemp)"
headers_file="$(mktemp)"
response_file="$(mktemp)"
trap 'rm -f "$payload_file" "$headers_file" "$response_file"' EXIT

jq --null-input \
  --arg sender "$sender" \
  --arg recipient "$recipient" \
  --arg subject "$subject" \
  --rawfile body "$body_file" \
  '{
    senderAddress: $sender,
    recipients: {
      to: [{ address: $recipient }]
    },
    content: {
      subject: $subject,
      plainText: $body
    }
  }' > "$payload_file"

access_token="$(
  az account get-access-token \
    --resource "$ACS_TOKEN_RESOURCE" \
    --query accessToken \
    --output tsv
)"
[[ -n "$access_token" ]] || {
  echo "::error::Azure CLI did not return an Azure Communication Services access token."
  exit 1
}

send_status="$(
  curl \
    --silent \
    --show-error \
    --dump-header "$headers_file" \
    --output "$response_file" \
    --write-out '%{http_code}' \
    --request POST \
    --header "Authorization: Bearer $access_token" \
    --header 'Content-Type: application/json' \
    --data-binary "@$payload_file" \
    "$endpoint/emails:send?api-version=$ACS_EMAIL_API_VERSION"
)"
if [[ "$send_status" != '202' ]]; then
  echo "::error::Azure Communication Services rejected the email request (HTTP $send_status)."
  cat "$response_file"
  exit 1
fi

operation_url="$(
  awk '
    BEGIN { IGNORECASE = 1 }
    /^Operation-Location:/ {
      sub(/^[^:]+:[[:space:]]*/, "")
      sub(/\r$/, "")
      print
      exit
    }
  ' "$headers_file"
)"
if [[ -z "$operation_url" ]]; then
  operation_id="$(
    awk '
      BEGIN { IGNORECASE = 1 }
      /^Operation-Id:/ {
        sub(/^[^:]+:[[:space:]]*/, "")
        sub(/\r$/, "")
        print
        exit
      }
    ' "$headers_file"
  )"
  [[ -n "$operation_id" ]] || {
    echo "::error::Azure Communication Services did not return an operation location."
    exit 1
  }
  operation_url="$endpoint/emails/operations/$operation_id?api-version=$ACS_EMAIL_API_VERSION"
elif [[ "$operation_url" == /* ]]; then
  operation_url="$endpoint$operation_url"
fi

deadline=$((SECONDS + POLL_TIMEOUT_SECONDS))
while (( SECONDS < deadline )); do
  sleep 2
  poll_status="$(
    curl \
      --silent \
      --show-error \
      --output "$response_file" \
      --write-out '%{http_code}' \
      --header "Authorization: Bearer $access_token" \
      "$operation_url"
  )"
  if [[ "$poll_status" != '200' ]]; then
    echo "::error::Azure Communication Services email status check failed (HTTP $poll_status)."
    cat "$response_file"
    exit 1
  fi

  delivery_status="$(jq --raw-output '.status // empty' "$response_file")"
  case "$delivery_status" in
    Succeeded)
      echo "Azure Communication Services accepted the email notification."
      exit 0
      ;;
    Failed|Canceled|Cancelled)
      echo "::error::Azure Communication Services email delivery ended with status $delivery_status."
      jq '.error // .' "$response_file"
      exit 1
      ;;
    NotStarted|Running)
      ;;
    *)
      echo "::error::Azure Communication Services returned an unexpected email status: ${delivery_status:-missing}."
      cat "$response_file"
      exit 1
      ;;
  esac
done

echo "::error::Azure Communication Services email delivery did not finish within $POLL_TIMEOUT_SECONDS seconds."
exit 1
