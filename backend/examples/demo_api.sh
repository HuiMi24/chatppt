#!/usr/bin/env bash
set -euo pipefail

API=${1:-http://127.0.0.1:8000}
PPT_PATH=${2:-./sample.pptx}

echo "1) Parse PPT"
curl -s "$API/api/ppt?path=$PPT_PATH" | jq .

echo "2) Chat edit"
curl -s -X POST "$API/api/chat" \
  -H 'Content-Type: application/json' \
  -d "{\"path\":\"$PPT_PATH\",\"message\":\"把第1页标题改为通过API修改\"}" | jq .

echo "3) Theme apply"
curl -s -X POST "$API/api/theme/apply" \
  -H 'Content-Type: application/json' \
  -d "{\"path\":\"$PPT_PATH\"}" | jq .
