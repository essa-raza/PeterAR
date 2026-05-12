#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_PATH="$PROJECT_ROOT/.venv-macos"
ARTIFACTS_DIR="$PROJECT_ROOT/artifacts"
APP_NAME="Customer Notes Merger"
PYTHON_BIN="${PYTHON_BIN:-python3}"

"$PYTHON_BIN" -m venv "$VENV_PATH"
source "$VENV_PATH/bin/activate"

python -m pip install --upgrade pip
python -m pip install -r "$PROJECT_ROOT/requirements.txt"
python "$PROJECT_ROOT/scripts/make_icon.py"

rm -rf "$PROJECT_ROOT/build" "$PROJECT_ROOT/dist" "$ARTIFACTS_DIR"
mkdir -p "$ARTIFACTS_DIR"

pyinstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "$APP_NAME" \
  --icon "$PROJECT_ROOT/assets/customer_notes_merger.icns" \
  "$PROJECT_ROOT/app.py"

APP_PATH="$PROJECT_ROOT/dist/$APP_NAME.app"

if [[ -n "${APPLE_CERTIFICATE_P12_BASE64:-}" ]] && [[ -n "${APPLE_CERTIFICATE_PASSWORD:-}" ]] && [[ -n "${APPLE_SIGNING_IDENTITY:-}" ]] && [[ -n "${APPLE_KEYCHAIN_PASSWORD:-}" ]]; then
  CERT_PATH="$RUNNER_TEMP/certificate.p12"
  KEYCHAIN_PATH="$RUNNER_TEMP/build.keychain-db"
  echo "$APPLE_CERTIFICATE_P12_BASE64" | base64 --decode > "$CERT_PATH"

  security create-keychain -p "$APPLE_KEYCHAIN_PASSWORD" "$KEYCHAIN_PATH"
  security set-keychain-settings -lut 21600 "$KEYCHAIN_PATH"
  security unlock-keychain -p "$APPLE_KEYCHAIN_PASSWORD" "$KEYCHAIN_PATH"
  security import "$CERT_PATH" -P "$APPLE_CERTIFICATE_PASSWORD" -A -t cert -f pkcs12 -k "$KEYCHAIN_PATH"
  security list-keychain -d user -s "$KEYCHAIN_PATH"

  codesign --force --deep --options runtime --timestamp \
    --sign "$APPLE_SIGNING_IDENTITY" \
    "$APP_PATH"

  if [[ -n "${APPLE_ID:-}" ]] && [[ -n "${APPLE_TEAM_ID:-}" ]] && [[ -n "${APPLE_APP_SPECIFIC_PASSWORD:-}" ]]; then
    ditto -c -k --keepParent "$APP_PATH" "$ARTIFACTS_DIR/$APP_NAME-notarization.zip"
    xcrun notarytool submit "$ARTIFACTS_DIR/$APP_NAME-notarization.zip" \
      --apple-id "$APPLE_ID" \
      --team-id "$APPLE_TEAM_ID" \
      --password "$APPLE_APP_SPECIFIC_PASSWORD" \
      --wait
    xcrun stapler staple "$APP_PATH"
  fi
fi

ditto -c -k --keepParent "$APP_PATH" "$ARTIFACTS_DIR/${APP_NAME}.zip"
echo "Created artifact: $ARTIFACTS_DIR/${APP_NAME}.zip"
