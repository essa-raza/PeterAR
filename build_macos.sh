#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_PATH="$PROJECT_ROOT/.venv-macos"
ARTIFACTS_DIR="$PROJECT_ROOT/artifacts"
APP_NAME="Customer Notes Merger"
PYTHON_BIN="${PYTHON_BIN:-python3}"
TARGET_ARCH="${TARGET_ARCH:-}"
DMG_SUFFIX="${DMG_SUFFIX:-macos}"
SAMPLE_CSV="$PROJECT_ROOT/Docs/Bestand uit Yuki - Openstaande transacties - one2three.csv"
SAMPLE_XLSX="$PROJECT_ROOT/Docs/Openstaande transacties  april (oud).xlsx"
SAMPLE_OUTPUT="$ARTIFACTS_DIR/sample-output-${DMG_SUFFIX}.xlsx"

"$PYTHON_BIN" -m venv "$VENV_PATH"
source "$VENV_PATH/bin/activate"

python -m pip install --upgrade pip
python -m pip install -r "$PROJECT_ROOT/requirements.txt"
python "$PROJECT_ROOT/scripts/make_icon.py"

rm -rf "$PROJECT_ROOT/build" "$PROJECT_ROOT/dist" "$ARTIFACTS_DIR"
mkdir -p "$ARTIFACTS_DIR"

PYINSTALLER_ARGS=(
  --noconfirm
  --clean
  --windowed
  --name "$APP_NAME"
  --icon "$PROJECT_ROOT/assets/customer_notes_merger.icns"
)

if [[ -n "$TARGET_ARCH" ]]; then
  PYINSTALLER_ARGS+=(--target-arch "$TARGET_ARCH")
fi

pyinstaller "${PYINSTALLER_ARGS[@]}" "$PROJECT_ROOT/app.py"

APP_PATH="$PROJECT_ROOT/dist/$APP_NAME.app"
APP_EXECUTABLE="$APP_PATH/Contents/MacOS/$APP_NAME"

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

"$APP_EXECUTABLE" --cli --csv "$SAMPLE_CSV" --excel "$SAMPLE_XLSX" --output "$SAMPLE_OUTPUT"

open "$APP_PATH"
sleep 8
screencapture -x "$ARTIFACTS_DIR/${APP_NAME}-${DMG_SUFFIX}-open.png"
pkill -f "$APP_NAME" || true

DMG_STAGING="$PROJECT_ROOT/dmg-staging"
rm -rf "$DMG_STAGING"
mkdir -p "$DMG_STAGING"
cp -R "$APP_PATH" "$DMG_STAGING/"
ln -s /Applications "$DMG_STAGING/Applications"

DMG_PATH="$ARTIFACTS_DIR/${APP_NAME}-${DMG_SUFFIX}.dmg"
hdiutil create -volname "$APP_NAME" -srcfolder "$DMG_STAGING" -ov -format UDZO "$DMG_PATH"

ditto -c -k --keepParent "$APP_PATH" "$ARTIFACTS_DIR/${APP_NAME}-${DMG_SUFFIX}.zip"
echo "Created artifacts:"
echo " - $DMG_PATH"
echo " - $SAMPLE_OUTPUT"
