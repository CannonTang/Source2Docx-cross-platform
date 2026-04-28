#!/bin/bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
PROJECT_FILE="$ROOT_DIR/Source2Docx.csproj"
APP_NAME="Source2Docx"
RID="${1:-osx-arm64}"
CONFIGURATION="${CONFIGURATION:-Release}"
APP_VERSION="${APP_VERSION:-0.1}"
PUBLISH_DIR="$ROOT_DIR/dist/publish/$RID"
APP_DIR="$ROOT_DIR/dist/$APP_NAME-$RID.app"
ZIP_PATH="$ROOT_DIR/dist/$APP_NAME-$RID.zip"
EXECUTABLE_PATH="$APP_DIR/Contents/MacOS/$APP_NAME"

mkdir -p "$ROOT_DIR/dist"

dotnet publish "$PROJECT_FILE" \
  -c "$CONFIGURATION" \
  -r "$RID" \
  --self-contained true \
  -o "$PUBLISH_DIR"

rm -rf "$APP_DIR" "$ZIP_PATH"
mkdir -p "$APP_DIR/Contents/MacOS" "$APP_DIR/Contents/Resources"

rsync -a "$PUBLISH_DIR/" "$APP_DIR/Contents/MacOS/"
chmod +x "$EXECUTABLE_PATH"

cat > "$APP_DIR/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleDevelopmentRegion</key>
    <string>en</string>
    <key>CFBundleDisplayName</key>
    <string>$APP_NAME</string>
    <key>CFBundleExecutable</key>
    <string>$APP_NAME</string>
    <key>CFBundleIdentifier</key>
    <string>com.tangzhi.source2docx</string>
    <key>CFBundleInfoDictionaryVersion</key>
    <string>6.0</string>
    <key>CFBundleName</key>
    <string>$APP_NAME</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>$APP_VERSION</string>
    <key>CFBundleVersion</key>
    <string>$APP_VERSION</string>
    <key>LSMinimumSystemVersion</key>
    <string>12.0</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
EOF

# Re-sign the assembled bundle after files are copied into Contents/MacOS.
# Without this, the embedded apphost signature becomes inconsistent and
# Gatekeeper may report the app as damaged instead of simply unsigned.
codesign --force --deep --sign - "$APP_DIR"

ditto -c -k --keepParent "$APP_DIR" "$ZIP_PATH"

echo "App bundle: $APP_DIR"
echo "Zip package: $ZIP_PATH"
