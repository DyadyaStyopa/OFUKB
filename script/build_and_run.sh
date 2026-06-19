#!/usr/bin/env bash
set -euo pipefail

MODE="${1:-run}"
APP_NAME="OFUKBMacGUI"
DISPLAY_NAME="OFUKB CBR PQ"
BUNDLE_ID="ru.ofukb.cbr-pq"
MIN_SYSTEM_VERSION="13.0"

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DIST_DIR="$ROOT_DIR/dist"
FINAL_APP_BUNDLE="$DIST_DIR/$DISPLAY_NAME.app"
STAGE_DIR="${TMPDIR:-/tmp}/ofukb_macos_app_build"
APP_BUNDLE="$STAGE_DIR/$DISPLAY_NAME.app"
APP_CONTENTS="$APP_BUNDLE/Contents"
APP_MACOS="$APP_CONTENTS/MacOS"
APP_RESOURCES="$APP_CONTENTS/Resources"
APP_BINARY="$APP_MACOS/$APP_NAME"
INFO_PLIST="$APP_CONTENTS/Info.plist"
BUILD_DIR="$DIST_DIR/build/$APP_NAME"
PYINSTALLER_DIST="$BUILD_DIR/pyinstaller"

pkill -x "$APP_NAME" >/dev/null 2>&1 || true

cd "$ROOT_DIR"
rm -rf "$FINAL_APP_BUNDLE" "$STAGE_DIR"
rm -rf "$BUILD_DIR"
mkdir -p "$APP_MACOS" "$APP_RESOURCES" "$BUILD_DIR"

SWIFT_FILES=()
while IFS= read -r -d '' file; do
  SWIFT_FILES+=("$file")
done < <(find "$ROOT_DIR/Sources/OFUKBMacGUI" -name '*.swift' -print0 | sort -z)

xcrun swiftc -parse-as-library -O -target arm64-apple-macosx"$MIN_SYSTEM_VERSION" "${SWIFT_FILES[@]}" -o "$BUILD_DIR/$APP_NAME-arm64"
xcrun swiftc -parse-as-library -O -target x86_64-apple-macosx"$MIN_SYSTEM_VERSION" "${SWIFT_FILES[@]}" -o "$BUILD_DIR/$APP_NAME-x86_64"
xcrun lipo -create "$BUILD_DIR/$APP_NAME-arm64" "$BUILD_DIR/$APP_NAME-x86_64" -output "$APP_BINARY"
chmod +x "$APP_BINARY"
cp "$ROOT_DIR/OFUKB_CBR_PQ_alt_parser.py" "$APP_RESOURCES/OFUKB_CBR_PQ_alt_parser.py"
cp "$ROOT_DIR/cbr_sqlite_export.py" "$APP_RESOURCES/cbr_sqlite_export.py"
cp "$ROOT_DIR/ofukb_cli.py" "$APP_RESOURCES/ofukb_cli.py"
cp "$ROOT_DIR/assets/app_icon.icns" "$APP_RESOURCES/app_icon.icns"
cp "$ROOT_DIR/assets/app_icon.png" "$APP_RESOURCES/app_icon.png"
cp "$ROOT_DIR/requirements.txt" "$APP_RESOURCES/requirements.txt"

python3 -m PyInstaller \
  --noconfirm \
  --clean \
  --onefile \
  --name ofukb_cli \
  --exclude-module matplotlib \
  --exclude-module scipy \
  --exclude-module IPython \
  --exclude-module jedi \
  --exclude-module pygments \
  --exclude-module PIL \
  --exclude-module tkinter \
  --distpath "$PYINSTALLER_DIST" \
  --workpath "$BUILD_DIR/pyinstaller-build" \
  --specpath "$BUILD_DIR" \
  "$ROOT_DIR/ofukb_cli.py"
cp "$PYINSTALLER_DIST/ofukb_cli" "$APP_RESOURCES/ofukb_cli"
chmod +x "$APP_RESOURCES/ofukb_cli"

cat >"$INFO_PLIST" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleExecutable</key>
  <string>$APP_NAME</string>
  <key>CFBundleIdentifier</key>
  <string>$BUNDLE_ID</string>
  <key>CFBundleName</key>
  <string>$DISPLAY_NAME</string>
  <key>CFBundleDisplayName</key>
  <string>$DISPLAY_NAME</string>
  <key>CFBundleIconFile</key>
  <string>app_icon</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>LSMinimumSystemVersion</key>
  <string>$MIN_SYSTEM_VERSION</string>
  <key>NSHumanReadableCopyright</key>
  <string>OFUKB</string>
  <key>NSPrincipalClass</key>
  <string>NSApplication</string>
</dict>
</plist>
PLIST

/usr/bin/xattr -cr "$APP_BUNDLE"
/usr/bin/xattr -c "$APP_BUNDLE"
while IFS= read -r -d '' path; do
  /usr/bin/xattr -d com.apple.FinderInfo "$path" >/dev/null 2>&1 || true
  /usr/bin/xattr -d com.apple.ResourceFork "$path" >/dev/null 2>&1 || true
  /usr/bin/xattr -d com.apple.macl "$path" >/dev/null 2>&1 || true
  /usr/bin/xattr -d 'com.apple.fileprovider.fpfs#P' "$path" >/dev/null 2>&1 || true
done < <(find "$APP_BUNDLE" -print0)
/usr/bin/codesign --force --deep --sign - "$APP_BUNDLE"
/usr/bin/codesign --verify --deep --strict --verbose=2 "$APP_BUNDLE"

/usr/bin/ditto --norsrc "$APP_BUNDLE" "$FINAL_APP_BUNDLE"
APP_BUNDLE="$FINAL_APP_BUNDLE"
/usr/bin/xattr -cr "$APP_BUNDLE" >/dev/null 2>&1 || true
/usr/bin/xattr -c "$APP_BUNDLE" >/dev/null 2>&1 || true
while IFS= read -r -d '' path; do
  /usr/bin/xattr -d com.apple.FinderInfo "$path" >/dev/null 2>&1 || true
  /usr/bin/xattr -d com.apple.ResourceFork "$path" >/dev/null 2>&1 || true
  /usr/bin/xattr -d com.apple.macl "$path" >/dev/null 2>&1 || true
  /usr/bin/xattr -d 'com.apple.fileprovider.fpfs#P' "$path" >/dev/null 2>&1 || true
done < <(find "$APP_BUNDLE" -print0)
/usr/bin/codesign --verify --deep --strict --verbose=2 "$APP_BUNDLE"

open_app() {
  /usr/bin/open -n "$APP_BUNDLE"
}

case "$MODE" in
  --build-only|build)
    ;;
  run)
    open_app
    ;;
  --debug|debug)
    lldb -- "$APP_BINARY"
    ;;
  --logs|logs)
    open_app
    /usr/bin/log stream --info --style compact --predicate "process == \"$APP_NAME\""
    ;;
  --telemetry|telemetry)
    open_app
    /usr/bin/log stream --info --style compact --predicate "subsystem == \"$BUNDLE_ID\""
    ;;
  --verify|verify)
    open_app
    sleep 1
    pgrep -x "$APP_NAME" >/dev/null
    ;;
  *)
    echo "usage: $0 [run|--build-only|--debug|--logs|--telemetry|--verify]" >&2
    exit 2
    ;;
esac
