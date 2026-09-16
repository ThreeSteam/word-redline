#!/bin/bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
APP="$ROOT/build/Word Redline.app"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
SDK="$(xcrun --sdk macosx --show-sdk-path)"
for ARCH in arm64 x86_64; do
  xcrun swiftc -swift-version 5 -O -parse-as-library -sdk "$SDK" -target "$ARCH-apple-macosx13.0" \
    "$ROOT/Sources/Matcher.swift" "$ROOT/Sources/WordWorker.swift" "$ROOT/Sources/App.swift" \
    -o "$ROOT/build/WordRedline-$ARCH"
done
lipo -create "$ROOT/build/WordRedline-arm64" "$ROOT/build/WordRedline-x86_64" -output "$APP/Contents/MacOS/WordRedline"
cp "$ROOT/Resources/Info.plist" "$APP/Contents/Info.plist"
cp "$ROOT/Resources/Compare.applescript" "$APP/Contents/Resources/Compare.applescript"
codesign --force --deep --sign - "$APP"
"$APP/Contents/MacOS/WordRedline" --self-test
codesign --verify --deep --strict "$APP"
ditto -c -k --sequesterRsrc --keepParent "$APP" "$ROOT/build/Word-Redline-2.0-macOS-Universal.zip"
echo "Built: $ROOT/build/Word-Redline-2.0-macOS-Universal.zip"
