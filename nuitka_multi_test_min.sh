#!/bin/bash

# Define the variable for the Nuitka distribution directory
NUITKA_DIST_DIR="nuitka_test_min_signed/arm"
APP_NAME="scripti"
APP_BUNDLE="$NUITKA_DIST_DIR/$APP_NAME.app"
DMG_NAME="scripti52.dmg"
DMG_OUTPUT_DIR=$NUITKA_DIST_DIR
ICON_FILE="/Users/jd/dev/consulting/python_script_parser/appicon.icns" 

# Ensure the output directories exist
mkdir -p $NUITKA_DIST_DIR
mkdir -p $DMG_OUTPUT_DIR

# Run Nuitka to create the standalone application bundle
#python3 -m nuitka  --standalone --macos-create-app-bundle  --follow-imports scripti.py --include-data-dir=examples=examples --include-data-dir=icons=icons --noinclude-data-file=tcl/opt0.4 --noinclude-data-file=tcl/http1.0 --include-module=pandas --enable-plugin=tk-inter --macos-disable-console --macos-app-icon=$ICON_FILE --output-dir=$NUITKA_DIST_DIR  --noinclude-data-file=tcl/opt0.4 --noinclude-data-file=tcl/http1.0 
python3 -m nuitka  --standalone --macos-create-app-bundle  --follow-imports test.py  --macos-disable-console --macos-app-icon=$ICON_FILE --output-dir=$NUITKA_DIST_DIR  --macos-sign-identity='Developer ID Application: WeAnd Ltd (3UCPV3W9SM)' --macos-sign-notarization  --macos-signed-app-name="uk.co.weand.scriptparser_test_min" 

#python3 -m nuitka --standalone --macos-create-app-bundle --onefile --follow-imports scripti.py --include-data-dir=examples=examples --include-data-dir=icons=icons --noinclude-data-file=tcl/opt0.4 --noinclude-data-file=tcl/http1.0 --include-module=pandas --enable-plugin=tk-inter --macos-disable-console --macos-app-icon=$ICON_FILE --output-dir=$NUITKA_DIST_DIR  --macos-sign-identity="Developer ID Application: WeAnd Ltd (3UCPV3W9SM)" --macos-sign-notarization --disable-console

# Copy Info.plist to the app bundle
#cp Info.plist $APP_BUNDLE/Contents

# Sign the application bundle
#codesign --deep --force --verbose --options runtime  --sign "Developer ID Application: WeAnd Ltd (3UCPV3W9SM)" $APP_BUNDLE --timestamp
#codesign --deep --force --verbose --options runtime --entitlements entitlements.plist --sign "Developer ID Application: WeAnd Ltd (3UCPV3W9SM)" $APP_BUNDLE --timestamp

# Create the DMG file
#hdiutil create -volname "Scripti" -srcfolder "$APP_BUNDLE" -ov -format UDZO -o "$DMG_OUTPUT_DIR/$DMG_NAME"

# Sign the DMG file
#codesign --deep --force --verbose --options runtime --entitlements entitlements.plist --sign "Developer ID Application: WeAnd Ltd (3UCPV3W9SM)" "$DMG_OUTPUT_DIR/$DMG_NAME" --timestamp
