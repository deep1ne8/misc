#!/bin/bash

# Elevate to root if not already
if [[ $EUID -ne 0 ]]; then
   echo "Please enter your password to continue installing AnyDesk..."
   sudo "$0" "$@"
   exit $?
fi

echo "Downloading AnyDesk..."
curl -L -o /tmp/anydesk.dmg https://download.anydesk.com/anydesk.dmg

echo "Mounting AnyDesk DMG..."
hdiutil attach /tmp/anydesk.dmg -nobrowse -quiet

echo "Copying AnyDesk to Applications..."
cp -R "/Volumes/AnyDesk/AnyDesk.app" /Applications/

echo "Unmounting DMG..."
hdiutil detach "/Volumes/AnyDesk" -quiet

echo "Cleaning up..."
rm /tmp/anydesk.dmg

# Launch AnyDesk
echo "Launching AnyDesk to initiate permissions prompt..."
open -a "/Applications/AnyDesk.app"

sleep 3

echo "Opening System Preferences for permission configuration..."

# Open Screen Recording permissions
open "x-apple.systempreferences:com.apple.preference.security?Privacy_ScreenCapture"
sleep 2

# Open Accessibility permissions
open "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility"
sleep 2

# Open Full Disk Access permissions
open "x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles"
sleep 2

echo "Installation complete. Please enable permissions for AnyDesk in each of the System Preference panes just opened."

exit 0
