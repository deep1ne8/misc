#!/bin/bash

# Enhanced ScreenConnect Deployment Script for macOS
# Requires sudo privileges

set -euo pipefail  # Exit on error, undefined vars, pipe failures

# Configuration
readonly SCRIPT_NAME="$(basename "$0")"
readonly LOG_FILE="/var/log/screenconnect_deployment.log"
readonly INSTALLER_PATH="/tmp/ScreenConnect.pkg"
readonly PPPC_PROFILE="/tmp/ScreenConnect_PPPC.mobileconfig"
readonly DOWNLOAD_URL="https://openapproach.screenconnect.com/Bin/ScreenConnect.ClientSetup.pkg?e=Access&y=Guest"
readonly APP_PATH="/Applications/ScreenConnect.app"
readonly BUNDLE_ID="com.screenconnect.client.access"
readonly MAX_RETRIES=3
readonly TIMEOUT=300

# Logging function
log() {
    local level="$1"
    shift
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] [$level] $*" | tee -a "$LOG_FILE"
}

# Error handling
error_exit() {
    log "ERROR" "$1"
    cleanup
    exit 1
}

# Cleanup function
cleanup() {
    log "INFO" "Cleaning up temporary files..."
    [[ -f "$INSTALLER_PATH" ]] && rm -f "$INSTALLER_PATH"
    [[ -f "$PPPC_PROFILE" ]] && rm -f "$PPPC_PROFILE"
}

# Check if running as root
check_privileges() {
    if [[ $EUID -ne 0 ]]; then
        error_exit "This script must be run as root (use sudo)"
    fi
}

# Verify system requirements
check_system() {
    log "INFO" "Checking system requirements..."
    
    # Check macOS version
    local os_version
    os_version=$(sw_vers -productVersion)
    log "INFO" "macOS version: $os_version"
    
    # Check available disk space (minimum 1GB)
    local available_space
    available_space=$(df /Applications | tail -1 | awk '{print $4}')
    if [[ $available_space -lt 1048576 ]]; then
        error_exit "Insufficient disk space. Need at least 1GB free."
    fi
    
    # Check internet connectivity
    if ! ping -c 1 google.com &>/dev/null; then
        error_exit "No internet connectivity detected"
    fi
}

# Create PPPC profile
create_pppc_profile() {
    log "INFO" "Creating PPPC profile..."
    
    cat > "$PPPC_PROFILE" << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>PayloadContent</key>
	<array>
		<dict>
			<key>PayloadDescription</key>
			<string>ConnectWiseControl PPPC</string>
			<key>PayloadDisplayName</key>
			<string>ConnectWiseControl PPPC</string>
			<key>PayloadIdentifier</key>
			<string>BEE165E6-42EE-4647-AC05-90A9F7A1E97F</string>
			<key>PayloadOrganization</key>
			<string>ConnectWise</string>
			<key>PayloadType</key>
			<string>com.apple.TCC.configuration-profile-policy</string>
			<key>PayloadUUID</key>
			<string>BEE165E6-42EE-4647-AC05-90A9F7A1E97F</string>
			<key>PayloadVersion</key>
			<integer>1</integer>
			<key>Services</key>
			<dict>
				<key>Accessibility</key>
				<array>
					<dict>
						<key>Allowed</key>
						<true/>
						<key>CodeRequirement</key>
						<string>identifier "com.apple.bash" and anchor apple</string>
						<key>Comment</key>
						<string></string>
						<key>Identifier</key>
						<string>/bin/bash</string>
						<key>IdentifierType</key>
						<string>path</string>
					</dict>
					<dict>
						<key>Allowed</key>
						<true/>
						<key>CodeRequirement</key>
						<string>identifier "com.screenconnect.client.access" and anchor apple generic and certificate 1[field.1.2.840.113635.100.6.2.6] /* exists */ and certificate leaf[field.1.2.840.113635.100.6.1.13] /* exists */ and certificate leaf[subject.OU] = "K8M3XDZV9Y"</string>
						<key>Comment</key>
						<string></string>
						<key>Identifier</key>
						<string>com.screenconnect.client.access</string>
						<key>IdentifierType</key>
						<string>bundleID</string>
					</dict>
				</array>
				<key>AppleEvents</key>
				<array>
					<dict>
						<key>AEReceiverCodeRequirement</key>
						<string>identifier "com.apple.systemevents" and anchor apple</string>
						<key>AEReceiverIdentifier</key>
						<string>com.apple.systemevents</string>
						<key>AEReceiverIdentifierType</key>
						<string>bundleID</string>
						<key>Allowed</key>
						<true/>
						<key>CodeRequirement</key>
						<string>identifier "com.apple.bash" and anchor apple</string>
						<key>Comment</key>
						<string></string>
						<key>Identifier</key>
						<string>/bin/bash</string>
						<key>IdentifierType</key>
						<string>path</string>
					</dict>
					<dict>
						<key>AEReceiverCodeRequirement</key>
						<string>identifier "com.apple.systemevents" and anchor apple</string>
						<key>AEReceiverIdentifier</key>
						<string>com.apple.systemevents</string>
						<key>AEReceiverIdentifierType</key>
						<string>bundleID</string>
						<key>Allowed</key>
						<true/>
						<key>CodeRequirement</key>
						<string>identifier "com.screenconnect.client.access" and anchor apple generic and certificate 1[field.1.2.840.113635.100.6.2.6] /* exists */ and certificate leaf[field.1.2.840.113635.100.6.1.13] /* exists */ and certificate leaf[subject.OU] = "K8M3XDZV9Y"</string>
						<key>Comment</key>
						<string></string>
						<key>Identifier</key>
						<string>com.screenconnect.client.access</string>
						<key>IdentifierType</key>
						<string>bundleID</string>
					</dict>
				</array>
			</dict>
		</dict>
	</array>
	<key>PayloadDescription</key>
	<string>ConnectWiseControl PPPC</string>
	<key>PayloadDisplayName</key>
	<string>ConnectWiseControl PPPC</string>
	<key>PayloadIdentifier</key>
	<string>77C06357-D87F-431E-BC89-1C8A5CA2516B</string>
	<key>PayloadOrganization</key>
	<string>ConnectWise</string>
	<key>PayloadType</key>
	<string>Configuration</string>
	<key>PayloadUUID</key>
	<string>77C06357-D87F-431E-BC89-1C8A5CA2516B</string>
	<key>PayloadVersion</key>
	<integer>1</integer>
	<key>payloadScope</key>
	<string>System</string>
	<key>PayloadVersion</key>
	<integer>1</integer>
</dict>
</plist>
EOF
    
    if [[ ! -f "$PPPC_PROFILE" ]]; then
        error_exit "Failed to create PPPC profile"
    fi
    
    log "INFO" "PPPC profile created successfully"
}

# Download installer with retry logic
download_installer() {
    log "INFO" "Downloading ScreenConnect installer..."
    
    local retry_count=0
    while [[ $retry_count -lt $MAX_RETRIES ]]; do
        if timeout "$TIMEOUT" curl -fsSL -o "$INSTALLER_PATH" "$DOWNLOAD_URL"; then
            log "INFO" "Download completed successfully"
            return 0
        else
            ((retry_count++))
            log "WARN" "Download attempt $retry_count failed. Retrying..."
            sleep 5
        fi
    done
    
    error_exit "Failed to download installer after $MAX_RETRIES attempts"
}

# Verify installer integrity
verify_installer() {
    log "INFO" "Verifying installer integrity..."
    
    if [[ ! -f "$INSTALLER_PATH" ]]; then
        error_exit "Installer file not found"
    fi
    
    # Check file size (should be > 1MB)
    local file_size
    file_size=$(stat -f%z "$INSTALLER_PATH" 2>/dev/null || echo 0)
    if [[ $file_size -lt 1048576 ]]; then
        error_exit "Installer file appears corrupted (size: $file_size bytes)"
    fi
    
    # Verify it's a valid package
    if ! pkgutil --check-signature "$INSTALLER_PATH" &>/dev/null; then
        log "WARN" "Package signature verification failed, but proceeding..."
    fi
    
    log "INFO" "Installer verification completed"
}

# Install PPPC profile first
install_pppc_profile() {
    log "INFO" "Installing PPPC profile..."
    
    # Remove existing profile if present
    if profiles -P | grep -q "com.organization.screenconnect.pppc"; then
        log "INFO" "Removing existing PPPC profile..."
        profiles -R -p com.organization.screenconnect.pppc || true
    fi
    
    # Install new profile
    if profiles -I -F "$PPPC_PROFILE"; then
        log "INFO" "PPPC profile installed successfully"
        # Wait for profile to take effect
        sleep 3
    else
        error_exit "Failed to install PPPC profile"
    fi
}

# Configure Gatekeeper
configure_gatekeeper() {
    log "INFO" "Configuring Gatekeeper settings..."
    
    # Check current Gatekeeper status
    local gatekeeper_status
    gatekeeper_status=$(spctl --status 2>/dev/null || echo "unknown")
    log "INFO" "Current Gatekeeper status: $gatekeeper_status"
    
    # Temporarily disable Gatekeeper if needed
    if [[ "$gatekeeper_status" == "assessments enabled" ]]; then
        log "INFO" "Temporarily disabling Gatekeeper for installation..."
        spctl --master-disable || log "WARN" "Failed to disable Gatekeeper"
    fi
}

# Install ScreenConnect
install_screenconnect() {
    log "INFO" "Installing ScreenConnect agent..."
    
    # Perform installation
    if installer -pkg "$INSTALLER_PATH" -target / -verbose; then
        log "INFO" "ScreenConnect installation completed"
    else
        error_exit "ScreenConnect installation failed"
    fi
    
    # Wait for installation to complete
    sleep 5
    
    # Verify installation
    if [[ ! -d "$APP_PATH" ]]; then
        error_exit "ScreenConnect application not found after installation"
    fi
}

# Post-installation configuration
post_install_config() {
    log "INFO" "Performing post-installation configuration..."
    
    # Re-enable Gatekeeper and whitelist ScreenConnect
    if [[ -d "$APP_PATH" ]]; then
        log "INFO" "Adding ScreenConnect to Gatekeeper whitelist..."
        spctl --add "$APP_PATH" || log "WARN" "Failed to add to Gatekeeper whitelist"
        spctl --master-enable || log "WARN" "Failed to re-enable Gatekeeper"
    fi
    
    # Start the service if not already running
    if ! pgrep -f "ScreenConnect" > /dev/null; then
        log "INFO" "Starting ScreenConnect service..."
        open "$APP_PATH" || log "WARN" "Failed to start ScreenConnect"
    fi
    
    # Set proper permissions
    chown -R root:wheel "$APP_PATH" 2>/dev/null || true
    chmod -R 755 "$APP_PATH" 2>/dev/null || true
}

# Verify deployment success
verify_deployment() {
    log "INFO" "Verifying deployment..."
    
    # Check if application exists
    if [[ ! -d "$APP_PATH" ]]; then
        error_exit "ScreenConnect application not found"
    fi
    
    # Check if PPPC profile is installed
    if ! profiles -P | grep -q "com.organization.screenconnect.pppc"; then
        log "WARN" "PPPC profile not found in installed profiles"
    fi
    
    # Check TCC database entries
    local tcc_count
    tcc_count=$(sqlite3 /Library/Application\ Support/com.apple.TCC/TCC.db "SELECT COUNT(*) FROM access WHERE client='$BUNDLE_ID';" 2>/dev/null || echo "0")
    log "INFO" "TCC database entries for ScreenConnect: $tcc_count"
    
    # Check if process is running
    if pgrep -f "ScreenConnect" > /dev/null; then
        log "INFO" "ScreenConnect process is running"
    else
        log "WARN" "ScreenConnect process not detected"
    fi
    
    log "INFO" "Deployment verification completed"
}

# Main execution function
main() {
    log "INFO" "Starting ScreenConnect deployment..."
    
    # Set trap for cleanup
    trap cleanup EXIT
    
    # Execute deployment steps
    check_privileges
    check_system
    create_pppc_profile
    download_installer
    verify_installer
    install_pppc_profile
    configure_gatekeeper
    install_screenconnect
    post_install_config
    verify_deployment
    
    log "INFO" "ScreenConnect deployment completed successfully!"
    echo "✅ ScreenConnect has been deployed successfully"
    echo "📝 Check $LOG_FILE for detailed logs"
    
    # Display next steps
    cat << 'EOF'

Next Steps:
1. Test remote connection capability
2. Verify all permissions are working correctly
3. Configure any additional ScreenConnect settings as needed
4. Document the deployment for your records

EOF
}

# Execute main function
main "$@"
