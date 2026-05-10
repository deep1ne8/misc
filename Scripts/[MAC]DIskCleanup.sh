#!/usr/bin/env bash
# cleanup_mac.sh — Comprehensive Disk Cleanup for macOS 14.3+
# Author: Earl "deep1ne" Daniels
# Date: 2025-07-07
# Usage: sudo ./cleanup_mac.sh [--dry-run] [--verbose] [--help]

set -euo pipefail

# Configuration
readonly OS_MIN_VERSION="14.3"
readonly LOGFILE="/var/log/cleanup_mac_$(date +%Y%m%d_%H%M%S).log"
readonly SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Global flags
DRY_RUN=false
VERBOSE=false
FREED_SPACE=0

# Color codes for output
readonly RED='\033[0;31m'
readonly GREEN='\033[0;32m'
readonly YELLOW='\033[1;33m'
readonly BLUE='\033[0;34m'
readonly NC='\033[0m' # No Color

# Functions
log() {
    local level="$1"; shift
    local message="$*"
    local timestamp="$(date '+%Y-%m-%d %H:%M:%S')"
    local color=""
    
    case "$level" in
        "ERROR") color="$RED" ;;
        "WARN")  color="$YELLOW" ;;
        "INFO")  color="$GREEN" ;;
        "DEBUG") color="$BLUE" ;;
        "DRY")   color="$YELLOW" ;;
        *)       color="$NC" ;;
    esac
    
    printf "%b[%s] %s: %s%b\n" "$color" "$timestamp" "$level" "$message" "$NC" | tee -a "$LOGFILE"
}

error_exit() {
    log "ERROR" "$1"
    exit "${2:-1}"
}

require_root() {
    if [[ $EUID -ne 0 ]]; then
        error_exit "This script must be run as root. Use: sudo $0"
    fi
}

check_os_version() {
    local current_version
    current_version=$(sw_vers -productVersion)
    
    # Convert version strings to comparable format
    local current_major current_minor
    IFS='.' read -r current_major current_minor _ <<< "$current_version"
    
    local min_major min_minor
    IFS='.' read -r min_major min_minor _ <<< "$OS_MIN_VERSION"
    
    if [[ $current_major -lt $min_major ]] || 
       [[ $current_major -eq $min_major && $current_minor -lt $min_minor ]]; then
        error_exit "macOS $current_version detected. Minimum required: $OS_MIN_VERSION"
    fi
    
    log "INFO" "Running on macOS $current_version (Compatible)"
}

parse_args() {
    while (( $# )); do
        case $1 in
            --dry-run)
                DRY_RUN=true
                log "INFO" "Dry-run mode enabled: no files will be deleted"
                ;;
            --verbose)
                VERBOSE=true
                set -x
                ;;
            -h|--help)
                cat <<EOF
macOS Cleanup Script - Comprehensive disk cleanup utility

Usage: sudo $0 [OPTIONS]

OPTIONS:
    --dry-run    Show actions without deleting anything
    --verbose    Print detailed command execution
    -h, --help   Show this help message

DESCRIPTION:
    This script performs comprehensive cleanup of:
    - System and user caches
    - Log files
    - Trash bins
    - Homebrew cache
    - Docker system data
    - Xcode derived data
    - npm cache
    - Various temporary files

REQUIREMENTS:
    - macOS $OS_MIN_VERSION or later
    - Must be run as root (sudo)

EXAMPLES:
    sudo $0                    # Run full cleanup
    sudo $0 --dry-run          # Preview what would be cleaned
    sudo $0 --verbose          # Run with detailed output

EOF
                exit 0
                ;;
            *)
                log "WARN" "Unknown argument: $1 (use --help for usage)"
                ;;
        esac
        shift
    done
}

get_directory_size() {
    local dir="$1"
    if [[ -d "$dir" ]]; then
        du -sk "$dir" 2>/dev/null | cut -f1 || echo "0"
    else
        echo "0"
    fi
}

run_cmd() {
    local cmd="$*"
    if $DRY_RUN; then
        log "DRY" "$cmd"
        return 0
    else
        if $VERBOSE; then
            log "DEBUG" "Executing: $cmd"
        fi
        eval "$cmd"
    fi
}

clean_directory() {
    local dir="$1"
    local description="${2:-$dir}"
    
    if [[ ! -d "$dir" ]]; then
        log "WARN" "Directory not found: $dir"
        return 1
    fi
    
    # Check if directory is empty
    if [[ -z "$(ls -A "$dir" 2>/dev/null)" ]]; then
        log "INFO" "$description is already empty"
        return 0
    fi
    
    # Get size before cleanup
    local size_before
    size_before=$(get_directory_size "$dir")
    
    log "INFO" "Cleaning $description ($(numfmt --to=iec-i --suffix=B --padding=7 $((size_before * 1024))))"
    
    # Use find with -delete for safer deletion
    if $DRY_RUN; then
        local file_count
        file_count=$(find "$dir" -mindepth 1 -maxdepth 1 2>/dev/null | wc -l)
        log "DRY" "Would delete $file_count items from $dir"
    else
        # Create a temporary marker to preserve the directory
        local temp_marker="$dir/.cleanup_preserve_$(date +%s)"
        touch "$temp_marker" 2>/dev/null || true
        
        # Delete contents but preserve directory structure
        find "$dir" -mindepth 1 -maxdepth 1 ! -name "$(basename "$temp_marker")" -exec rm -rf {} + 2>/dev/null || true
        
        # Remove marker
        rm -f "$temp_marker" 2>/dev/null || true
        
        # Calculate freed space
        local size_after
        size_after=$(get_directory_size "$dir")
        local freed=$((size_before - size_after))
        FREED_SPACE=$((FREED_SPACE + freed))
        
        if [[ $freed -gt 0 ]]; then
            log "INFO" "Freed $(numfmt --to=iec-i --suffix=B --padding=7 $((freed * 1024))) from $description"
        fi
    fi
}

clean_caches() {
    log "INFO" "=== Clearing System & User Caches ==="
    
    # System caches
    clean_directory "/Library/Caches" "System Library Caches"
    clean_directory "/System/Library/Caches" "System Caches"
    
    # User caches for all users
    for user_home in /Users/*; do
        [[ -d "$user_home" ]] || continue
        local username
        username=$(basename "$user_home")
        [[ "$username" != "Shared" ]] || continue
        
        clean_directory "$user_home/Library/Caches" "$username's User Caches"
        clean_directory "$user_home/Library/Application Support/CrashReporter" "$username's Crash Reports"
        clean_directory "$user_home/Library/Saved Application State" "$username's Saved App States"
    done
    
    # Additional system cache locations
    clean_directory "/private/var/folders" "Temporary Folders"
    clean_directory "/private/tmp" "System Temporary Files"
}

clean_logs() {
    log "INFO" "=== Clearing System & User Logs ==="
    
    # Rotate system logs first
    run_cmd "log collect --output /dev/null 2>/dev/null || true"
    
    # Clean log directories
    clean_directory "/private/var/log" "System Logs"
    clean_directory "/Library/Logs" "Library Logs"
    
    # User logs
    for user_home in /Users/*; do
        [[ -d "$user_home" ]] || continue
        local username
        username=$(basename "$user_home")
        [[ "$username" != "Shared" ]] || continue
        
        clean_directory "$user_home/Library/Logs" "$username's User Logs"
    done
}

clean_trash() {
    log "INFO" "=== Emptying Trash ==="
    
    # Empty trash for all users
    for user_home in /Users/*; do
        [[ -d "$user_home" ]] || continue
        local username
        username=$(basename "$user_home")
        [[ "$username" != "Shared" ]] || continue
        
        clean_directory "$user_home/.Trash" "$username's Trash"
    done
    
    # Clean mounted volume trashes
    for vol in /Volumes/*; do
        [[ -d "$vol" ]] || continue
        [[ -d "$vol/.Trashes" ]] && clean_directory "$vol/.Trashes" "$(basename "$vol") Volume Trash"
    done
}

clean_homebrew() {
    if ! command -v brew &> /dev/null; then
        log "INFO" "Homebrew not installed; skipping"
        return 0
    fi
    
    log "INFO" "=== Running Homebrew cleanup ==="
    
    # Get brew cache size before cleanup
    local brew_cache_dir
    brew_cache_dir=$(brew --cache 2>/dev/null || echo "")
    
    if [[ -n "$brew_cache_dir" && -d "$brew_cache_dir" ]]; then
        local size_before
        size_before=$(get_directory_size "$brew_cache_dir")
        
        run_cmd "brew cleanup --prune=all"
        run_cmd "brew autoremove"
        
        if ! $DRY_RUN; then
            local size_after
            size_after=$(get_directory_size "$brew_cache_dir")
            local freed=$((size_before - size_after))
            FREED_SPACE=$((FREED_SPACE + freed))
            
            if [[ $freed -gt 0 ]]; then
                log "INFO" "Freed $(numfmt --to=iec-i --suffix=B --padding=7 $((freed * 1024))) from Homebrew"
            fi
        fi
    else
        run_cmd "brew cleanup --prune=all"
        run_cmd "brew autoremove"
    fi
}

clean_docker() {
    if ! command -v docker &> /dev/null; then
        log "INFO" "Docker not installed; skipping"
        return 0
    fi
    
    # Check if Docker daemon is running
    if ! docker info &> /dev/null; then
        log "WARN" "Docker daemon not running; skipping Docker cleanup"
        return 0
    fi
    
    log "INFO" "=== Pruning Docker system ==="
    
    # Get Docker system info before cleanup
    local docker_size_before
    docker_size_before=$(docker system df --format "{{.Size}}" 2>/dev/null | head -1 || echo "0B")
    
    run_cmd "docker system prune --all --volumes --force"
    
    if ! $DRY_RUN; then
        local docker_size_after
        docker_size_after=$(docker system df --format "{{.Size}}" 2>/dev/null | head -1 || echo "0B")
        log "INFO" "Docker cleanup completed (was: $docker_size_before, now: $docker_size_after)"
    fi
}

clean_xcode() {
    if ! xcode-select -p &> /dev/null; then
        log "INFO" "Xcode not detected; skipping"
        return 0
    fi
    
    log "INFO" "=== Cleaning Xcode data ==="
    
    # Clean for all users
    for user_home in /Users/*; do
        [[ -d "$user_home" ]] || continue
        local username
        username=$(basename "$user_home")
        [[ "$username" != "Shared" ]] || continue
        
        clean_directory "$user_home/Library/Developer/Xcode/DerivedData" "$username's Xcode DerivedData"
        clean_directory "$user_home/Library/Developer/Xcode/Archives" "$username's Xcode Archives"
        clean_directory "$user_home/Library/Developer/Xcode/iOS DeviceSupport" "$username's iOS DeviceSupport"
        clean_directory "$user_home/Library/Developer/CoreSimulator/Caches" "$username's Simulator Caches"
    done
}

clean_npm() {
    if ! command -v npm &> /dev/null; then
        log "INFO" "npm not installed; skipping"
        return 0
    fi
    
    log "INFO" "=== Cleaning npm cache ==="
    run_cmd "npm cache clean --force"
    
    # Clean yarn cache if available
    if command -v yarn &> /dev/null; then
        log "INFO" "=== Cleaning yarn cache ==="
        run_cmd "yarn cache clean --force"
    fi
}

clean_additional() {
    log "INFO" "=== Cleaning additional system files ==="
    
    # Clean QuickLook thumbnails
    clean_directory "/private/var/folders/*/*/C/com.apple.QuickLook.thumbnailcache" "QuickLook Thumbnails"
    
    # Clean Spotlight indexes (will be rebuilt automatically)
    run_cmd "mdutil -E / || true"
    
    # Clean font caches
    run_cmd "atsutil databases -remove || true"
    
    # Clean DNS cache
    run_cmd "dscacheutil -flushcache || true"
    
    # Clean system font cache
    run_cmd "fc-cache -f -v || true"
}

display_summary() {
    log "INFO" "=== Cleanup Summary ==="
    
    if $DRY_RUN; then
        log "INFO" "Dry-run completed - no files were actually deleted"
    else
        if [[ $FREED_SPACE -gt 0 ]]; then
            log "INFO" "Total space freed: $(numfmt --to=iec-i --suffix=B --padding=7 $((FREED_SPACE * 1024)))"
        else
            log "INFO" "No significant space was freed"
        fi
        
        # Show current disk usage
        local disk_usage
        disk_usage=$(df -h / | tail -1 | awk '{print $5 " used (" $3 " of " $2 ")"}')
        log "INFO" "Current disk usage: $disk_usage"
        
        log "INFO" "Log file saved to: $LOGFILE"
    fi
}

cleanup_on_exit() {
    local exit_code=$?
    if [[ $exit_code -ne 0 ]]; then
        log "ERROR" "Script exited with error code $exit_code"
    fi
    display_summary
    exit $exit_code
}

# Main execution
main() {
    # Set up error handling
    trap cleanup_on_exit EXIT
    
    # Initialize log file
    touch "$LOGFILE" 2>/dev/null || error_exit "Cannot create log file: $LOGFILE"
    
    log "INFO" "Starting macOS cleanup script"
    log "INFO" "Script version: $(date '+%Y-%m-%d')"
    log "INFO" "Running on: $(sw_vers -productName) $(sw_vers -productVersion)"
    
    # Parse arguments and validate environment
    parse_args "$@"
    require_root
    check_os_version
    
    # Execute cleanup functions
    clean_caches
    clean_logs
    clean_trash
    clean_homebrew
    clean_docker
    clean_xcode
    clean_npm
    clean_additional
    
    log "INFO" "Cleanup completed successfully"
}

# Execute main function with all arguments
main "$@"
