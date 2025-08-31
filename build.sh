#!/bin/bash

# Voucher Processor Build Script
# This script helps build the application for different platforms

set -e  # Exit on any error

echo "=== Voucher Processor Build Script ==="

# Detect OS
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    OS="linux"
elif [[ "$OSTYPE" == "darwin"* ]]; then
    OS="macos"
elif [[ "$OSTYPE" == "msys" || "$OSTYPE" == "win32" ]]; then
    OS="windows"
else
    OS="unknown"
fi

echo "Detected OS: $OS"

# Function to install dependencies
install_deps() {
    echo "Installing Python dependencies..."
    pip install -r requirements.txt
    
    if [ "$OS" == "linux" ]; then
        echo "Checking for system dependencies..."
        # Check for LibreOffice
        if ! command -v libreoffice &> /dev/null; then
            echo "⚠️  LibreOffice not found. Please install it:"
            echo "   Ubuntu/Debian: sudo apt-get install libreoffice"
            echo "   Fedora: sudo dnf install libreoffice"
        else
            echo "✅ LibreOffice found"
        fi
    fi
}

# Function to build cx_Freeze executable
build_executable() {
    echo "Building standalone executable..."
    python setup.py build
    
    echo "✅ Executable built in the 'build/' directory"
}

# Function to build AppImage (Linux only)
build_appimage() {
    if [ "$OS" != "linux" ]; then
        echo "❌ AppImage can only be built on Linux"
        return 1
    fi
    
    echo "Building AppImage..."
    python build_appimage.py
}

# Function to show help
show_help() {
    echo "Usage: $0 [OPTIONS]"
    echo "Build script for Voucher Processor application"
    echo ""
    echo "Options:"
    echo "  deps          Install Python dependencies"
    echo "  exe           Build standalone executable"
    echo "  appimage      Build AppImage (Linux only)"
    echo "  all           Build everything"
    echo "  help          Show this help message"
    echo ""
    echo "Examples:"
    echo "  $0 deps       Install dependencies"
    echo "  $0 exe        Build executable"
    echo "  $0 all        Build everything"
}

# Main logic
case "${1:-help}" in
    deps)
        install_deps
        ;;
    exe)
        build_executable
        ;;
    appimage)
        build_appimage
        ;;
    all)
        install_deps
        build_executable
        if [ "$OS" == "linux" ]; then
            build_appimage
        fi
        ;;
    help|*)
        show_help
        ;;
esac

echo "=== Build process completed ==="