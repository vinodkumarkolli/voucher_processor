#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AppImage build script for Voucher Processor GUI
This script creates AppImage packages for Fedora and Ubuntu Linux distributions.
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def check_prerequisites():
    """Check if required tools are installed"""
    print("🔍 Checking prerequisites...")
    
    required_tools = ['python3', 'pip', 'wget']
    missing_tools = []
    
    for tool in required_tools:
        if shutil.which(tool) is None:
            missing_tools.append(tool)
    
    if missing_tools:
        print(f"❌ Missing required tools: {', '.join(missing_tools)}")
        print("Please install them before proceeding.")
        return False
    
    print("✅ All required tools are available")
    return True

def install_dependencies():
    """Install Python dependencies to local directory"""
    print("\n📦 Installing Python dependencies...")
    
    try:
        # Create AppDir first if it doesn't exist
        appdir = Path('AppDir')
        appdir.mkdir(exist_ok=True)
        
        # Create local lib directory
        lib_dir = appdir / 'usr' / 'lib'
        lib_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"ℹ️  Installing dependencies to: {lib_dir.absolute()}")
        
        # Use system Python instead of virtual environment Python
        python_executable = shutil.which('python3')
        if not python_executable:
            python_executable = sys.executable
            
        print(f"ℹ️  Using Python executable: {python_executable}")
        
        # Install dependencies to local directory
        result = subprocess.run([
            python_executable, '-m', 'pip', 'install',
            '--target', str(lib_dir),
            '--no-user',
            '-r', 'requirements.txt'
        ], capture_output=True, text=True)
        
        print(f"ℹ️  Pip install return code: {result.returncode}")
        if result.stdout:
            print(f"ℹ️  Pip install stdout: {result.stdout}")
        if result.stderr:
            print(f"ℹ️  Pip install stderr: {result.stderr}")
        
        if result.returncode == 0:
            print("✅ Python dependencies installed to local directory")
            # List installed packages for debugging
            try:
                installed_packages = subprocess.run([
                    python_executable, '-m', 'pip', 'list', '--path', str(lib_dir)
                ], capture_output=True, text=True)
                if installed_packages.returncode == 0:
                    print("📋 Installed packages:")
                    for line in installed_packages.stdout.split('\n')[:10]:  # Show first 10 lines
                        if line.strip():
                            print(f"  {line}")
                else:
                    print(f"❌ Failed to list installed packages: {installed_packages.stderr}")
            except Exception as e:
                print(f"ℹ️  Could not list installed packages: {e}")
            
            # Verify lib directory contents
            print(f"✅ Dependencies installed successfully")
                
            return True
        else:
            print(f"❌ Failed to install Python dependencies: {result.stderr}")
            return False
    except Exception as e:
        print(f"❌ Failed to install Python dependencies: {e}")
        return False

def create_appimage():
    """Create AppImage package"""
    print("\n🔧 Creating AppImage...")
    
    try:
        # Check if AppDir already exists with dependencies
        appdir = Path('AppDir')
        lib_dir = appdir / 'usr' / 'lib'
        
        # Create AppDir structure if it doesn't exist
        if not appdir.exists():
            appdir.mkdir(exist_ok=True)
        
        # Create basic directory structure
        (appdir / 'usr/bin').mkdir(parents=True, exist_ok=True)
        
        # Copy main application files
        shutil.copy2('voucher_gui.py', appdir / 'usr/bin')
        shutil.copy2('converter.py', appdir / 'usr/bin')
        
        # Check if dependencies exist and create lib directory if needed
        if lib_dir.exists():
            print("ℹ️  Dependencies already installed in AppDir")
        else:
            print("⚠️  No dependencies found in AppDir")
        
        # Create AppRun script
        apprun_content = """#!/bin/bash
HERE="$(dirname "$(readlink -f "${0}")")"
export PYTHONPATH="${HERE}/usr/lib:${HERE}/usr/bin:${PYTHONPATH}"
python3 "${HERE}/usr/bin/voucher_gui.py" "$@"
"""
        
        with open(appdir / 'AppRun', 'w') as f:
            f.write(apprun_content)
        os.chmod(appdir / 'AppRun', 0o755)
        
        # Create desktop file
        desktop_content = """[Desktop Entry]
Name=Voucher Processor
Comment=Process voucher templates and CSV data into PDF output files
Exec=voucher_gui.py
Icon=voucher_processor
Terminal=false
Type=Application
Categories=Office;
"""
        
        with open(appdir / 'voucher_processor.desktop', 'w') as f:
            f.write(desktop_content)
        
        # Create a simple icon
        icon_content = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">
  <rect width="256" height="256" fill="#3498db"/>
  <text x="128" y="140" font-family="Arial" font-size="100" fill="white" text-anchor="middle">VP</text>
</svg>"""
        
        with open(appdir / 'voucher_processor.png', 'w') as f:
            f.write(icon_content)
        
        print("✅ AppDir structure created")
        
        # Download appimagetool if it doesn't exist
        if not os.path.exists('appimagetool'):
            print("📥 Downloading appimagetool...")
            subprocess.run([
                'wget', '-O', 'appimagetool', 
                'https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage'
            ], check=True)
            os.chmod('appimagetool', 0o755)
        
        # Create AppImage
        print("📦 Creating AppImage...")
        try:
            # Use shell command as a string instead of list with shell=True
            cmd = "ARCH=x86_64 ./appimagetool AppDir VoucherProcessor-x86_64.AppImage"
            result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True)
            
            # Verify the AppImage was created
            appimage_path = Path('VoucherProcessor-x86_64.AppImage')
            if appimage_path.exists():
                print("✅ AppImage created successfully")
                print(f"📁 AppImage location: {appimage_path.absolute()}")
                return True
            else:
                print("❌ AppImage file was not created")
                return False
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to create AppImage: {e}")
            if e.stderr:
                print(f"Error output: {e.stderr}")
            return False
        except Exception as e:
            print(f"❌ Unexpected error during AppImage creation: {e}")
            return False
        
    except Exception as e:
        print(f"❌ Failed to create AppImage: {e}")
        return False

def main():
    """Main build function"""
    print("=== Voucher Processor AppImage Build Script ===\n")
    
    # Check prerequisites
    if not check_prerequisites():
        return False
    
    # Install dependencies
    if not install_dependencies():
        return False
    
    # Create AppImage
    if create_appimage():
        print("\n🎉 AppImage created successfully!")
        print("📁 Your AppImage is ready: VoucherProcessor-x86_64.AppImage")
        return True
    
    print("\n❌ AppImage build failed")
    return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)