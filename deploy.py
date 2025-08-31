#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Deployment Script for Voucher Processor GUI Application

This script helps package the application for distribution.
"""

import os
import sys
import subprocess
import shutil
import zipfile
from pathlib import Path

def check_dependencies():
    """Check if required dependencies are installed"""
    print("Checking dependencies...")
    
    try:
        import pandas
        print("✓ pandas installed")
    except ImportError:
        print("✗ pandas not installed. Run: pip install pandas")
        return False
    
    try:
        import openpyxl
        print("✓ openpyxl installed")
    except ImportError:
        print("✗ openpyxl not installed. Run: pip install openpyxl")
        return False
    
    try:
        import tkinter
        print("✓ tkinter available")
    except ImportError:
        print("✗ tkinter not available. Please install tkinter")
        return False
    
    # Check LibreOffice
    try:
        result = subprocess.run(['libreoffice', '--version'], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            print("✓ LibreOffice installed")
        else:
            print("✗ LibreOffice not found. Please install LibreOffice")
            return False
    except FileNotFoundError:
        print("✗ LibreOffice not found. Please install LibreOffice")
        return False
    
    return True

def install_build_dependencies():
    """Install cx_Freeze for building executable"""
    print("\nInstalling build dependencies...")
    try:
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'cx_Freeze'], 
                      check=True)
        print("✓ cx_Freeze installed")
        return True
    except subprocess.CalledProcessError:
        print("✗ Failed to install cx_Freeze")
        return False

def build_executable():
    """Build standalone executable"""
    print("\nBuilding standalone executable...")
    
    try:
        # Clean previous build
        if os.path.exists('build'):
            shutil.rmtree('build')
            print("Cleaned previous build")
        
        # Build executable
        subprocess.run([sys.executable, 'setup.py', 'build'], check=True)
        print("✓ Executable built successfully")
        
        # List build contents
        build_dir = Path('build')
        if build_dir.exists():
            exe_dirs = list(build_dir.glob('exe.*'))
            if exe_dirs:
                exe_dir = exe_dirs[0]
                print(f"Build created in: {exe_dir}")
                
                # List main files
                files = list(exe_dir.glob('*'))
                print("Build contents:")
                for file in files[:10]:  # Show first 10 files
                    print(f"  - {file.name}")
                if len(files) > 10:
                    print(f"  ... and {len(files) - 10} more files")
                
                return str(exe_dir)
        
        return None
        
    except subprocess.CalledProcessError as e:
        print(f"✗ Build failed: {e}")
        return None

def create_distribution_package(build_dir):
    """Create distribution package"""
    print("\nCreating distribution package...")
    
    try:
        # Create distribution directory
        dist_name = "VoucherProcessor_Standalone"
        if os.path.exists(dist_name):
            shutil.rmtree(dist_name)
        
        os.makedirs(dist_name)
        
        # Copy build contents
        build_path = Path(build_dir)
        for item in build_path.iterdir():
            if item.is_file():
                shutil.copy2(item, dist_name)
            else:
                shutil.copytree(item, os.path.join(dist_name, item.name))
        
        # Copy additional files
        additional_files = ['README.md', 'requirements.txt']
        for file in additional_files:
            if os.path.exists(file):
                shutil.copy2(file, dist_name)
        
        # Create zip package
        zip_name = f"{dist_name}.zip"
        if os.path.exists(zip_name):
            os.remove(zip_name)
        
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(dist_name):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, dist_name)
                    zipf.write(file_path, arc_name)
        
        print(f"✓ Distribution package created: {zip_name}")
        print(f"✓ Distribution folder created: {dist_name}/")
        
        return zip_name
        
    except Exception as e:
        print(f"✗ Failed to create distribution package: {e}")
        return None

def create_python_package():
    """Create Python source package"""
    print("\nCreating Python source package...")
    
    try:
        package_name = "VoucherProcessor_Python"
        if os.path.exists(package_name):
            shutil.rmtree(package_name)
        
        os.makedirs(package_name)
        
        # Copy Python files
        python_files = [
            'voucher_gui.py',
            'converter.py', 
            'requirements.txt',
            'README.md',
            'setup.py'
        ]
        
        for file in python_files:
            if os.path.exists(file):
                shutil.copy2(file, package_name)
        
        # Create run script
        if sys.platform == "win32":
            run_script = "run.bat"
            content = "@echo off\npython voucher_gui.py\npause"
        else:
            run_script = "run.sh"
            content = "#!/bin/bash\npython3 voucher_gui.py"
        
        with open(os.path.join(package_name, run_script), 'w') as f:
            f.write(content)
        
        if not sys.platform == "win32":
            os.chmod(os.path.join(package_name, run_script), 0o755)
        
        # Create zip
        zip_name = f"{package_name}.zip"
        if os.path.exists(zip_name):
            os.remove(zip_name)
        
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(package_name):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, package_name)
                    zipf.write(file_path, arc_name)
        
        print(f"✓ Python package created: {zip_name}")
        print(f"✓ Python folder created: {package_name}/")
        
        return zip_name
        
    except Exception as e:
        print(f"✗ Failed to create Python package: {e}")
        return None

def main():
    """Main deployment function"""
    print("=== Voucher Processor Deployment Script ===\n")
    
    # Check dependencies
    if not check_dependencies():
        print("\n✗ Dependency check failed. Please install missing dependencies.")
        return False
    
    print("\n✓ All dependencies satisfied")
    
    # Ask user what to build
    print("\nDeployment Options:")
    print("1. Create Python source package (requires Python on target system)")
    print("2. Create standalone executable (includes Python runtime)")
    print("3. Create both packages")
    
    while True:
        choice = input("\nSelect option (1/2/3): ").strip()
        if choice in ['1', '2', '3']:
            break
        print("Invalid choice. Please enter 1, 2, or 3.")
    
    success = True
    
    # Create Python package
    if choice in ['1', '3']:
        python_package = create_python_package()
        if not python_package:
            success = False
    
    # Create executable
    if choice in ['2', '3']:
        if not install_build_dependencies():
            success = False
        else:
            build_dir = build_executable()
            if build_dir:
                exe_package = create_distribution_package(build_dir)
                if not exe_package:
                    success = False
            else:
                success = False
    
    if success:
        print("\n=== Deployment Completed Successfully! ===")
        print("\nDistribution packages are ready for deployment.")
        
        if choice in ['1', '3']:
            print(f"Python package: VoucherProcessor_Python.zip")
        if choice in ['2', '3']:
            print(f"Executable package: VoucherProcessor_Standalone.zip")
        
        print("\nNext steps:")
        print("1. Test the packages on target systems")
        print("2. Ensure LibreOffice is installed on target systems")
        print("3. Distribute the appropriate package for each system")
        
    else:
        print("\n✗ Deployment failed. Please check error messages above.")
    
    return success

if __name__ == "__main__":
    main()