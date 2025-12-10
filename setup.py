# Create setup.py
"""
Office Optimizer Pro - Installation Script
==========================================
PROPRIETARY SOFTWARE - SHILEZI
Copyright (c) 2025 Shilezi (https://github.com/shilezi)
==========================================
"""

import subprocess
import sys
import os

def install_requirements():
    """Install required packages"""
    print("Installing requirements for Office Optimizer Pro...")
    print("=" * 50)
    
    requirements = [
        "customtkinter>=5.2.0",
        "Pillow>=10.0.0",
        "requests>=2.31.0",
        "imageio>=2.31.0"
    ]
    
    # Add pywin32 only on Windows
    if sys.platform == 'win32':
        requirements.append("pywin32>=306")
    
    for req in requirements:
        try:
            print(f"Installing {req}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", req])
        except subprocess.CalledProcessError as e:
            print(f"Failed to install {req}: {e}")
            return False
    
    print("\n" + "=" * 50)
    print("Installation complete!")
    print("\nTo run Office Optimizer Pro:")
    print("  python office_optimizer_pro.py")
    print("\nTo download FFmpeg (optional, for video compression):")
    print("  python download_minimal_ffmpeg.py")
    return True

def display_welcome():
    """Display welcome message and branding"""
    print("\n" + "=" * 60)
    print("OFFICE OPTIMIZER PRO v5.2")
    print("Created by: Shilezi (https://github.com/shilezi)")
    print("=" * 60)
    print("\n⚠️  PROPRIETARY SOFTWARE")
    print("   Unauthorized distribution is prohibited.")
    print("=" * 60 + "\n")

if __name__ == "__main__":
    display_welcome()
    
    response = input("Do you want to install required packages? (y/n): ")
    if response.lower() in ['y', 'yes']:
        install_requirements()
    else:
        print("\nYou can install requirements manually:")
        print("  pip install -r requirements.txt")
    
    input("\nPress Enter to exit...")
