"""
================================================================================
Office Optimizer Pro v5.4 - Installation Script (2025)
================================================================================
CREATED BY: SHILEZI (https://github.com/shilezi)
PROPRIETARY SOFTWARE - ALL RIGHTS RESERVED
================================================================================
"""

import subprocess
import sys
import os
import platform

def display_welcome():
    """Display welcome message and branding"""
    print("\n" + "="*70)
    print("⚡ OFFICE OPTIMIZER PRO v5.4 (2025)")
    print("   Created by: SHILEZI (https://github.com/shilezi)")
    print("="*70)
    print("\n⚠️  PROPRIETARY SOFTWARE")
    print("   Unauthorized distribution is prohibited.")
    print("="*70 + "\n")

def check_python_version():
    """Check Python version compatibility"""
    print("🔍 Checking Python version...")
    major, minor, *_ = sys.version_info
    
    if major < 3 or (major == 3 and minor < 8):
        print(f"❌ Python 3.8+ required (you have {major}.{minor})")
        print("   Download Python from: https://www.python.org/downloads/")
        return False
    
    print(f"✅ Python {major}.{minor} detected")
    return True

def check_platform():
    """Check if running on Windows (required for some features)"""
    print("\n🔍 Checking platform...")
    
    if platform.system() != "Windows":
        print("⚠️  Warning: Some features require Windows")
        print("   - PowerPoint structure optimization")
        print("   - Direct COM integration")
    else:
        print("✅ Windows detected")
    
    return True

def install_requirements():
    """Install required packages"""
    print("\n📦 Installing requirements for Office Optimizer Pro v5.4...")
    print("-" * 50)
    
    requirements = [
        "customtkinter>=5.2.0",
        "Pillow>=10.0.0",
        "requests>=2.31.0",
        "imageio>=2.31.0"
    ]
    
    # Add pywin32 only on Windows
    if platform.system() == 'Windows':
        requirements.append("pywin32>=306")
    
    success_count = 0
    fail_count = 0
    
    for req in requirements:
        try:
            print(f"Installing {req}...", end=' ', flush=True)
            subprocess.check_call([sys.executable, "-m", "pip", "install", req])
            print("✅")
            success_count += 1
        except subprocess.CalledProcessError as e:
            print(f"❌ (Error: {e})")
            fail_count += 1
        except Exception as e:
            print(f"❌ (Error: {e})")
            fail_count += 1
    
    print("\n" + "=" * 50)
    
    if fail_count == 0:
        print(f"✅ All {success_count} packages installed successfully!")
        return True
    else:
        print(f"⚠️  Installed {success_count} packages, {fail_count} failed")
        print("   You can install manually: pip install -r requirements.txt")
        return False

def setup_ffmpeg():
    """Ask user about FFmpeg setup"""
    print("\n🎬 FFmpeg Setup (Optional)")
    print("-" * 50)
    print("FFmpeg enables video and audio compression features.")
    print("Without FFmpeg, only image compression will be available.")
    
    response = input("\nDo you want to download and setup FFmpeg? (y/n): ").lower().strip()
    
    if response in ['y', 'yes']:
        print("\nSetting up FFmpeg...")
        
        # Check if download script exists
        if os.path.exists("download_minimal_ffmpeg.py"):
            try:
                import download_minimal_ffmpeg
                download_minimal_ffmpeg.download_minimal_ffmpeg()
            except:
                print("❌ Could not run FFmpeg setup script")
                print("   Run manually: python download_minimal_ffmpeg.py")
        else:
            print("❌ FFmpeg setup script not found")
            print("   Download it from the repository")
    
    else:
        print("⚠️  FFmpeg setup skipped")
        print("   You can run it later: python download_minimal_ffmpeg.py")

def create_shortcuts():
    """Create desktop shortcuts (Windows only)"""
    if platform.system() != "Windows":
        return
    
    response = input("\nCreate desktop shortcut? (y/n): ").lower().strip()
    
    if response in ['y', 'yes']:
        try:
            import winshell
            from win32com.client import Dispatch
            
            desktop = winshell.desktop()
            shortcut_path = os.path.join(desktop, "Office Optimizer Pro.lnk")
            
            target = sys.executable
            wDir = os.path.dirname(os.path.abspath(__file__))
            icon = os.path.join(wDir, "icon.ico") if os.path.exists("icon.ico") else ""
            
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = target
            shortcut.Arguments = f'"{os.path.join(wDir, "office_optimizer_pro.py")}"'
            shortcut.WorkingDirectory = wDir
            if icon:
                shortcut.IconLocation = icon
            shortcut.save()
            
            print("✅ Desktop shortcut created")
        except:
            print("⚠️  Could not create shortcut (requires admin privileges)")

def display_completion():
    """Display completion message"""
    print("\n" + "="*70)
    print("🎉 INSTALLATION COMPLETE!")
    print("="*70)
    
    print("\nTo run Office Optimizer Pro:")
    print("  python office_optimizer_pro.py")
    
    print("\nOptional setup:")
    print("  • FFmpeg for video compression: python download_minimal_ffmpeg.py")
    print("  • Desktop shortcut: Right-click → Send to → Desktop")
    
    print("\n" + "="*70)
    print("⚡ Thank you for using Office Optimizer Pro v5.4!")
    print("   Created by: SHILEZI (https://github.com/shilezi)")
    print("="*70 + "\n")

def main():
    """Main installation function"""
    display_welcome()
    
    # Check Python version
    if not check_python_version():
        input("\nPress Enter to exit...")
        return
    
    # Check platform
    check_platform()
    
    # Ask about installation
    response = input("\nProceed with installation? (y/n): ").lower().strip()
    
    if response not in ['y', 'yes']:
        print("\nInstallation cancelled.")
        print("You can install manually: pip install -r requirements.txt")
        input("\nPress Enter to exit...")
        return
    
    # Install requirements
    if not install_requirements():
        input("\nPress Enter to exit...")
        return
    
    # Setup FFmpeg
    setup_ffmpeg()
    
    # Create shortcuts (Windows only)
    create_shortcuts()
    
    # Display completion
    display_completion()
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
