# download_minimal_ffmpeg.py
import requests
import zipfile
import os
import sys
import platform
import shutil
from pathlib import Path

def download_minimal_ffmpeg():
    """Download minimal FFmpeg for Windows"""
    
    print("Downloading minimal FFmpeg...")
    print("This will reduce the size from 209MB to ~20MB\n")
    
    # URL for minimal FFmpeg build from BtbN
    # Updated to use the correct URL structure
    url = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
    
    # Alternative URLs if the above doesn't work
    # url = "https://github.com/BtbN/FFmpeg-Builds/releases/download/autobuild-2024-12-24-12-00/ffmpeg-master-latest-win64-gpl.zip"
    
    try:
        # Create temp directory
        temp_dir = "ffmpeg_temp"
        os.makedirs(temp_dir, exist_ok=True)
        zip_path = os.path.join(temp_dir, "ffmpeg.zip")
        
        # Download the zip
        print(f"Downloading from: {url}")
        response = requests.get(url, stream=True)
        response.raise_for_status()
        
        # Save with progress
        total_size = int(response.headers.get('content-length', 0))
        downloaded = 0
        
        with open(zip_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        percent = (downloaded / total_size) * 100
                        print(f"\r  Progress: {percent:.1f}%", end='')
        
        print("\n\nExtracting files...")
        
        # Extract the zip
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # List contents to understand structure
            file_list = zip_ref.namelist()
            
            # Find the bin directory
            bin_dirs = [f for f in file_list if 'bin/' in f and not f.endswith('/')]
            
            if not bin_dirs:
                print("  No bin directory found. Trying alternative structure...")
                # Try to extract everything and look for ffmpeg.exe
                zip_ref.extractall(temp_dir)
                
                # Search for ffmpeg.exe in extracted files
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        if file.lower() == 'ffmpeg.exe':
                            ffmpeg_path = os.path.join(root, file)
                            print(f"  Found ffmpeg.exe at: {ffmpeg_path}")
                            
                            # Create target directory
                            target_dir = "ffmpeg"
                            os.makedirs(target_dir, exist_ok=True)
                            
                            # Copy ffmpeg.exe and look for DLLs in the same directory
                            shutil.copy2(ffmpeg_path, os.path.join(target_dir, "ffmpeg.exe"))
                            
                            # Copy DLLs from the same directory
                            for dll_file in files:
                                if dll_file.lower().endswith('.dll'):
                                    shutil.copy2(os.path.join(root, dll_file), 
                                                os.path.join(target_dir, dll_file))
                                    print(f"  Copied: {dll_file}")
                            
                            break
            else:
                # Extract only bin directory contents
                print(f"  Found {len(bin_dirs)} files in bin directory")
                
                # Create target directory
                target_dir = "ffmpeg"
                os.makedirs(target_dir, exist_ok=True)
                
                # Extract only essential files from bin
                essential_patterns = [
                    'ffmpeg.exe',
                    'avcodec',
                    'avformat',
                    'avutil',
                    'swresample',
                    'swscale'
                ]
                
                extracted_count = 0
                for file in bin_dirs:
                    filename = os.path.basename(file)
                    if any(pattern in filename.lower() for pattern in essential_patterns):
                        # Extract to target directory
                        zip_ref.extract(file, temp_dir)
                        
                        # Move to final location
                        src_path = os.path.join(temp_dir, file)
                        dst_path = os.path.join(target_dir, filename)
                        
                        if os.path.exists(src_path):
                            shutil.move(src_path, dst_path)
                            print(f"  Extracted: {filename}")
                            extracted_count += 1
        
        # Clean up
        shutil.rmtree(temp_dir, ignore_errors=True)
        
        # Check if we got the files
        ffmpeg_exe = os.path.join("ffmpeg", "ffmpeg.exe")
        if os.path.exists(ffmpeg_exe):
            size_mb = os.path.getsize(ffmpeg_exe) / (1024 * 1024)
            dll_count = len([f for f in os.listdir("ffmpeg") if f.endswith('.dll')])
            
            print(f"\n✅ Minimal FFmpeg downloaded successfully!")
            print(f"Location: {ffmpeg_exe}")
            print(f"Size: {size_mb:.1f}MB")
            print(f"DLLs: {dll_count} files")
            print(f"Total folder size: {get_folder_size('ffmpeg'):.1f}MB")
            
            # Test if it works
            print("\nTesting FFmpeg...")
            try:
                result = subprocess.run([ffmpeg_exe, "-version"], 
                                      capture_output=True, text=True, timeout=5)
                if result.returncode == 0:
                    version_line = result.stdout.split('\n')[0]
                    print(f"  ✓ Working: {version_line}")
                else:
                    print("  ⚠️ FFmpeg returned error")
            except:
                print("  ⚠️ Could not test FFmpeg (but file exists)")
            
            return True
        else:
            print("\n❌ Failed to download FFmpeg.")
            print("Please download manually from: https://github.com/BtbN/FFmpeg-Builds/releases")
            print("Look for: ffmpeg-master-latest-win64-gpl.zip")
            return False
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("\nAlternative: Download FFmpeg manually:")
        print("1. Go to: https://github.com/BtbN/FFmpeg-Builds/releases")
        print("2. Download: ffmpeg-master-latest-win64-gpl.zip")
        print("3. Extract and copy ffmpeg.exe and .dll files to a folder named 'ffmpeg'")
        return False

def get_folder_size(folder_path):
    """Calculate folder size in MB"""
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            total_size += os.path.getsize(fp)
    return total_size / (1024 * 1024)

def download_simple_alternative():
    """Alternative simple download - just get ffmpeg.exe"""
    print("\nTrying alternative download...")
    
    # Direct download link for ffmpeg.exe only (static build, no DLLs needed)
    # This is from the official FFmpeg Windows builds
    url = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"
    
    try:
        import requests
        import zipfile
        import io
        
        print("Downloading from official FFmpeg builds...")
        response = requests.get(url)
        
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            # Look for ffmpeg.exe in the archive
            for name in z.namelist():
                if 'ffmpeg.exe' in name and 'bin' in name:
                    # Create ffmpeg directory
                    os.makedirs("ffmpeg_simple", exist_ok=True)
                    
                    # Extract ffmpeg.exe
                    z.extract(name, "ffmpeg_temp")
                    
                    # Move to final location
                    src = os.path.join("ffmpeg_temp", name)
                    dst = os.path.join("ffmpeg_simple", "ffmpeg.exe")
                    shutil.move(src, dst)
                    
                    print(f"✅ Downloaded: {dst}")
                    
                    # Clean up
                    shutil.rmtree("ffmpeg_temp", ignore_ok=True)
                    return True
        
        print("❌ Could not find ffmpeg.exe in the archive")
        return False
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    if platform.system() != "Windows":
        print("Warning: This script is for Windows only.")
        print("For other platforms, install FFmpeg via package manager.")
    else:
        # Try the main download first
        if not download_minimal_ffmpeg():
            print("\nTrying alternative...")
            download_simple_alternative()