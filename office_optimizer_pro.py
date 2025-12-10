"""
================================================================================
Office Optimizer Pro v5.4
Professional tool to compress PowerPoint, Word, and Excel files
================================================================================
CREATED BY: SHILEZI (https://github.com/shilezi)
VERSION: 5.4.0 | RELEASE: 2025
================================================================================
PROPRIETARY SOFTWARE - ALL RIGHTS RESERVED
Copyright © 2025 Shilezi. Unauthorized distribution is prohibited.
================================================================================
"""

import zipfile
import os
import io
import shutil
import subprocess
import tempfile
import sys
import threading
import time
from datetime import datetime
from functools import lru_cache
import random

# GUI imports
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

# Image processing
from PIL import Image, ImageTk

# Try to import win32com for PowerPoint automation
try:
    import win32com.client
    HAS_COM = True
except ImportError:
    HAS_COM = False

# ============================================================================
# BRANDING PROTECTION
# ============================================================================

def verify_authenticity():
    """Verify this is an authentic Shilezi build"""
    try:
        # Check for Shilezi watermark in code
        watermark = "SHILEZI (https://github.com/shilezi)"
        with open(__file__, 'r', encoding='utf-8') as f:
            content = f.read(2000)  # Read first 2000 chars
            if watermark not in content:
                return False, "Unauthorized modification detected"
        
        # Check year
        if "2025" not in content:
            return False, "Invalid build year"
            
        return True, "Authentic Shilezi v5.4 Build (2025)"
    except:
        return False, "Integrity check failed"

# ============================================================================
# CONFIGURATION CONSTANTS
# ============================================================================

CONFIG = {
    "version": "5.4.0",
    "year": "2025",
    "author": "Shilezi",
    "repository": "https://github.com/shilezi/office-optimizer-pro",
    "max_file_size": 2 * 1024 * 1024 * 1024,  # 2GB
    "chunk_size": 10 * 1024 * 1024,
    "temp_backup_dir": os.path.join(tempfile.gettempdir(), "office_optimizer_backups"),
    "presets": {
        "Balanced (Recommended)": {"quality": 70, "max_width": 1920},
        "Strong (Smallest)": {"quality": 50, "max_width": 1280},
        "High Quality (Print)": {"quality": 90, "max_width": 3840},
        "Email (Light)": {"quality": 60, "max_width": 1024},
        "Archive (Lossless)": {"quality": 95, "max_width": 1920}
    }
}

# Display authenticity check on import
is_authentic, auth_message = verify_authenticity()
if not is_authentic:
    print(f"⚠️ WARNING: {auth_message}")
    print("Download the official version from: https://github.com/shilezi/office-optimizer-pro")

# ============================================================================
# CORE COMPRESSION ENGINE
# ============================================================================

class OfficeCompressor:
    """Main compression engine with enhanced features - Shilezi v5.4 (2025)"""
    
    def __init__(self, quality=70, max_width=1920, compress_video=False, 
                 png_smart_convert=False, enable_backup=True):
        self.quality = quality
        self.max_width = max_width
        self.compress_video_flag = compress_video
        self.png_smart_convert = png_smart_convert
        self.enable_backup = enable_backup
        self.chunk_size = CONFIG["chunk_size"]
        self.stats = {
            "files_processed": 0,
            "total_savings_bytes": 0,
            "total_original_size": 0,
            "processing_time": 0
        }
        
        # Locate FFmpeg
        self.ffmpeg_path = self._find_ffmpeg()
        
        # Create backup directory
        if enable_backup and not os.path.exists(CONFIG["temp_backup_dir"]):
            os.makedirs(CONFIG["temp_backup_dir"], exist_ok=True)
    
    def _find_ffmpeg(self):
        """Find FFmpeg executable in various locations"""
        # Check in current directory
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        # Check local directory first
        local_ffmpeg = os.path.join(base_path, "ffmpeg.exe")
        if os.path.exists(local_ffmpeg):
            return local_ffmpeg
        
        # Check in ffmpeg/ subdirectory
        ffmpeg_dir = os.path.join(base_path, "ffmpeg")
        if os.path.exists(ffmpeg_dir):
            local_ffmpeg = os.path.join(ffmpeg_dir, "ffmpeg.exe")
            if os.path.exists(local_ffmpeg):
                return local_ffmpeg
        
        # Check system PATH
        return shutil.which("ffmpeg")
    
    def check_ffmpeg(self):
        """Check if FFmpeg is available and working"""
        if not self.ffmpeg_path:
            return False, "FFmpeg not found. Video/Audio compression disabled."
        
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            result = subprocess.run(
                [self.ffmpeg_path, "-version"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                startupinfo=startupinfo,
                timeout=5
            )
            
            if result.returncode == 0:
                return True, f"Ready: {os.path.basename(self.ffmpeg_path)}"
            else:
                return False, "FFmpeg returned error code"
                
        except subprocess.TimeoutExpired:
            return False, "FFmpeg check timeout"
        except Exception as e:
            return False, f"FFmpeg Error: {str(e)}"
    
    def validate_file(self, filepath):
        """Validate if file is a valid Office file and within size limits"""
        # Check file exists
        if not os.path.exists(filepath):
            return False, "File does not exist"
        
        # Check file size
        size = os.path.getsize(filepath)
        if size > CONFIG["max_file_size"]:
            return False, f"File too large ({self._format_bytes(size)} > {self._format_bytes(CONFIG['max_file_size'])})"
        
        # Check if it's a valid Office file (zip with specific structure)
        if not filepath.lower().endswith(('.pptx', '.docx', '.xlsx')):
            return False, "Not a supported Office file (.pptx, .docx, .xlsx)"
        
        try:
            with zipfile.ZipFile(filepath, 'r') as zf:
                # Quick validation by checking for required Office file structure
                required = ['[Content_Types].xml']
                has_required = any(f in zf.namelist() for f in required)
                if not has_required:
                    return False, "Not a valid Office file (missing required structure)"
            return True, "Valid"
        except zipfile.BadZipFile:
            return False, "Not a valid ZIP/Office file"
        except Exception as e:
            return False, f"Validation error: {str(e)}"
    
    def create_backup(self, filepath):
        """Create timestamped backup of original file"""
        if not self.enable_backup:
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{os.path.basename(filepath)}.backup_{timestamp}"
        backup_path = os.path.join(CONFIG["temp_backup_dir"], backup_name)
        
        try:
            shutil.copy2(filepath, backup_path)
            return backup_path
        except Exception:
            return None
    
    def restore_backup(self, backup_path, original_path):
        """Restore file from backup"""
        if backup_path and os.path.exists(backup_path):
            try:
                shutil.copy2(backup_path, original_path)
                return True
            except Exception:
                return False
        return False
    
    def _has_actual_transparency(self, img):
        """Check if PNG actually uses transparency (not just has alpha channel)"""
        if img.mode in ('RGBA', 'LA', 'PA'):
            # For performance, sample random pixels instead of checking entire image
            sample_size = min(500, img.width * img.height)
            
            # Create a list of random positions to check
            positions = []
            for _ in range(sample_size):
                x = random.randint(0, img.width - 1)
                y = random.randint(0, img.height - 1)
                positions.append((x, y))
            
            # Check alpha channel
            if img.mode == 'RGBA':
                alpha = img.getchannel('A')
                for x, y in positions:
                    if alpha.getpixel((x, y)) < 250:  # Not fully opaque
                        return True
            elif img.mode in ('LA', 'PA'):
                # LA = Luminance + Alpha, PA = Palette + Alpha
                for x, y in positions:
                    pixel = img.getpixel((x, y))
                    if len(pixel) > 1 and pixel[1] < 250:
                        return True
        elif img.mode == 'P' and 'transparency' in img.info:
            # Paletted image with transparency
            return True
        
        return False
    
    def compress(self, input_path, output_path, progress_callback=None, log_callback=None):
        """Main compression method with enhanced error handling"""
        start_time = time.time()
        original_size = os.path.getsize(input_path)
        working_input = input_path
        temp_cleaned = None
        backup_path = None
        
        try:
            # Validate input file
            is_valid, msg = self.validate_file(input_path)
            if not is_valid:
                if log_callback:
                    log_callback(f"Validation failed: {msg}")
                return False
            
            # Create backup if enabled
            if self.enable_backup:
                backup_path = self.create_backup(input_path)
                if backup_path and log_callback:
                    log_callback(f"Backup created: {os.path.basename(backup_path)}")
            
            # Step 0: Structure Clean (PowerPoint only)
            if HAS_COM and input_path.lower().endswith('.pptx'):
                working_input = self._clean_presentation_structure(input_path, log_callback)
                if working_input != input_path:
                    temp_cleaned = working_input
            
            # Open input and output ZIP files
            with zipfile.ZipFile(working_input, 'r') as in_zip:
                with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as out_zip:
                    
                    file_list = in_zip.infolist()
                    total_files = len(file_list)
                    
                    for i, item in enumerate(file_list):
                        # Update progress
                        if progress_callback and i % 10 == 0:
                            progress_pct = (i / total_files) * 100
                            progress_callback(progress_pct)
                        
                        f_lower = item.filename.lower()
                        
                        # Process based on file type
                        if self._is_image(f_lower):
                            self._process_image(item, in_zip, out_zip, log_callback)
                        elif self.compress_video_flag and self._is_video(f_lower):
                            if self.ffmpeg_path:
                                if log_callback:
                                    log_callback(f"Video: {self._truncate_name(item.filename)}...")
                                self._process_video(item, in_zip, out_zip)
                            else:
                                self._copy_file(item, in_zip, out_zip)
                        elif self.compress_video_flag and self._is_audio(f_lower):
                            if self.ffmpeg_path:
                                if log_callback:
                                    log_callback(f"Audio: {self._truncate_name(item.filename)}...")
                                self._process_audio(item, in_zip, out_zip)
                            else:
                                self._copy_file(item, in_zip, out_zip)
                        else:
                            self._copy_file(item, in_zip, out_zip)
            
            # Calculate statistics
            compressed_size = os.path.getsize(output_path)
            self.stats["files_processed"] += 1
            self.stats["total_original_size"] += original_size
            self.stats["total_savings_bytes"] += (original_size - compressed_size)
            self.stats["processing_time"] += (time.time() - start_time)
            
            if log_callback:
                savings_pct = ((original_size - compressed_size) / original_size * 100) if original_size > 0 else 0
                log_callback(f"Complete: Saved {self._format_bytes(original_size - compressed_size)} ({savings_pct:.1f}%)")
            
            # Clean up temporary files
            if temp_cleaned and os.path.exists(temp_cleaned):
                try:
                    shutil.rmtree(os.path.dirname(temp_cleaned), ignore_errors=True)
                except:
                    pass
            
            return True
            
        except Exception as e:
            if log_callback:
                log_callback(f"Error: {str(e)}")
            
            # Try to restore from backup on error
            if backup_path and os.path.exists(backup_path):
                self.restore_backup(backup_path, input_path)
                if log_callback:
                    log_callback("Restored from backup due to error")
            
            return False
    
    def _clean_presentation_structure(self, input_path, log_callback=None):
        """Clean PowerPoint presentation structure (remove unused layouts)"""
        if not HAS_COM:
            return input_path
        
        try:
            temp_dir = tempfile.mkdtemp()
            temp_pptx = os.path.join(temp_dir, "clean_" + os.path.basename(input_path))
            shutil.copy2(input_path, temp_pptx)
            
            if log_callback:
                log_callback("Optimizing PowerPoint structure...")
            
            ppt_app = None
            presentation = None
            
            try:
                # Initialize PowerPoint
                ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                ppt_app.Visible = False
                ppt_app.DisplayAlerts = False
                
                # Open presentation
                abs_path = os.path.abspath(temp_pptx)
                presentation = ppt_app.Presentations.Open(abs_path, WithWindow=False)
                
                # Save cleaned copy
                cleaned_path = os.path.join(temp_dir, "cleaned_structure.pptx")
                presentation.SaveCopyAs(os.path.abspath(cleaned_path))
                
                return cleaned_path
                
            except Exception as e:
                if log_callback:
                    log_callback(f"PowerPoint optimization skipped: {str(e)}")
                return input_path
                    
            finally:
                # Clean up COM objects properly
                if presentation:
                    presentation.Close()
                if ppt_app:
                    ppt_app.Quit()
                
                # Force garbage collection
                import gc
                gc.collect()
                
        except Exception:
            return input_path
    
    def _process_image(self, zip_info, in_zip, out_zip, log_callback=None):
        """Process and compress image files"""
        try:
            img_data = in_zip.read(zip_info.filename)
            
            # Use BytesIO for in-memory processing
            with Image.open(io.BytesIO(img_data)) as img:
                original_mode = img.mode
                original_size = len(img_data)
                
                # Resize if needed
                if img.width > self.max_width or img.height > self.max_width:
                    img.thumbnail((self.max_width, self.max_width), Image.Resampling.LANCZOS)
                
                out_buffer = io.BytesIO()
                is_png = zip_info.filename.lower().endswith('.png')
                save_format = 'JPEG'
                
                # PNG handling with smart conversion
                if is_png:
                    save_format = 'PNG'  # Default
                    
                    if self.png_smart_convert:
                        # Check if PNG actually uses transparency
                        has_transparency = self._has_actual_transparency(img)
                        
                        if not has_transparency:
                            # Opaque PNG - convert to JPEG for better compression
                            save_format = 'JPEG'
                            img = img.convert('RGB')
                            if log_callback:
                                log_callback(f"  Converted PNG to JPEG: {os.path.basename(zip_info.filename)}")
                
                # Handle other formats
                if not is_png and img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Save with appropriate settings
                if save_format == 'JPEG':
                    img.save(out_buffer, format='JPEG', quality=self.quality, optimize=True)
                else:
                    # Optimize PNG (quantize if RGBA)
                    if img.mode == 'RGBA':
                        img = img.quantize(colors=256, method=2)
                    img.save(out_buffer, format='PNG', optimize=True)
                
                compressed_size = out_buffer.tell()
                
                # Only replace if we actually saved space
                if compressed_size < original_size:
                    out_zip.writestr(zip_info.filename, out_buffer.getvalue())
                    if log_callback:
                        savings = original_size - compressed_size
                        log_callback(f"  Compressed: {os.path.basename(zip_info.filename)} (-{self._format_bytes(savings)})")
                else:
                    # Keep original if compression didn't help
                    self._copy_file(zip_info, in_zip, out_zip)
                    
        except Exception as e:
            if log_callback:
                log_callback(f"  Image processing error: {str(e)}")
            self._copy_file(zip_info, in_zip, out_zip)
    
    def _process_video(self, zip_info, in_zip, out_zip):
        """Compress video files using FFmpeg"""
        temp_dir = tempfile.mkdtemp()
        original = os.path.join(temp_dir, "orig" + os.path.splitext(zip_info.filename)[1])
        compressed = os.path.join(temp_dir, "comp.mp4")
        
        try:
            # Extract original video
            with in_zip.open(zip_info) as src, open(original, 'wb') as dst:
                shutil.copyfileobj(src, dst)
            
            original_size = os.path.getsize(original)
            
            # Determine compression settings based on quality
            if self.quality <= 50:  # Strong compression
                crf_val = '12'
                audio_bitrate = '192k'
                scale_width = 1280
                fps_filter = ",fps=30"
                preset = 'fast'
            elif self.quality <= 70:  # Balanced
                crf_val = '4'
                audio_bitrate = '256k'
                scale_width = 1920
                fps_filter = ""
                preset = 'medium'
            else:  # High quality
                crf_val = '1'
                audio_bitrate = '320k'
                scale_width = 3840
                fps_filter = ""
                preset = 'slow'
            
            # Limit scale width
            if self.max_width < scale_width:
                scale_width = self.max_width
            
            # Build FFmpeg command
            cmd = [
                self.ffmpeg_path, '-y', '-i', original,
                '-vcodec', 'libx264', '-crf', crf_val,
                '-preset', preset,
                '-vf', f"scale='min({scale_width},iw)':-2{fps_filter}",
                '-ac', '2', '-b:a', audio_bitrate,
                '-movflags', '+faststart',
                compressed
            ]
            
            # Run FFmpeg
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            subprocess.run(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
                startupinfo=startupinfo,
                timeout=300  # 5 minute timeout
            )
            
            # Check if compression was beneficial
            if os.path.exists(compressed):
                compressed_size = os.path.getsize(compressed)
                if compressed_size < original_size * 0.95:  # At least 5% savings
                    with open(compressed, 'rb') as f:
                        out_zip.writestr(zip_info.filename, f.read())
                else:
                    self._copy_file(zip_info, in_zip, out_zip)
            else:
                self._copy_file(zip_info, in_zip, out_zip)
                
        except subprocess.TimeoutExpired:
            self._copy_file(zip_info, in_zip, out_zip)
        except Exception:
            self._copy_file(zip_info, in_zip, out_zip)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    def _process_audio(self, zip_info, in_zip, out_zip):
        """Compress audio files using FFmpeg"""
        temp_dir = tempfile.mkdtemp()
        original = os.path.join(temp_dir, "orig" + os.path.splitext(zip_info.filename)[1])
        compressed = os.path.join(temp_dir, "comp.mp3")
        
        try:
            # Extract original audio
            with in_zip.open(zip_info) as src, open(original, 'wb') as dst:
                shutil.copyfileobj(src, dst)
            
            # Build FFmpeg command based on file type
            cmd = []
            ext = zip_info.filename.lower()
            
            if ext.endswith('.wav'):
                # Convert WAV to MP3
                bitrate = '128k' if self.quality <= 50 else '192k'
                cmd = [self.ffmpeg_path, '-y', '-i', original, '-codec:a', 'libmp3lame',
                      '-b:a', bitrate, '-ac', '2', '-ar', '44100', compressed]
            else:
                # Re-encode other formats
                bitrate = '128k' if self.quality <= 50 else ('192k' if self.quality <= 70 else '256k')
                cmd = [self.ffmpeg_path, '-y', '-i', original, '-b:a', bitrate, compressed]
            
            # Run FFmpeg
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            subprocess.run(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
                startupinfo=startupinfo,
                timeout=60
            )
            
            # Replace if compressed version is smaller
            if os.path.exists(compressed) and os.path.getsize(compressed) < os.path.getsize(original):
                with open(compressed, 'rb') as f:
                    out_zip.writestr(zip_info.filename, f.read())
            else:
                self._copy_file(zip_info, in_zip, out_zip)
                
        except Exception:
            self._copy_file(zip_info, in_zip, out_zip)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    def _copy_file(self, zip_info, in_zip, out_zip):
        """Copy file without modification"""
        with in_zip.open(zip_info) as src, out_zip.open(zip_info, 'w') as dst:
            shutil.copyfileobj(src, dst, self.chunk_size)
    
    def get_statistics(self):
        """Get compression statistics"""
        if self.stats["total_original_size"] == 0:
            return {}
        
        savings_pct = (self.stats["total_savings_bytes"] / self.stats["total_original_size"] * 100)
        
        return {
            "files_processed": self.stats["files_processed"],
            "original_size": self._format_bytes(self.stats["total_original_size"]),
            "savings_bytes": self._format_bytes(self.stats["total_savings_bytes"]),
            "savings_percent": f"{savings_pct:.1f}%",
            "processing_time": f"{self.stats['processing_time']:.1f}s",
            "average_speed": self._format_bytes(self.stats["total_original_size"] / max(self.stats["processing_time"], 1)) + "/s"
        }
    
    def _is_image(self, filename):
        return 'media/' in filename and filename.endswith(('.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp'))
    
    def _is_video(self, filename):
        return 'media/' in filename and filename.endswith(('.mp4', '.m4v', '.mov', '.avi', '.wmv', '.mkv', '.flv', '.webm'))
    
    def _is_audio(self, filename):
        return 'media/' in filename and filename.endswith(('.wav', '.mp3', '.m4a', '.wma', '.ogg', '.flac'))
    
    def _truncate_name(self, text, limit=40):
        return text[:limit-3] + "..." if len(text) > limit else text
    
    def _format_bytes(self, size):
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"


# ============================================================================
# MODERN GUI APPLICATION
# ============================================================================

class OfficeOptimizerApp(ctk.CTk):
    """Modern GUI application for Office file optimization - Shilezi v5.4 (2025)"""
    
    def __init__(self):
        super().__init__()
        
        # Verify authenticity
        is_authentic, auth_message = verify_authenticity()
        if not is_authentic:
            messagebox.showerror("Unauthorized Build", 
                                f"⚠️ WARNING: {auth_message}\n\n"
                                f"Please download the official version from:\n"
                                f"{CONFIG['repository']}")
            sys.exit(1)
        
        # Configure window
        self.title(f"Office Optimizer Pro v{CONFIG['version']} (Shilezi © {CONFIG['year']})")
        self.geometry("1000x800")
        self.minsize(900, 700)
        
        # Set theme
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        
        # Application state
        self.files = []
        self.row_widgets = {}
        self.is_processing = False
        self.compression_stats = {}
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Create UI components
        self._create_header()
        self._create_file_area()
        self._create_settings_area()
        self._create_footer()
        self._create_status_bar()
        
        # Check system on load
        self.after(500, self._check_system)
    
    def _create_header(self):
        """Create application header with Shilezi branding"""
        self.header = ctk.CTkFrame(self, height=100, corner_radius=0, fg_color="#1a365d")
        self.header.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.header.grid_columnconfigure(0, weight=1)
        
        # Main title with Shilezi branding
        title_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        title_frame.pack(expand=True, fill="both", padx=30, pady=10)
        
        # Shilezi logo/text
        ctk.CTkLabel(
            title_frame,
            text="⚡ SHILEZI",
            font=("Segoe UI", 16, "bold"),
            text_color="#60a5fa"
        ).pack(side="left", pady=5)
        
        # Main title
        ctk.CTkLabel(
            title_frame,
            text="Office Optimizer Pro",
            font=("Segoe UI", 28, "bold"),
            text_color="white"
        ).pack(side="left", padx=(20, 0), pady=5)
        
        # Version and year
        version_frame = ctk.CTkFrame(title_frame, fg_color="transparent")
        version_frame.pack(side="left", padx=(15, 0), pady=5)
        
        ctk.CTkLabel(
            version_frame,
            text=f"v{CONFIG['version']}",
            font=("Segoe UI", 12, "bold"),
            text_color="#93c5fd"
        ).pack(side="top")
        
        ctk.CTkLabel(
            version_frame,
            text=f"© {CONFIG['year']}",
            font=("Segoe UI", 10),
            text_color="#bfdbfe"
        ).pack(side="top")
        
        # Quick actions
        action_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        action_frame.pack(fill="x", padx=30, pady=(0, 10))
        
        ctk.CTkButton(
            action_frame,
            text="📊 Statistics",
            width=100,
            height=30,
            command=self._show_statistics
        ).pack(side="right", padx=(5, 0))
        
        ctk.CTkButton(
            action_frame,
            text="⚙️ Settings",
            width=100,
            height=30,
            command=self._open_settings
        ).pack(side="right")
        
        # Demo watermark (non-removable)
        ctk.CTkLabel(
            action_frame,
            text="🚀 Shilezi Official Build",
            font=("Consolas", 9, "bold"),
            text_color="#fbbf24"
        ).pack(side="left")
    
    def _create_file_area(self):
        """Create file selection and queue area"""
        self.file_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.file_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        self.file_frame.grid_columnconfigure(0, weight=1)
        self.file_frame.grid_rowconfigure(1, weight=1)
        
        # Toolbar
        toolbar = ctk.CTkFrame(self.file_frame, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        ctk.CTkButton(
            toolbar,
            text="📁 Add Files",
            font=("Segoe UI", 13, "bold"),
            command=self._add_files,
            height=38,
            width=120
        ).pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(
            toolbar,
            text="📂 Add Folder",
            font=("Segoe UI", 13),
            command=self._add_folder,
            height=38,
            width=120
        ).pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(
            toolbar,
            text="🗑️ Clear All",
            fg_color="#334155",
            hover_color="#475569",
            command=self._clear_files,
            height=38,
            width=100
        ).pack(side="left")
        
        # File count label
        self.lbl_file_count = ctk.CTkLabel(
            toolbar,
            text="0 files selected",
            text_color="gray",
            font=("Segoe UI", 11)
        )
        self.lbl_file_count.pack(side="right", padx=10)
        
        # Total size label
        self.lbl_total_size = ctk.CTkLabel(
            toolbar,
            text="Total: 0 B",
            text_color="gray",
            font=("Segoe UI", 11)
        )
        self.lbl_total_size.pack(side="right", padx=10)
        
        # File list with scroll
        self.scroll_frame = ctk.CTkScrollableFrame(
            self.file_frame,
            label_text="File Queue",
            label_font=("Segoe UI", 12, "bold")
        )
        self.scroll_frame.grid(row=1, column=0, sticky="nsew")
    
    def _create_settings_area(self):
        """Create settings and controls area"""
        self.settings_frame = ctk.CTkFrame(self, fg_color="#1e293b", corner_radius=12)
        self.settings_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        self.settings_frame.grid_columnconfigure(1, weight=1)
        self.settings_frame.grid_columnconfigure(3, weight=1)
        
        # Profile selection
        ctk.CTkLabel(
            self.settings_frame,
            text="Optimization Profile",
            font=("Segoe UI", 12, "bold")
        ).grid(row=0, column=0, padx=20, pady=15, sticky="w")
        
        self.profile_var = ctk.StringVar(value="Balanced (Recommended)")
        self.combo_profile = ctk.CTkComboBox(
            self.settings_frame,
            values=list(CONFIG["presets"].keys()),
            variable=self.profile_var,
            command=self._on_profile_change,
            width=220,
            state="readonly"
        )
        self.combo_profile.grid(row=0, column=1, padx=10, sticky="w")
        
        # Save method
        ctk.CTkLabel(
            self.settings_frame,
            text="Save Method",
            font=("Segoe UI", 12, "bold")
        ).grid(row=0, column=2, padx=20, sticky="w")
        
        self.save_var = ctk.StringVar(value="Create Optimized Copy")
        self.combo_save = ctk.CTkComboBox(
            self.settings_frame,
            values=["Create Optimized Copy", "Replace Original (with backup)"],
            variable=self.save_var,
            width=220,
            state="readonly"
        )
        self.combo_save.grid(row=0, column=3, padx=10, sticky="w")
        
        # Options frame
        options_frame = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        options_frame.grid(row=1, column=0, columnspan=4, sticky="ew", padx=20, pady=(0, 15))
        
        # Video compression
        self.chk_video = ctk.CTkSwitch(
            options_frame,
            text="Compress Videos",
            onvalue=True,
            offvalue=False
        )
        self.chk_video.pack(side="left", padx=(0, 20))
        
        # PNG smart conversion
        self.chk_png = ctk.CTkSwitch(
            options_frame,
            text="Smart PNG-to-JPG",
            onvalue=True,
            offvalue=False
        )
        self.chk_png.pack(side="left", padx=(0, 20))
        
        # Create backups
        self.chk_backup = ctk.CTkSwitch(
            options_frame,
            text="Create Backups",
            onvalue=True,
            offvalue=False
        )
        self.chk_backup.pack(side="left")
        self.chk_backup.select()  # Enabled by default
        
        # Profile description
        self.lbl_profile_desc = ctk.CTkLabel(
            self.settings_frame,
            text="Target: 70% Quality | 1920px (Full HD)",
            text_color="#94a3b8",
            font=("Consolas", 11)
        )
        self.lbl_profile_desc.grid(row=2, column=0, columnspan=4, padx=20, pady=(0, 15), sticky="w")
    
    def _create_footer(self):
        """Create footer with progress and action buttons"""
        self.footer = ctk.CTkFrame(self, fg_color="transparent")
        self.footer.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 10))
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self.footer, height=20)
        self.progress_bar.pack(fill="x", pady=(0, 15))
        self.progress_bar.set(0)
        
        # Action buttons frame
        btn_frame = ctk.CTkFrame(self.footer, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        # Stop button (hidden initially)
        self.btn_stop = ctk.CTkButton(
            btn_frame,
            text="⏹️ Stop",
            fg_color="#dc2626",
            hover_color="#b91c1c",
            command=self._stop_processing,
            height=40,
            width=100,
            state="disabled"
        )
        self.btn_stop.pack(side="left", padx=(0, 10))
        
        # Start button
        self.btn_start = ctk.CTkButton(
            btn_frame,
            text="🚀 START OPTIMIZATION",
            font=("Segoe UI", 14, "bold"),
            height=45,
            width=250,
            command=self._start_optimization
        )
        self.btn_start.pack(side="right")
    
    def _create_status_bar(self):
        """Create status bar at bottom"""
        self.status_frame = ctk.CTkFrame(self, height=40, corner_radius=0, fg_color="#0f172a")
        self.status_frame.grid(row=4, column=0, sticky="ew", padx=0, pady=0)
        self.status_frame.pack_propagate(False)
        
        # Status label
        self.lbl_status = ctk.CTkLabel(
            self.status_frame,
            text="Ready - Shilezi v5.4 (2025)",
            text_color="#94a3b8",
            font=("Consolas", 10)
        )
        self.lbl_status.pack(side="left", padx=20)
        
        # FFmpeg status
        self.lbl_ffmpeg = ctk.CTkLabel(
            self.status_frame,
            text="FFmpeg: Checking...",
            text_color="#64748b",
            font=("Consolas", 10)
        )
        self.lbl_ffmpeg.pack(side="right", padx=20)
    
    def _update_file_status(self, filepath, text, color):
        """Safely update file status without creating new widgets"""
        if filepath in self.row_widgets:
            self.row_widgets[filepath]["status"].configure(text=text, text_color=color)
        else:
            # Log error but don't crash
            print(f"Warning: File not found in row_widgets: {os.path.basename(filepath)}")
    
    def _check_system(self):
        """Check system requirements and FFmpeg"""
        engine = OfficeCompressor()
        is_ready, message = engine.check_ffmpeg()
        
        if is_ready:
            self.lbl_ffmpeg.configure(text=f"FFmpeg: {message}", text_color="#4ade80")
            self.chk_video.select()
        else:
            self.lbl_ffmpeg.configure(text=f"FFmpeg: {message}", text_color="#f87171")
            self.chk_video.deselect()
            
            # Show warning if FFmpeg not found
            if "not found" in message.lower():
                self.after(1000, lambda: messagebox.showwarning(
                    "FFmpeg Not Found",
                    "FFmpeg not found. Video/audio compression will be disabled.\n\n"
                    "To enable video compression:\n"
                    "1. Download ffmpeg.exe from ffmpeg.org\n"
                    "2. Place it in the same folder as this application\n"
                    "3. Restart the application"
                ))
    
    def _add_files(self):
        """Add files to queue"""
        filenames = filedialog.askopenfilenames(
            title="Select Office Files",
            filetypes=[
                ("Office Files", "*.pptx *.docx *.xlsx"),
                ("PowerPoint", "*.pptx"),
                ("Word", "*.docx"),
                ("Excel", "*.xlsx"),
                ("All Files", "*.*")
            ]
        )
        
        for f in filenames:
            if f not in self.files:
                self.files.append(f)
                self._add_file_row(f)
        
        self._update_file_summary()
    
    def _add_folder(self):
        """Add all Office files from a folder"""
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith(('.pptx', '.docx', '.xlsx')):
                        f = os.path.join(root, file)
                        if f not in self.files:
                            self.files.append(f)
                            self._add_file_row(f)
            
            self._update_file_summary()
    
    def _add_file_row(self, filepath):
        """Add file entry to the list"""
        row = ctk.CTkFrame(self.scroll_frame, fg_color="#334155", corner_radius=6)
        row.pack(fill="x", pady=2, padx=5)
        
        # File icon and name
        ctk.CTkLabel(row, text="📄", font=("Segoe UI", 16)).pack(side="left", padx=10, pady=8)
        
        # Truncated filename
        display_name = self._truncate_filename(filepath, 50)
        ctk.CTkLabel(
            row,
            text=display_name,
            font=("Segoe UI", 11),
            anchor="w"
        ).pack(side="left", padx=5, fill="x", expand=True)
        
        # File size
        size = self._format_bytes(os.path.getsize(filepath))
        ctk.CTkLabel(
            row,
            text=size,
            text_color="gray",
            font=("Consolas", 10)
        ).pack(side="right", padx=10)
        
        # Status label
        status = ctk.CTkLabel(
            row,
            text="Pending",
            text_color="#94a3b8",
            font=("Segoe UI", 10)
        )
        status.pack(side="right", padx=15)
        
        # Store reference
        self.row_widgets[filepath] = {
            "status": status,
            "row": row
        }
    
    def _clear_files(self):
        """Clear all files from queue"""
        if self.is_processing:
            messagebox.showwarning("Processing", "Cannot clear files while processing")
            return
        
        self.files = []
        self.row_widgets = {}
        
        # Clear scroll frame
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        
        self._update_file_summary()
    
    def _update_file_summary(self):
        """Update file count and total size"""
        count = len(self.files)
        total_size = sum(os.path.getsize(f) for f in self.files if os.path.exists(f))
        
        self.lbl_file_count.configure(text=f"{count} file{'s' if count != 1 else ''} selected")
        self.lbl_total_size.configure(text=f"Total: {self._format_bytes(total_size)}")
    
    def _on_profile_change(self, choice):
        """Update profile description"""
        if choice in CONFIG["presets"]:
            preset = CONFIG["presets"][choice]
            self.lbl_profile_desc.configure(
                text=f"Target: {preset['quality']}% Quality | {preset['max_width']}px"
            )
    
    def _start_optimization(self):
        """Start the optimization process"""
        if not self.files:
            messagebox.showwarning("No Files", "Please add files first!")
            return
        
        if self.is_processing:
            return
        
        # Get settings from UI
        profile = self.profile_var.get()
        if profile not in CONFIG["presets"]:
            profile = "Balanced (Recommended)"
        
        preset = CONFIG["presets"][profile]
        replace_original = "Replace" in self.save_var.get()
        compress_video = self.chk_video.get()
        png_smart = self.chk_png.get()
        enable_backup = self.chk_backup.get()
        
        # Reset UI
        self.is_processing = True
        self.btn_start.configure(state="disabled", text="PROCESSING...")
        self.btn_stop.configure(state="normal")
        self.progress_bar.set(0)
        self.lbl_status.configure(text="Starting...", text_color="#60a5fa")
        
        # Start processing in separate thread
        thread = threading.Thread(
            target=self._run_optimization,
            args=(preset["quality"], preset["max_width"], replace_original, 
                  compress_video, png_smart, enable_backup),
            daemon=True
        )
        thread.start()
    
    def _run_optimization(self, quality, max_width, replace_original, 
                         compress_video, png_smart, enable_backup):
        """Run optimization engine in background thread"""
        engine = OfficeCompressor(
            quality=quality,
            max_width=max_width,
            compress_video=compress_video,
            png_smart_convert=png_smart,
            enable_backup=enable_backup
        )
        
        total_files = len(self.files)
        
        for idx, filepath in enumerate(self.files):
            if not self.is_processing:
                break
            
            # Update current file status using safe method
            self._thread_safe_update(
                lambda f=filepath: self._update_file_status(f, "Processing...", "#60a5fa")
            )
            
            # Determine output path
            if replace_original:
                out_path = filepath + ".optimized"
            else:
                base, ext = os.path.splitext(filepath)
                out_path = f"{base}_Optimized{ext}"
            
            # Create progress callback
            def progress_callback(p):
                overall_progress = (idx / total_files) + (p / 100 / total_files)
                self._thread_safe_update(lambda: self.progress_bar.set(overall_progress))
            
            # Create log callback
            def log_callback(msg):
                self._thread_safe_update(lambda: self.lbl_status.configure(text=msg))
            
            # Process the file
            success = engine.compress(
                filepath, 
                out_path, 
                progress_callback, 
                log_callback
            )
            
            # Update file status with safe method
            if success:
                if replace_original:
                    try:
                        os.replace(out_path, filepath)
                        status_text = "Replaced"
                    except Exception as e:
                        status_text = "Error"
                        success = False
                        if log_callback:
                            log_callback(f"Replace failed: {str(e)}")
                else:
                    status_text = "Saved"
                
                if success:
                    self._thread_safe_update(
                        lambda f=filepath, t=status_text: self._update_file_status(f, t, "#4ade80")
                    )
                else:
                    self._thread_safe_update(
                        lambda f=filepath: self._update_file_status(f, "Error", "#f87171")
                    )
            else:
                self._thread_safe_update(
                    lambda f=filepath: self._update_file_status(f, "Error", "#f87171")
                )
        
        # Update final status
        if self.is_processing:
            self.compression_stats = engine.get_statistics()
            
            self._thread_safe_update(lambda: self.progress_bar.set(1.0))
            self._thread_safe_update(lambda: self.lbl_status.configure(
                text=f"Complete! Processed {total_files} file{'s' if total_files != 1 else ''}",
                text_color="#4ade80"
            ))
            
            # Show statistics
            if self.compression_stats:
                self.after(500, self._show_statistics)
        
        # Reset UI
        self._thread_safe_update(lambda: self.btn_start.configure(
            state="normal", text="🚀 START OPTIMIZATION"
        ))
        self._thread_safe_update(lambda: self.btn_stop.configure(state="disabled"))
        self.is_processing = False
    
    def _stop_processing(self):
        """Stop the current processing"""
        if self.is_processing:
            self.is_processing = False
            self.lbl_status.configure(text="Stopping...", text_color="#fbbf24")
            self.btn_stop.configure(state="disabled")
    
    def _show_statistics(self):
        """Show compression statistics dialog"""
        if not self.compression_stats:
            return
        
        # Create dialog
        dialog = ctk.CTkToplevel(self)
        dialog.title("Compression Statistics")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Content
        ctk.CTkLabel(
            dialog,
            text="📊 Compression Report",
            font=("Segoe UI", 18, "bold")
        ).pack(pady=(20, 10))
        
        # Stats frame
        stats_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        stats_frame.pack(fill="both", expand=True, padx=30, pady=10)
        
        stats = self.compression_stats
        for label, value in [
            ("Files Processed:", stats.get("files_processed", 0)),
            ("Original Size:", stats.get("original_size", "0 B")),
            ("Total Savings:", f"{stats.get('savings_bytes', '0 B')} ({stats.get('savings_percent', '0%')})"),
            ("Processing Time:", stats.get("processing_time", "0s")),
            ("Average Speed:", stats.get("average_speed", "0 B/s"))
        ]:
            row = ctk.CTkFrame(stats_frame, fg_color="transparent")
            row.pack(fill="x", pady=5)
            
            ctk.CTkLabel(
                row,
                text=label,
                font=("Segoe UI", 11),
                width=150,
                anchor="w"
            ).pack(side="left")
            
            ctk.CTkLabel(
                row,
                text=str(value),
                font=("Consolas", 11, "bold"),
                text_color="#60a5fa"
            ).pack(side="left")
        
        # Close button
        ctk.CTkButton(
            dialog,
            text="Close",
            command=dialog.destroy,
            width=100
        ).pack(pady=20)
    
    def _open_settings(self):
        """Open settings dialog"""
        # Simple settings dialog - can be expanded
        dialog = ctk.CTkToplevel(self)
        dialog.title("Settings")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Content
        ctk.CTkLabel(
            dialog,
            text="Settings",
            font=("Segoe UI", 18, "bold")
        ).pack(pady=20)
        
        ctk.CTkLabel(
            dialog,
            text="Additional settings will be available in future versions.",
            text_color="gray"
        ).pack(pady=10)
        
        ctk.CTkButton(
            dialog,
            text="Close",
            command=dialog.destroy,
            width=100
        ).pack(pady=20)
    
    def _thread_safe_update(self, func):
        """Execute function in main thread (thread-safe GUI updates)"""
        self.after(0, func)
    
    def _truncate_filename(self, path, limit=40):
        """Truncate filename for display"""
        name = os.path.basename(path)
        if len(name) > limit:
            base, ext = os.path.splitext(name)
            return base[:limit - len(ext) - 3] + "..." + ext
        return name
    
    def _format_bytes(self, size):
        """Format bytes to human readable string"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point"""
    # Display Shilezi branding
    print("\n" + "="*70)
    print("⚡ OFFICE OPTIMIZER PRO v5.4")
    print("   Created by: SHILEZI (https://github.com/shilezi)")
    print(f"   Version: 5.4.0 | Year: 2025 | Build: Official")
    print("="*70)
    
    # Check for required modules
    try:
        import customtkinter
    except ImportError:
        print("Error: customtkinter is not installed.")
        print("Please install it using: pip install customtkinter")
        input("Press Enter to exit...")
        return
    
    try:
        from PIL import Image
    except ImportError:
        print("Error: Pillow is not installed.")
        print("Please install it using: pip install pillow")
        input("Press Enter to exit...")
        return
    
    # Verify authenticity
    is_authentic, auth_message = verify_authenticity()
    if not is_authentic:
        print(f"\n⚠️  WARNING: {auth_message}")
        print(f"   Please download the official version from:")
        print(f"   https://github.com/shilezi/office-optimizer-pro")
        input("\nPress Enter to exit...")
        return
    
    # Create and run application
    app = OfficeOptimizerApp()
    
    # Set application icon if available
    icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
    if os.path.exists(icon_path):
        app.iconbitmap(icon_path)
    
    # Start main loop
    app.mainloop()


if __name__ == "__main__":
    main()
