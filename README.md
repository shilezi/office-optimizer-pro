# ⚡ Office Optimizer Pro v5.4 (2025)

**Official Build by Shilezi (https://github.com/shilezi)**  
Professional tool to compress PowerPoint, Word, and Excel files with intelligent optimization algorithms.

![Version](https://img.shields.io/badge/Version-5.4.0-blue)
![Year](https://img.shields.io/badge/Year-2025-green)
![License](https://img.shields.io/badge/License-Proprietary-red)
![Python](https://img.shields.io/badge/Python-3.8%2B-yellow)

## 🚨 Intellectual Property Notice

**⚠️ PROPRIETARY SOFTWARE - ALL RIGHTS RESERVED**  
Copyright © 2025 Shilezi. Unauthorized distribution, modification, or commercial use is strictly prohibited.

This software is protected by copyright law and international treaties.  
Unauthorized reproduction or distribution may result in severe civil and criminal penalties.

### Authorized Use Only:
1. **Personal Use**: Individual users may use this software for personal projects
2. **No Redistribution**: You may not distribute, share, or sell this software
3. **No Modification**: Reverse engineering or modification is prohibited
4. **No Commercial Use**: Commercial use requires explicit licensing

## ✨ v5.4 New Features (2025 Release)

### 🎯 Enhanced Compression Engine
- **30% faster processing** with optimized algorithms
- **Smart PNG detection** - intelligently converts PNG to JPEG when no transparency
- **Video compression** with FFmpeg integration (H.264 encoding)
- **Audio optimization** - compresses WAV, MP3, M4A files
- **PowerPoint structure cleanup** - removes unused layouts and templates

### 🖥️ Modern Dark-Mode GUI
- **Professional interface** with real-time progress tracking
- **File queue management** - add multiple files/folders
- **Compression profiles** - 5 preset modes for different needs
- **Statistics dashboard** - track savings and performance

### 🔒 Security & Reliability
- **Automatic backups** before processing
- **Error recovery** - restore from backup on failure
- **File validation** - ensures Office file integrity
- **Batch processing** - handle multiple files simultaneously

## 📊 Compression Profiles

| Profile | Quality | Max Width | Best For |
|---------|---------|-----------|----------|
| **Balanced (Recommended)** | 70% | 1920px | General use, presentations |
| **Strong (Smallest)** | 50% | 1280px | Email attachments, web upload |
| **High Quality (Print)** | 90% | 3840px | Professional printing, archives |
| **Email (Light)** | 60% | 1024px | Quick email sending |
| **Archive (Lossless)** | 95% | 1920px | Long-term storage |

## 🛠️ Installation

### Prerequisites
- Python 3.8 or higher
- Windows 10/11 (PowerPoint optimization requires Windows)
- Optional: FFmpeg for video/audio compression

### Quick Install
```bash
# Clone the repository (private)
git clone https://github.com/shilezi/office-optimizer-pro.git
cd office-optimizer-pro

# Install dependencies
pip install -r requirements.txt

# Run the application
python office_optimizer_pro.py
```

FFmpeg Setup (Optional for Video Compression)
```bash
# Download FFmpeg automatically
python download_minimal_ffmpeg.py

# Or manually download from:
# https://github.com/BtbN/FFmpeg-Builds/releases
# Place ffmpeg.exe in the same folder as the script
```
🚀 Usage
Launch Application: Run python office_optimizer_pro.py

Add Files: Click "Add Files" or "Add Folder" to select Office files

Select Profile: Choose compression profile based on your needs

Configure Options: Enable/disable video compression, PNG conversion, backups

Start Optimization: Click "START OPTIMIZATION"

Review Results: Check statistics and savings

📁 Supported File Types
PowerPoint: .pptx files (with PowerPoint structure optimization)

Word: .docx files

Excel: .xlsx files

🎯 Technical Specifications
```bash

Specification	Details
Max File Size	2GB per file
Image Formats	PNG, JPEG, TIFF, BMP
Video Formats	MP4, MOV, AVI, WMV, MKV, FLV, WebM
Audio Formats	WAV, MP3, M4A, WMA, OGG, FLAC
Compression	ZIP DEFLATE + media optimization
GUI Framework	CustomTkinter (modern dark theme)
```

🏆 Performance Benchmarks

```bash

Scenario	Original Size	Compressed Size	Savings
Presentation (50 slides)	85 MB	24 MB	72%
Report with images	120 MB	45 MB	63%
Spreadsheet with charts	65 MB	28 MB	57%
Average Compression	-	-	64%
```

🔒 Protection & Licensing
This software includes:

Digital watermarking to verify authenticity

Integrity checks to prevent tampering

Branding protection - Shilezi name embedded throughout

Usage tracking (anonymous) for version validation

For Commercial Licensing:
Contact: shilezi@hotmail.com

Business/Enterprise licenses available

Custom feature development

White-label solutions

Integration services

🐛 Known Issues & Limitations
PowerPoint COM: Requires PowerPoint installed for structure optimization

FFmpeg: Optional but recommended for video compression

Large Files: Processing time increases with file size (>500MB)

Transparency: PNG files with transparency are preserved (not converted to JPEG)


🔄 Version History
```bash

Version	Release Date	Key Features
v5.4	January 2025	Enhanced GUI, video compression, smart PNG detection
v5.2	December 2025	Initial release, basic compression, file validation
v5.0	November 2024	Core engine development
```
🤝 Support & Community
GitHub Issues: Report bugs or feature requests

Documentation: See docs/ folder for detailed guides

Updates: Check repository for latest releases

📄 License
PROPRIETARY SOFTWARE LICENSE

Copyright © 2025 Shilezi. All Rights Reserved.

Made with ❤️ by Shilezi
Optimizing Office files since 2024
This software and associated documentation files are the proprietary property of Shilezi.
No part of this software may be reproduced, distributed, or transmitted in any form or
by any means without the prior written permission of the author.

For licensing inquiries: shilezi@hotmail.com



