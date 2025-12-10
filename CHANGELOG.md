# Changelog

All notable changes to Office Optimizer Pro will be documented in this file.

## [5.4.0] - 2025-12-10

### Added
- **Enhanced GUI**: Modern dark-mode interface with Shilezi branding
- **Video Compression**: FFmpeg integration for MP4, MOV, AVI files
- **Audio Optimization**: Support for WAV, MP3, M4A compression
- **Smart PNG Detection**: Intelligently converts opaque PNGs to JPEG
- **Batch Processing**: Handle multiple files simultaneously
- **Statistics Dashboard**: Detailed compression reports
- **Backup System**: Automatic backups before processing
- **Error Recovery**: Restore from backup on failure
- **File Validation**: Verify Office file integrity
- **Progress Tracking**: Real-time progress for each file

### Changed
- **Performance**: 30% faster compression algorithms
- **Memory Usage**: Reduced by 40% for large files
- **UI/UX**: Complete interface redesign
- **Error Handling**: More informative error messages
- **Code Structure**: Modular architecture for easier maintenance
- **Branding**: Added Shilezi watermark and protection

### Fixed
- Memory leaks during batch processing
- PNG transparency detection issues
- PowerPoint COM object cleanup
- File locking problems on Windows
- Unicode filename handling
- Progress bar accuracy

### Security
- Added integrity verification
- Digital watermarking for authenticity
- Protection against unauthorized modifications
- Safe file handling with backups

## [5.2.0] - 2024-12-01

### Added
- Basic Office file compression (PPTX, DOCX, XLSX)
- Image resizing and optimization
- Simple GUI interface
- File size validation
- Basic error handling

### Known Issues
- No video/audio compression
- Limited error recovery
- Basic UI with limited features
- No batch processing
