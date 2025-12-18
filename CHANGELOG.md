# Changelog

All notable changes to the RSM PDF Viewer PCF Control will be documented in this file.

## [1.1.2] - 2025-12-18

### Fixed

- Page input no longer glitches when typing - uses local state until blur/Enter
- Page input now limits digit length to prevent excessively long numbers
- Input selects all text on focus for easy replacement
- Escape key cancels page input and resets to current page

## [1.1.1] - 2025-12-17

### Fixed

- PDF now centers properly when sidebar is closed
- Improved sidebar collapse animation with smooth transitions

## [1.1.0] - 2025-12-17

### Fixed

- Thumbnail images no longer appear stretched (added proper aspect ratio handling)
- File switching now works correctly - selecting a different file from dropdown loads the new document
- Fixed "Cannot use the same canvas during multiple render() operations" error by using React key props to force canvas remount
- Proper state reset before loading new documents (clears rendered pages, text content, viewports, etc.)
- Fixed pdfService.destroy() to be properly awaited

### Technical Details

- Added `documentKey` state to force PdfCanvas and ThumbnailPanel remount on document switch
- Added CSS `object-fit: contain` and `max-height` for thumbnail images
- Improved cleanup flow in loadDocument function

## [1.0.16] - 2025-12-15

### Added

- Page caching in PdfService for improved performance
- Parallelized page viewport fetching (10 pages at a time)
- GPU acceleration hints for smoother scrolling

### Fixed

- Various performance optimizations
- Improved error handling

## [1.0.0] - 2024-12-15

### Added

- Initial release of RSM PDF Viewer PCF Control
- PDF rendering using PDF.js v4.10.38
- Image file support (JPG, PNG, GIF, BMP, WebP, SVG, TIFF)
- Zoom controls (25%-500%, fit page, fit width, auto)
- Page navigation with scroll tracking
- Text search with highlights (Ctrl+F)
- Thumbnails sidebar panel
- Outline/bookmarks navigation
- Document rotation (left/right)
- Dark mode toggle
- Fullscreen mode
- Print functionality
- Download functionality
- Keyboard shortcuts
- Auto-discovery of Dataverse file columns
- Default file column configuration
- Responsive design
- Test mode for development (URL input, file upload)

### Technical Details

- Built with React and TypeScript
- PCF Virtual Control (React-based)
- Bundle size: ~524 KB (minified)
- PDF.js worker loaded from CDN

## Roadmap

### Planned Features

- [ ] Annotation support (highlights, comments)
- [ ] Form field support
- [ ] Drag-to-pan navigation
- [ ] Pinch-to-zoom on touch devices
- [ ] Multi-document tabs
- [ ] Recent documents list
- [ ] Document comparison view

### Under Consideration

- OCR for scanned PDFs (via Azure AI Document Intelligence)
- Digital signature verification
- Watermark overlay
- Custom stamp annotations
