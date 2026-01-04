# Changelog

All notable changes to the PDF Viewer PCF Control will be documented in this file.

## [1.2.4] - 2026-01-04

### Bug Fixes (Navigation Synchronization)

- **Fixed sidebar/page mismatch** - Sidebar thumbnail now correctly matches current page in main viewer
  - Added `isNavigatingRef` flag to prevent IntersectionObserver from overriding programmatic navigation
  - IntersectionObserver skips page updates during navigation operations

- **Fixed thumbnail click navigation** - Clicking sidebar thumbnail now scrolls to correct page on first try
  - Rewrote `goToPage()` to use pending scroll pattern that waits for virtual window re-render
  - Added `useEffect` that scrolls AFTER state update ensures target page is in DOM
  - Uses `requestAnimationFrame` to guarantee DOM has updated before scrolling
  - Changed scroll behavior to 'instant' for programmatic navigation (eliminates race conditions)

- **Fixed text/canvas disappearing** - Page content no longer flickers during re-rendering
  - Added `everRenderedRef` to track pages that have been rendered at least once
  - Canvas stays visible even during re-rendering if it was previously rendered
  - Placeholder only shows for pages that have never been rendered

- **Fixed page number field navigation** - Arrow keys and page input now work correctly on first try

### Performance Improvements

- **Expanded virtual window buffer** - Increased from 3/5 pages to 5/8 pages for smoother navigation
- **Increased virtual scroll threshold** - Changed from 20 to 30 pages before enabling virtual scrolling
  - Documents with <=30 pages now render all pages (no virtual scrolling needed)

### Technical Details

- New refs: `isNavigatingRef`, `pendingScrollRef`, `everRenderedRef`
- Navigation flag prevents IntersectionObserver debounce from overriding programmatic navigation
- Pending scroll pattern ensures DOM is ready before attempting scroll
- New CONFIG constant: `VIRTUAL_SCROLL_THRESHOLD`

## [1.2.3] - 2025-12-30

### Performance Improvements (Large Document Optimization)

- **Virtual scrolling for pages** - Only renders pages near the current view
  - Reduces DOM nodes from 100+ to ~9 for large documents
  - Auto-activates for documents with >20 pages
  - Configurable buffer: 3 pages before, 5 pages after current page
  - Spacer divs maintain scroll height for smooth navigation

- **Thumbnail virtualization** - Only renders visible thumbnails in sidebar
  - Auto-activates for documents with >30 pages
  - Scroll-based range tracking with requestAnimationFrame throttling
  - Buffer of 5 thumbnails before and 10 after visible range

- **Parallel Dataverse requests** - Column discovery now runs in parallel
  - File and Image attribute queries execute simultaneously via Promise.all()
  - Reduces initialization time by 100-200ms

- **Memoized search match filtering** - O(1) lookup instead of O(n) filtering
  - Pre-computed `matchesByPage` Map using useMemo
  - Pre-computed `currentMatch` for instant comparison
  - Significant improvement for documents with many search matches

### Technical Details

- New CONFIG constants: `VIRTUAL_SCROLL_BUFFER_BEFORE`, `VIRTUAL_SCROLL_BUFFER_AFTER`
- Virtual scrolling uses spacer divs to maintain correct scroll height
- Thumbnail virtualization tracks visible range via scroll events
- Search matches grouped by page number for O(1) page lookup

## [1.2.2] - 2025-12-18

### Performance Improvements

- **IntersectionObserver for page visibility** - Replaced manual scroll position calculations with native browser API
  - Eliminates 60+ DOM queries per scroll event
  - More accurate current page detection using intersection ratios
  - Automatic pre-loading via rootMargin buffer
- **requestIdleCallback for thumbnail rendering** - Non-blocking thumbnail generation
  - Thumbnails near current page render immediately (high priority)
  - Remaining thumbnails render during browser idle time
  - Falls back to setTimeout for Safari and older browsers
  - Proper cleanup of idle callbacks on unmount

### Technical Details

- Multiple intersection thresholds (0, 0.1, 0.25, 0.5, 0.75, 1.0) for smooth current page tracking
- Thumbnail rendering split into priority batches for better perceived performance

## [1.2.1] - 2025-12-18

### Performance Improvements

- Added AbortController to all DataverseService fetch calls for proper request cancellation on unmount
- Added render task timeout cleanup - stale tasks auto-cancel after 10 seconds
- Periodic cleanup interval checks for stuck render tasks every 5 seconds

### Bug Fixes

- Fixed potential memory leak: Dataverse fetch requests now properly abort on component unmount
- Fixed potential stuck render tasks that could accumulate during rapid scrolling

## [1.2.0] - 2025-12-18

### Performance Improvements

- Added React.memo to ThumbnailPanel and OutlinePanel components (~20-30% render reduction)
- Wrapped all handler functions with useCallback to prevent cascading re-renders
- Added useMemo for columnsWithFiles computed value
- Removed redundant CSS will-change property
- Replaced magic numbers with CONFIG constants for better maintainability

### Bug Fixes

- Fixed memory leak: Blob URLs are now revoked when imageUrl changes or component unmounts
- Fixed race condition: Document load now uses loadIdRef to prevent stale callbacks
- Fixed PDF.js version mismatch: Updated pdfjs-dist to 4.10.38 to match CDN worker

### Code Quality

- Removed unused singleton exports from PdfService and DataverseService
- Removed console.error calls from production code
- Added CONFIG constants for thumbnail rendering delays and batch sizes

## [1.1.3] - 2025-12-18

### Fixed

- Fixed rendering glitches where text from different pages would overlap/merge
- Added canvas clearing before each render to prevent ghosting
- Added rendering lock (ref-based) to prevent race conditions during concurrent renders
- Added scroll throttling to prevent excessive render calls during fast scrolling
- Proper cleanup of rendering state on component unmount or scale/rotation changes

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

- Initial release of PDF Viewer PCF Control
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
