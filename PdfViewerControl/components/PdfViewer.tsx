/**
 * PdfViewer - Main React component for the PDF Viewer PCF control
 */

import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import { PdfService, PdfMetadata, OutlineItem, TextContent, PageViewport, RenderTask } from '../services/PdfService';
import { DataverseService, FileColumn, ViewerConfig } from '../services/DataverseService';

// Configuration constants
const CONFIG = {
    DEFAULT_SCALE: 1.0,
    MIN_SCALE: 0.25,
    MAX_SCALE: 5.0,
    ZOOM_STEP: 1.25,
    RENDER_BUFFER_PX: 500,
    THUMBNAIL_SCALE: 0.2,
    SCROLL_THROTTLE_MS: 100,
    IMAGE_EXTENSIONS: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'svg', 'tiff', 'tif'],
};

export interface IPdfViewerProps {
    defaultFileColumn?: string;
    showToolbar?: boolean;
    showSidebar?: boolean;
    defaultZoom?: string;
    theme?: string;
    width?: number;
    height?: number;
    onPageChange?: (currentPage: number, totalPages: number) => void;
    onColumnChange?: (columnName: string) => void;
}

type ViewerState = 'loading' | 'config' | 'viewing' | 'error';
type FileType = 'pdf' | 'image' | null;

export const PdfViewer: React.FC<IPdfViewerProps> = ({
    defaultFileColumn,
    showToolbar = true,
    showSidebar = true,
    defaultZoom = 'auto',
    theme = 'light',
    width,
    height,
    onPageChange,
    onColumnChange
}) => {
    // Services
    const pdfServiceRef = useRef<PdfService>(new PdfService());
    const dataverseServiceRef = useRef<DataverseService>(new DataverseService());

    // State
    const [viewerState, setViewerState] = useState<ViewerState>('loading');
    const [loadingMessage, setLoadingMessage] = useState('Initializing...');
    const [loadingProgress, setLoadingProgress] = useState(0);
    const [errorMessage, setErrorMessage] = useState('');

    // Viewer config state
    const [viewerConfig, setViewerConfig] = useState<ViewerConfig | null>(null);
    const [selectedColumn, setSelectedColumn] = useState<string | null>(null);

    // Document state
    const [fileType, setFileType] = useState<FileType>(null);
    const [currentPage, setCurrentPage] = useState(1);
    const [totalPages, setTotalPages] = useState(0);
    const [currentScale, setCurrentScale] = useState(CONFIG.DEFAULT_SCALE);
    const [currentRotation, setCurrentRotation] = useState(0);
    const [fitMode, setFitMode] = useState<string | null>(defaultZoom);
    const [metadata, setMetadata] = useState<PdfMetadata | null>(null);

    // UI state - sidebar hidden by default for performance
    const [sidebarVisible, setSidebarVisible] = useState(false);
    const [sidebarTab, setSidebarTab] = useState<'thumbnails' | 'outline'>('thumbnails');
    const [darkMode, setDarkMode] = useState(false);
    const [showDocInfo, setShowDocInfo] = useState(false);
    const [showShortcuts, setShowShortcuts] = useState(false);

    // Search state
    const [searchText, setSearchText] = useState('');
    const [searchMatches, setSearchMatches] = useState<{ pageNum: number; itemIndex: number; startOffset: number; length: number }[]>([]);
    const [currentMatchIndex, setCurrentMatchIndex] = useState(-1);
    const [showFindPanel, setShowFindPanel] = useState(false);
    const searchInputRef = useRef<HTMLInputElement>(null);
    const searchDebounceRef = useRef<NodeJS.Timeout | null>(null);

    // Rendered pages tracking
    const [pageViewports, setPageViewports] = useState<PageViewport[]>([]);
    const [renderedPages, setRenderedPages] = useState<Map<number, 'rendering' | 'done'>>(new Map());
    const [textContent, setTextContent] = useState<Map<number, TextContent>>(new Map());
    const [outline, setOutline] = useState<OutlineItem[]>([]);

    // Image state
    const [imageUrl, setImageUrl] = useState<string | null>(null);

    // Document key - changes when loading a new document to force canvas remount
    const [documentKey, setDocumentKey] = useState(0);

    // Test mode state
    const [testUrl, setTestUrl] = useState('https://mozilla.github.io/pdf.js/web/compressed.tracemonkey-pldi-09.pdf');
    const fileInputRef = useRef<HTMLInputElement>(null);

    // Refs
    const canvasContainerRef = useRef<HTMLDivElement>(null);
    const containerRef = useRef<HTMLDivElement>(null);

    // Theme detection
    useEffect(() => {
        if (theme === 'auto') {
            const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
            setDarkMode(prefersDark);
        } else {
            setDarkMode(theme === 'dark');
        }
    }, [theme]);

    // Initialize - discover file columns via Dataverse
    useEffect(() => {
        const initialize = async () => {
            setLoadingMessage('Detecting form context...');
            const dataverseService = dataverseServiceRef.current;

            if (defaultFileColumn) {
                dataverseService.setDefaultColumn(defaultFileColumn);
            }

            const success = await dataverseService.initialize();
            const config = dataverseService.getConfig();
            setViewerConfig(config);

            if (!success) {
                // Not in Dataverse context - show test mode
                setViewerState('config');
                return;
            }

            const columnsWithFiles = dataverseService.getColumnsWithFiles();
            if (columnsWithFiles.length === 0) {
                setViewerState('config');
                return;
            }

            // Try to auto-select a column
            const autoSelectColumn = dataverseService.getAutoSelectColumn();
            if (autoSelectColumn) {
                loadDocument(autoSelectColumn);
            } else {
                setViewerState('config');
            }
        };

        initialize();

        return () => {
            pdfServiceRef.current.destroy();
        };
    }, [defaultFileColumn]);

    // Notify parent of page changes
    useEffect(() => {
        onPageChange?.(currentPage, totalPages);
    }, [currentPage, totalPages, onPageChange]);

    // Load document from selected column
    const loadDocument = useCallback(async (columnName: string) => {
        setViewerState('loading');
        setLoadingMessage('Loading document...');
        setLoadingProgress(0);
        setSelectedColumn(columnName);
        onColumnChange?.(columnName);

        const dataverseService = dataverseServiceRef.current;
        const pdfService = pdfServiceRef.current;

        // Reset all state before loading new document
        await pdfService.destroy();
        setRenderedPages(new Map());
        setTextContent(new Map());
        setPageViewports([]);
        setOutline([]);
        setMetadata(null);
        setImageUrl(null);
        setFileType(null);
        setCurrentPage(1);
        setTotalPages(0);
        // Increment document key to force canvas remount (prevents canvas reuse errors)
        setDocumentKey(prev => prev + 1);

        try {
            const fileUrl = dataverseService.getFileUrl(columnName);

            // Skip HEAD request - detect file type from column config instead (200-500ms faster)
            // Check if this is an image column based on metadata
            if (isImageFile(columnName)) {
                // Load as image
                setFileType('image');
                setImageUrl(fileUrl + '?_=' + Date.now());
                setTotalPages(1);
                setCurrentPage(1);
                setViewerState('viewing');
            } else {
                // Load as PDF - optimized for speed
                setFileType('pdf');
                setLoadingMessage('Loading PDF...');

                await pdfService.loadFromUrl(fileUrl, (loaded, total) => {
                    if (total > 0) setLoadingProgress((loaded / total) * 100);
                });

                const pageCount = pdfService.getPageCount();
                setTotalPages(pageCount);
                setCurrentPage(1);

                // Get first page viewport and show immediately
                const firstViewport = await pdfService.getPageViewport(1, 1);
                setPageViewports([firstViewport]);

                // Show viewer NOW - don't wait for anything else
                setViewerState('viewing');

                // Load ALL remaining viewports in a single batch after render
                requestAnimationFrame(() => {
                    (async () => {
                        const allViewports: PageViewport[] = [firstViewport];
                        for (let i = 2; i <= pageCount; i++) {
                            const vp = await pdfService.getPageViewport(i, 1);
                            allViewports.push(vp);
                        }
                        setPageViewports(allViewports);

                        // Load metadata after viewports
                        pdfService.getMetadata().then(setMetadata).catch(() => { /* ignore */ });
                        pdfService.getOutline().then(setOutline).catch(() => { /* ignore */ });
                    })();
                });
            }
        } catch (error) {
            console.error('Error loading document:', error);
            const errMsg = error instanceof Error ? error.message : 'Unknown error';
            if (errMsg.includes('400') || errMsg.includes('404') || errMsg.includes('not found')) {
                setErrorMessage('No file found in this column. Please ensure a file has been uploaded.');
            } else if (errMsg.includes('Invalid PDF')) {
                setErrorMessage('The file is not a valid PDF. Please check the file format.');
            } else {
                setErrorMessage(`Failed to load document: ${errMsg}`);
            }
            setViewerState('error');
        }
    }, [onColumnChange]);

    // Check if file is an image by extension
    const isImageFile = (columnName: string): boolean => {
        const col = viewerConfig?.fileColumns.find(c => c.logicalName === columnName);
        if (col?.type === 'Image') return true;
        // Could also check displayName for file extensions
        return false;
    };

    // Load PDF from URL (test mode)
    const loadFromUrl = async (url: string) => {
        setViewerState('loading');
        setLoadingMessage('Loading PDF from URL...');
        setLoadingProgress(0);
        setSelectedColumn('test-url');

        const pdfService = pdfServiceRef.current;

        try {
            setFileType('pdf');
            await pdfService.loadFromUrl(url, (loaded, total) => {
                setLoadingProgress((loaded / total) * 100);
            });

            const viewports = await pdfService.getAllPageViewports();
            setPageViewports(viewports);
            setTotalPages(pdfService.getPageCount());
            setCurrentPage(1);

            const meta = await pdfService.getMetadata();
            setMetadata(meta);

            const docOutline = await pdfService.getOutline();
            setOutline(docOutline);

            setViewerState('viewing');
        } catch (error) {
            console.error('Error loading PDF:', error);
            setErrorMessage('Failed to load PDF. Check the URL and try again.');
            setViewerState('error');
        }
    };

    // Determine if file is an image by MIME type or extension
    const isImageFileByNameOrType = (file: File): boolean => {
        // Check MIME type first
        if (file.type.startsWith('image/')) {
            return true;
        }
        // Check by extension (browser might not correctly detect MIME type)
        const extension = file.name.split('.').pop()?.toLowerCase() || '';
        return CONFIG.IMAGE_EXTENSIONS.includes(extension);
    };

    // Determine if file is a PDF by MIME type or extension
    const isPdfFileByNameOrType = (file: File): boolean => {
        // Check MIME type first
        if (file.type === 'application/pdf') {
            return true;
        }
        // Check by extension
        const extension = file.name.split('.').pop()?.toLowerCase() || '';
        return extension === 'pdf';
    };

    // Load PDF from local file (test mode)
    const loadFromFile = async (file: File) => {
        setViewerState('loading');
        setLoadingMessage(`Loading ${file.name}...`);
        setLoadingProgress(0);
        setSelectedColumn(file.name);

        const pdfService = pdfServiceRef.current;

        try {
            // Determine file type by MIME type or extension
            const isImage = isImageFileByNameOrType(file);
            const isPdf = isPdfFileByNameOrType(file);

            if (isImage) {
                // Handle image files
                setFileType('image');
                const url = URL.createObjectURL(file);
                setImageUrl(url);
                setTotalPages(1);
                setCurrentPage(1);
                setViewerState('viewing');
            } else if (isPdf || !isImage) {
                // Handle PDF files (or try as PDF if unknown type)
                setFileType('pdf');
                await pdfService.loadFromBlob(file, (loaded, total) => {
                    if (total > 0) {
                        setLoadingProgress((loaded / total) * 100);
                    }
                });

                const viewports = await pdfService.getAllPageViewports();
                setPageViewports(viewports);
                setTotalPages(pdfService.getPageCount());
                setCurrentPage(1);

                const meta = await pdfService.getMetadata();
                setMetadata(meta);

                const docOutline = await pdfService.getOutline();
                setOutline(docOutline);

                setViewerState('viewing');
            }
        } catch (error) {
            console.error('Error loading file:', error);
            const errorMsg = error instanceof Error ? error.message : 'Unknown error';
            setErrorMessage(`Failed to load file: ${errorMsg}`);
            setViewerState('error');
        }
    };

    // Handle file input change
    const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (file) {
            loadFromFile(file);
        }
    };

    // Calculate scale for fit modes
    const calculateScale = useCallback((mode: string): number => {
        if (!canvasContainerRef.current || pageViewports.length === 0) {
            return CONFIG.DEFAULT_SCALE;
        }

        const viewport = pageViewports[0];
        const containerWidth = canvasContainerRef.current.clientWidth - 60;
        const containerHeight = canvasContainerRef.current.clientHeight - 40;

        const rotation = currentRotation % 180 !== 0;
        const pageWidth = rotation ? viewport.height : viewport.width;
        const pageHeight = rotation ? viewport.width : viewport.height;

        switch (mode) {
            case 'page-width':
                return containerWidth / pageWidth;
            case 'page-fit':
                return Math.min(containerWidth / pageWidth, containerHeight / pageHeight);
            case 'auto':
                return Math.min(1, containerWidth / pageWidth);
            default:
                return parseFloat(mode) / 100 || CONFIG.DEFAULT_SCALE;
        }
    }, [pageViewports, currentRotation]);

    // Apply zoom
    const applyZoom = useCallback((scale: number, mode: string | null = null) => {
        const newScale = Math.max(CONFIG.MIN_SCALE, Math.min(CONFIG.MAX_SCALE, scale));
        setCurrentScale(newScale);
        setFitMode(mode);
        setRenderedPages(new Map());
        setTextContent(new Map());
    }, []);

    // Zoom handlers
    const zoomIn = () => {
        applyZoom(currentScale * CONFIG.ZOOM_STEP);
    };

    const zoomOut = () => {
        applyZoom(currentScale / CONFIG.ZOOM_STEP);
    };

    const setZoom = (value: string) => {
        if (['auto', 'page-fit', 'page-width'].includes(value)) {
            applyZoom(calculateScale(value), value);
        } else {
            applyZoom(parseFloat(value));
        }
    };

    // Navigation handlers - scroll within container, not browser window
    const goToPage = (page: number) => {
        const targetPage = Math.max(1, Math.min(totalPages, page));
        setCurrentPage(targetPage);

        // Scroll within the PDF container, not the browser window
        if (!canvasContainerRef.current) return;

        const container = canvasContainerRef.current.querySelector('.pdf-page-wrapper-container');
        const pageElement = canvasContainerRef.current.querySelector<HTMLElement>(`[data-page="${targetPage}"]`);

        if (container && pageElement) {
            const scrollTop = pageElement.offsetTop - 20; // 20px padding from top

            container.scrollTo({
                top: scrollTop,
                behavior: 'smooth'
            });
        }
    };

    const previousPage = () => goToPage(currentPage - 1);
    const nextPage = () => goToPage(currentPage + 1);

    // Rotation handlers
    const rotateLeft = () => {
        setCurrentRotation((prev) => (prev - 90 + 360) % 360);
        setRenderedPages(new Map());
        setTextContent(new Map());
    };

    const rotateRight = () => {
        setCurrentRotation((prev) => (prev + 90) % 360);
        setRenderedPages(new Map());
        setTextContent(new Map());
    };

    // Toggle handlers
    const toggleSidebar = () => setSidebarVisible(prev => !prev);
    const toggleDarkMode = () => setDarkMode(prev => !prev);
    const toggleFullscreen = () => {
        if (!document.fullscreenElement) {
            containerRef.current?.requestFullscreen();
        } else {
            document.exitFullscreen();
        }
    };

    // Search - find all matches with their positions (internal implementation)
    const performSearch = useCallback((text: string) => {
        if (!text || text.length < 1) {
            setSearchMatches([]);
            setCurrentMatchIndex(-1);
            return;
        }

        const matches: { pageNum: number; itemIndex: number; startOffset: number; length: number }[] = [];
        const searchLower = text.toLowerCase();

        textContent.forEach((content, pageNum) => {
            content.items.forEach((item, itemIndex) => {
                const itemLower = item.str.toLowerCase();
                let startIndex = 0;

                // Find all occurrences within this text item
                while (true) {
                    const foundIndex = itemLower.indexOf(searchLower, startIndex);
                    if (foundIndex === -1) break;

                    matches.push({
                        pageNum,
                        itemIndex,
                        startOffset: foundIndex,
                        length: text.length
                    });
                    startIndex = foundIndex + 1;
                }
            });
        });

        // Sort by page number, then by item index
        matches.sort((a, b) => {
            if (a.pageNum !== b.pageNum) return a.pageNum - b.pageNum;
            return a.itemIndex - b.itemIndex;
        });

        setSearchMatches(matches);

        if (matches.length > 0) {
            setCurrentMatchIndex(0);
            scrollToMatch(matches[0]);
        } else {
            setCurrentMatchIndex(-1);
        }
    }, [textContent]);

    // Debounced search handler - 150ms delay to prevent lag while typing
    const handleSearch = useCallback((text: string) => {
        setSearchText(text);

        // Clear any pending search
        if (searchDebounceRef.current) {
            clearTimeout(searchDebounceRef.current);
        }

        // Empty text - clear immediately
        if (!text || text.length < 1) {
            setSearchMatches([]);
            setCurrentMatchIndex(-1);
            return;
        }

        // Debounce the actual search
        searchDebounceRef.current = setTimeout(() => {
            performSearch(text);
        }, 150);
    }, [performSearch]);

    // Scroll to a specific match
    const scrollToMatch = useCallback((match: { pageNum: number; itemIndex: number }) => {
        if (!canvasContainerRef.current) return;

        const container = canvasContainerRef.current.querySelector('.pdf-page-wrapper-container');
        const pageElement = canvasContainerRef.current.querySelector<HTMLElement>(`[data-page="${match.pageNum}"]`);

        if (container && pageElement) {
            // Find the highlight element for this match
            const highlightEl = pageElement.querySelector<HTMLElement>(`.pdf-text-highlight[data-item="${match.itemIndex}"]`);

            if (highlightEl) {
                // Scroll so the highlight is visible (centered if possible)
                const containerHeight = container.clientHeight;
                const highlightTop = pageElement.offsetTop + highlightEl.offsetTop;
                const scrollTarget = highlightTop - containerHeight / 3;

                container.scrollTo({
                    top: Math.max(0, scrollTarget),
                    behavior: 'smooth'
                });
            } else {
                // Fallback to page scroll
                container.scrollTo({
                    top: pageElement.offsetTop - 20,
                    behavior: 'smooth'
                });
            }
        }
    }, []);

    const searchNext = useCallback(() => {
        if (searchMatches.length === 0) return;
        const nextIndex = (currentMatchIndex + 1) % searchMatches.length;
        setCurrentMatchIndex(nextIndex);
        scrollToMatch(searchMatches[nextIndex]);
    }, [searchMatches, currentMatchIndex, scrollToMatch]);

    const searchPrev = useCallback(() => {
        if (searchMatches.length === 0) return;
        const prevIndex = (currentMatchIndex - 1 + searchMatches.length) % searchMatches.length;
        setCurrentMatchIndex(prevIndex);
        scrollToMatch(searchMatches[prevIndex]);
    }, [searchMatches, currentMatchIndex, scrollToMatch]);

    // Open find panel
    const openFindPanel = useCallback(() => {
        setShowFindPanel(true);
        setTimeout(() => searchInputRef.current?.focus(), 100);
    }, []);

    // Close find panel
    const closeFindPanel = useCallback(() => {
        // Clear any pending search debounce
        if (searchDebounceRef.current) {
            clearTimeout(searchDebounceRef.current);
            searchDebounceRef.current = null;
        }
        setShowFindPanel(false);
        setSearchText('');
        setSearchMatches([]);
        setCurrentMatchIndex(-1);
    }, []);

    // Print
    const handlePrint = () => window.print();

    // Download
    const handleDownload = async () => {
        if (!selectedColumn || !viewerConfig) return;

        try {
            const dataverseService = dataverseServiceRef.current;
            const { blob, contentType } = await dataverseService.fetchFile(selectedColumn);

            const col = viewerConfig.fileColumns.find(c => c.logicalName === selectedColumn);
            let filename = col?.displayName || 'document';

            // Add extension if needed
            if (contentType.includes('pdf') && !filename.toLowerCase().endsWith('.pdf')) {
                filename += '.pdf';
            } else if (contentType.includes('image') && !(/\.(jpg|jpeg|png|gif|bmp|webp)$/i.exec(filename))) {
                filename += '.jpg';
            }

            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Download failed:', error);
        }
    };

    // Keyboard shortcuts - only when PDF viewer has focus
    useEffect(() => {
        const handleKeyDown = (e: KeyboardEvent) => {
            // Check if event target is within our container - if not, ignore
            const target = e.target as HTMLElement;
            if (!containerRef.current?.contains(target)) {
                return; // Event is outside PDF viewer, ignore it
            }

            // Skip shortcuts if typing in any input field (including contenteditable)
            if (target instanceof HTMLInputElement ||
                target instanceof HTMLTextAreaElement ||
                target.isContentEditable) {
                // Only allow Escape and Enter in our search input
                if (target !== searchInputRef.current) {
                    return;
                }
            }

            // Handle Escape - close panels
            if (e.key === 'Escape') {
                if (showFindPanel) {
                    e.preventDefault();
                    closeFindPanel();
                    return;
                }
                setShowDocInfo(false);
                setShowShortcuts(false);
                return;
            }

            // Handle Enter/Shift+Enter in search input for next/prev
            if (e.target === searchInputRef.current) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    if (e.shiftKey) {
                        searchPrev();
                    } else {
                        searchNext();
                    }
                }
                return;
            }

            // Ctrl/Cmd shortcuts
            if (e.ctrlKey || e.metaKey) {
                switch (e.key.toLowerCase()) {
                    case 'f':
                        e.preventDefault();
                        openFindPanel();
                        break;
                    case 'p':
                        e.preventDefault();
                        handlePrint();
                        break;
                    case 'g':
                        e.preventDefault();
                        if (e.shiftKey) {
                            searchPrev();
                        } else {
                            searchNext();
                        }
                        break;
                }
                return;
            }

            // Only safe shortcuts - no single letters that could interfere with typing
            switch (e.key) {
                case 'ArrowLeft':
                case 'PageUp':
                    e.preventDefault();
                    previousPage();
                    break;
                case 'ArrowRight':
                case 'PageDown':
                    e.preventDefault();
                    nextPage();
                    break;
                case 'Home':
                    e.preventDefault();
                    goToPage(1);
                    break;
                case 'End':
                    e.preventDefault();
                    goToPage(totalPages);
                    break;
            }
        };

        document.addEventListener('keydown', handleKeyDown);
        return () => document.removeEventListener('keydown', handleKeyDown);
    }, [currentPage, totalPages, showFindPanel, closeFindPanel, openFindPanel, searchNext, searchPrev]);

    // Render loading state
    if (viewerState === 'loading') {
        return (
            <div className="pdf-viewer-container" data-theme={darkMode ? 'dark' : 'light'} ref={containerRef}
                 style={{ width: width || '100%', height: height || 500 }}>
                <div className="pdf-loading-overlay">
                    <div className="pdf-loading-spinner" />
                    <div className="pdf-loading-text">{loadingMessage}</div>
                    {loadingProgress > 0 && (
                        <div className="pdf-progress-container">
                            <div className="pdf-progress-bar" style={{ width: `${loadingProgress}%` }} />
                        </div>
                    )}
                </div>
            </div>
        );
    }

    // Render config/picker state
    if (viewerState === 'config') {
        const columnsWithFiles = viewerConfig?.fileColumns.filter(c => c.hasFile) || [];
        const isTestMode = !viewerConfig?.tableName || !viewerConfig?.recordId;

        return (
            <div className="pdf-viewer-container" data-theme={darkMode ? 'dark' : 'light'} ref={containerRef}
                 style={{ width: width || '100%', height: height || 500 }}>
                <div className="pdf-config-panel">
                    <h2>📄 Document Viewer</h2>
                    <p className="subtitle">{isTestMode ? 'Test Mode - Load a PDF to preview' : 'Select which document to display'}</p>

                    {!isTestMode && (
                        <div className="pdf-config-detected">
                            <div className="pdf-config-detected-item">
                                <div className="label">Table</div>
                                <div className={`value ${viewerConfig?.tableName ? 'success' : 'error'}`}>
                                    {viewerConfig?.tableName || 'Not detected'}
                                </div>
                            </div>
                            <div className="pdf-config-detected-item">
                                <div className="label">Record</div>
                                <div className={`value ${viewerConfig?.recordId ? 'success' : 'error'}`}>
                                    {viewerConfig?.recordId ? viewerConfig.recordId.substring(0, 8) + '...' : 'Not detected'}
                                </div>
                            </div>
                        </div>
                    )}

                    {isTestMode ? (
                        <>
                            {/* Test Mode - URL Input */}
                            <div className="pdf-config-section">
                                <label>Load from URL</label>
                                <input
                                    type="text"
                                    className="pdf-config-select"
                                    value={testUrl}
                                    onChange={(e) => setTestUrl(e.target.value)}
                                    placeholder="Enter PDF URL..."
                                    style={{ marginBottom: 8 }}
                                />
                                <button
                                    className="pdf-config-btn"
                                    onClick={() => loadFromUrl(testUrl)}
                                    disabled={!testUrl}
                                >
                                    Load from URL
                                </button>
                            </div>

                            {/* Test Mode - File Upload */}
                            <div className="pdf-config-section" style={{ marginTop: 20 }}>
                                <label>Or upload a file</label>
                                <input
                                    type="file"
                                    ref={fileInputRef}
                                    onChange={handleFileSelect}
                                    accept=".pdf,.jpg,.jpeg,.png,.gif,.bmp,.webp"
                                    style={{ display: 'none' }}
                                />
                                <button
                                    className="pdf-config-btn"
                                    onClick={() => fileInputRef.current?.click()}
                                    style={{ background: 'transparent', color: 'var(--pdf-primary-color)', border: '1px solid var(--pdf-primary-color)' }}
                                >
                                    Choose File (PDF or Image)
                                </button>
                            </div>

                            <div className="pdf-config-info" style={{ marginTop: 20 }}>
                                <strong>Test Mode Active</strong>
                                <p style={{ margin: '4px 0 0 0', fontSize: 12 }}>
                                    No Dataverse context detected. When deployed to Power Apps, the control will automatically load files from Dataverse.
                                </p>
                            </div>
                        </>
                    ) : columnsWithFiles.length > 0 ? (
                        <>
                            <div className="pdf-config-section">
                                <label>Select File Column</label>
                                <select
                                    className="pdf-config-select"
                                    value={selectedColumn || ''}
                                    onChange={(e) => setSelectedColumn(e.target.value)}
                                >
                                    <option value="">-- Choose a column --</option>
                                    {viewerConfig?.fileColumns.map((col) => (
                                        <option
                                            key={col.logicalName}
                                            value={col.logicalName}
                                            disabled={!col.hasFile}
                                        >
                                            {col.displayName} {col.hasFile ? '✓' : '(empty)'}
                                            {col.logicalName === defaultFileColumn ? ' ⭐' : ''}
                                        </option>
                                    ))}
                                </select>
                                <div className="hint">Columns marked with ✓ have files attached</div>
                            </div>
                            <button
                                className="pdf-config-btn"
                                disabled={!selectedColumn}
                                onClick={() => selectedColumn && loadDocument(selectedColumn)}
                            >
                                Load Document
                            </button>
                        </>
                    ) : (
                        <div className="pdf-no-files-message">
                            <p>📭 No files are attached to this record.</p>
                        </div>
                    )}
                </div>
            </div>
        );
    }

    // Render error state
    if (viewerState === 'error') {
        return (
            <div className="pdf-viewer-container" data-theme={darkMode ? 'dark' : 'light'} ref={containerRef}
                 style={{ width: width || '100%', height: height || 500 }}>
                <div className="pdf-error-container">
                    <div className="pdf-error-icon">⚠️</div>
                    <div className="pdf-error-message">{errorMessage}</div>
                    <button className="pdf-config-btn" style={{ maxWidth: 200, margin: '0 auto' }}
                            onClick={() => setViewerState('config')}>
                        Try Again
                    </button>
                </div>
            </div>
        );
    }

    // Render viewer state
    const zoomPercent = Math.round(currentScale * 100);
    const columnsWithFiles = viewerConfig?.fileColumns.filter(c => c.hasFile) || [];

    return (
        <div className="pdf-viewer-container" data-theme={darkMode ? 'dark' : 'light'} ref={containerRef}
             style={{ width: width || '100%', height: height || 500 }}>

            {/* Toolbar */}
            {showToolbar && (
                <div className="pdf-toolbar">
                    {/* Sidebar Toggle */}
                    <div className="pdf-toolbar-group">
                        <button
                            className={`pdf-btn icon-btn toggle-btn ${sidebarVisible ? 'active' : ''}`}
                            onClick={toggleSidebar}
                            title="Toggle sidebar (B)"
                        >
                            ☰
                        </button>
                    </div>

                    {/* Column Selector */}
                    {columnsWithFiles.length > 1 && (
                        <div className="pdf-toolbar-group">
                            <select
                                className="pdf-select pdf-column-selector"
                                value={selectedColumn || ''}
                                onChange={(e) => loadDocument(e.target.value)}
                                title="Select file column"
                            >
                                {columnsWithFiles.map((col) => (
                                    <option key={col.logicalName} value={col.logicalName}>
                                        {col.displayName}
                                    </option>
                                ))}
                            </select>
                        </div>
                    )}

                    {/* Zoom Controls */}
                    <div className="pdf-toolbar-group">
                        <button className="pdf-btn icon-btn" onClick={zoomOut} title="Zoom out (-)">−</button>
                        <select
                            className="pdf-select"
                            value={fitMode || currentScale.toString()}
                            onChange={(e) => setZoom(e.target.value)}
                            title="Zoom level"
                        >
                            <option value="auto">Auto</option>
                            <option value="page-fit">Fit Page</option>
                            <option value="page-width">Fit Width</option>
                            <option value="0.5">50%</option>
                            <option value="0.75">75%</option>
                            <option value="1">100%</option>
                            <option value="1.25">125%</option>
                            <option value="1.5">150%</option>
                            <option value="2">200%</option>
                        </select>
                        <button className="pdf-btn icon-btn" onClick={zoomIn} title="Zoom in (+)">+</button>
                    </div>

                    {/* Page Navigation */}
                    {fileType === 'pdf' && (
                        <div className="pdf-toolbar-group pdf-page-nav">
                            <button
                                className="pdf-btn icon-btn"
                                onClick={previousPage}
                                disabled={currentPage <= 1}
                                title="Previous page"
                            >
                                ◀
                            </button>
                            <input
                                type="number"
                                className="pdf-input page-input"
                                value={currentPage}
                                min={1}
                                max={totalPages}
                                onChange={(e) => goToPage(parseInt(e.target.value, 10))}
                            />
                            <span>of {totalPages}</span>
                            <button
                                className="pdf-btn icon-btn"
                                onClick={nextPage}
                                disabled={currentPage >= totalPages}
                                title="Next page"
                            >
                                ▶
                            </button>
                        </div>
                    )}

                    {/* Rotation */}
                    <div className="pdf-toolbar-group">
                        <button className="pdf-btn icon-btn" onClick={rotateLeft} title="Rotate left">↶</button>
                        <button className="pdf-btn icon-btn" onClick={rotateRight} title="Rotate right (R)">↷</button>
                    </div>

                    {/* Search Button */}
                    {fileType === 'pdf' && (
                        <div className="pdf-toolbar-group">
                            <button
                                className={`pdf-btn icon-btn ${showFindPanel ? 'active' : ''}`}
                                onClick={() => showFindPanel ? closeFindPanel() : openFindPanel()}
                                title="Find (Ctrl+F)"
                            >
                                🔍
                            </button>
                        </div>
                    )}

                    {/* Action Buttons */}
                    <div className="pdf-toolbar-group">
                        <button
                            className={`pdf-btn icon-btn toggle-btn ${darkMode ? 'active' : ''}`}
                            onClick={toggleDarkMode}
                            title="Dark mode (D)"
                        >
                            🌙
                        </button>
                        <button className="pdf-btn icon-btn" onClick={handlePrint} title="Print">🖨️</button>
                        <button className="pdf-btn icon-btn" onClick={toggleFullscreen} title="Fullscreen (F)">⛶</button>
                        <button className="pdf-btn icon-btn" onClick={() => setShowDocInfo(true)} title="Document info">ℹ️</button>
                        <button className="pdf-btn icon-btn" onClick={handleDownload} title="Download">📥</button>
                        <button className="pdf-btn icon-btn" onClick={() => setShowShortcuts(true)} title="Shortcuts (?)">⌨️</button>
                    </div>
                </div>
            )}

            {/* Floating Find Panel */}
            {showFindPanel && fileType === 'pdf' && (
                <div className="pdf-find-panel">
                    <div className="pdf-find-input-wrapper">
                        <input
                            ref={searchInputRef}
                            type="text"
                            className="pdf-find-input"
                            placeholder="Find in document..."
                            value={searchText}
                            onChange={(e) => handleSearch(e.target.value)}
                            autoFocus
                        />
                        {searchText && (
                            <span className="pdf-find-count">
                                {searchMatches.length > 0
                                    ? `${currentMatchIndex + 1} of ${searchMatches.length}`
                                    : 'No matches'}
                            </span>
                        )}
                    </div>
                    <div className="pdf-find-buttons">
                        <button
                            className="pdf-find-btn"
                            onClick={searchPrev}
                            disabled={searchMatches.length === 0}
                            title="Previous match (Shift+Enter)"
                        >
                            ▲
                        </button>
                        <button
                            className="pdf-find-btn"
                            onClick={searchNext}
                            disabled={searchMatches.length === 0}
                            title="Next match (Enter)"
                        >
                            ▼
                        </button>
                        <button
                            className="pdf-find-btn pdf-find-close"
                            onClick={closeFindPanel}
                            title="Close (Esc)"
                        >
                            ✕
                        </button>
                    </div>
                </div>
            )}

            {/* Main Content */}
            <div className="pdf-main-content">
                {/* Sidebar */}
                {showSidebar && (
                    <div className={`pdf-sidebar ${!sidebarVisible ? 'collapsed' : ''}`}>
                        <div className="pdf-sidebar-tabs">
                            <button
                                className={`pdf-sidebar-tab ${sidebarTab === 'thumbnails' ? 'active' : ''}`}
                                onClick={() => setSidebarTab('thumbnails')}
                            >
                                Thumbnails
                            </button>
                            <button
                                className={`pdf-sidebar-tab ${sidebarTab === 'outline' ? 'active' : ''}`}
                                onClick={() => setSidebarTab('outline')}
                            >
                                Outline
                            </button>
                        </div>
                        <div className="pdf-sidebar-content">
                            {sidebarTab === 'thumbnails' && (
                                <div className="pdf-sidebar-panel active">
                                    <ThumbnailPanel
                                        key={documentKey}
                                        pdfService={pdfServiceRef.current}
                                        totalPages={totalPages}
                                        currentPage={currentPage}
                                        onPageClick={goToPage}
                                    />
                                </div>
                            )}
                            {sidebarTab === 'outline' && (
                                <div className="pdf-sidebar-panel active">
                                    <OutlinePanel
                                        outline={outline}
                                        pdfService={pdfServiceRef.current}
                                        onNavigate={goToPage}
                                    />
                                </div>
                            )}
                        </div>
                    </div>
                )}

                {/* Document Canvas */}
                <div className="pdf-canvas-container" ref={canvasContainerRef}>
                    {fileType === 'pdf' ? (
                        <PdfCanvas
                            key={documentKey}
                            pdfService={pdfServiceRef.current}
                            pageViewports={pageViewports}
                            scale={currentScale}
                            rotation={currentRotation}
                            currentPage={currentPage}
                            onPageChange={setCurrentPage}
                            renderedPages={renderedPages}
                            setRenderedPages={setRenderedPages}
                            textContent={textContent}
                            setTextContent={setTextContent}
                            searchText={searchText}
                            currentMatchIndex={currentMatchIndex}
                            searchMatches={searchMatches}
                        />
                    ) : (
                        <div className="pdf-image-viewer">
                            <img
                                src={imageUrl || ''}
                                alt="Document"
                                style={{
                                    transform: `scale(${currentScale}) rotate(${currentRotation}deg)`,
                                    transformOrigin: 'center center'
                                }}
                            />
                        </div>
                    )}
                </div>
            </div>

            {/* Document Info Panel */}
            {showDocInfo && (
                <div className="pdf-doc-info-panel">
                    <h3>📄 Document Information</h3>
                    <div className="pdf-doc-info-row">
                        <span className="label">File Column</span>
                        <span className="value">{selectedColumn}</span>
                    </div>
                    <div className="pdf-doc-info-row">
                        <span className="label">Pages</span>
                        <span className="value">{totalPages}</span>
                    </div>
                    {metadata?.title && (
                        <div className="pdf-doc-info-row">
                            <span className="label">Title</span>
                            <span className="value">{metadata.title}</span>
                        </div>
                    )}
                    {metadata?.author && (
                        <div className="pdf-doc-info-row">
                            <span className="label">Author</span>
                            <span className="value">{metadata.author}</span>
                        </div>
                    )}
                    <button
                        className="pdf-btn"
                        style={{ marginTop: 12 }}
                        onClick={() => setShowDocInfo(false)}
                    >
                        Close
                    </button>
                </div>
            )}

            {/* Shortcuts Modal */}
            {showShortcuts && (
                <div className="pdf-modal-overlay" onClick={() => setShowShortcuts(false)}>
                    <div className="pdf-modal-content" onClick={(e) => e.stopPropagation()}>
                        <div className="pdf-modal-header">
                            <h2>Keyboard Shortcuts</h2>
                            <button className="pdf-modal-close" onClick={() => setShowShortcuts(false)}>×</button>
                        </div>
                        <div className="pdf-shortcut-list">
                            <div className="pdf-shortcut-key"><kbd>Ctrl</kbd>+<kbd>F</kbd></div>
                            <div className="pdf-shortcut-desc">Find in document</div>

                            <div className="pdf-shortcut-key"><kbd>Ctrl</kbd>+<kbd>P</kbd></div>
                            <div className="pdf-shortcut-desc">Print</div>

                            <div className="pdf-shortcut-key"><kbd>Arrow Left/Right</kbd></div>
                            <div className="pdf-shortcut-desc">Previous / Next page</div>

                            <div className="pdf-shortcut-key"><kbd>Home</kbd> / <kbd>End</kbd></div>
                            <div className="pdf-shortcut-desc">First / Last page</div>

                            <div className="pdf-shortcut-key"><kbd>Esc</kbd></div>
                            <div className="pdf-shortcut-desc">Close dialogs</div>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

// Thumbnail Panel Component
interface ThumbnailPanelProps {
    pdfService: PdfService;
    totalPages: number;
    currentPage: number;
    onPageClick: (page: number) => void;
}

const ThumbnailPanel: React.FC<ThumbnailPanelProps> = ({ pdfService, totalPages, currentPage, onPageClick }) => {
    const [thumbnails, setThumbnails] = useState<Map<number, string>>(new Map());
    const thumbnailTasksRef = useRef<Map<number, RenderTask>>(new Map());
    const mountedRef = useRef(true);

    useEffect(() => {
        mountedRef.current = true;

        const renderThumbnails = async () => {
            // Render visible thumbnails first (around current page)
            const start = Math.max(1, currentPage - 3);
            const end = Math.min(totalPages, currentPage + 10);

            // First batch: pages near current view
            for (let i = start; i <= end && mountedRef.current; i++) {
                if (thumbnails.has(i)) continue;
                try {
                    const canvas = document.createElement('canvas');
                    const renderTask = await pdfService.renderPage(i, canvas, CONFIG.THUMBNAIL_SCALE, 0);
                    thumbnailTasksRef.current.set(i, renderTask);
                    await renderTask.promise;
                    thumbnailTasksRef.current.delete(i);
                    if (mountedRef.current) {
                        setThumbnails(prev => new Map(prev).set(i, canvas.toDataURL()));
                    }
                } catch {
                    // Skip failed thumbnails
                }
            }

            // Second batch: remaining pages (deferred, slower)
            await new Promise(resolve => setTimeout(resolve, 500));
            for (let i = 1; i <= Math.min(totalPages, 30) && mountedRef.current; i++) {
                if (thumbnails.has(i)) continue;
                try {
                    const canvas = document.createElement('canvas');
                    const renderTask = await pdfService.renderPage(i, canvas, CONFIG.THUMBNAIL_SCALE, 0);
                    thumbnailTasksRef.current.set(i, renderTask);
                    await renderTask.promise;
                    thumbnailTasksRef.current.delete(i);
                    if (mountedRef.current) {
                        setThumbnails(prev => new Map(prev).set(i, canvas.toDataURL()));
                    }
                    // Small delay between thumbnails to not block main thread
                    await new Promise(resolve => setTimeout(resolve, 50));
                } catch {
                    // Skip failed thumbnails
                }
            }
        };

        if (totalPages > 0) {
            // Defer thumbnail rendering to not block initial PDF render
            const timer = setTimeout(renderThumbnails, 300);
            return () => {
                clearTimeout(timer);
                mountedRef.current = false;
                thumbnailTasksRef.current.forEach((task) => task.cancel());
                thumbnailTasksRef.current.clear();
            };
        }

        return () => {
            mountedRef.current = false;
            thumbnailTasksRef.current.forEach((task) => task.cancel());
            thumbnailTasksRef.current.clear();
        };
    }, [pdfService, totalPages, currentPage]);

    return (
        <div className="pdf-thumbnail-container">
            {Array.from({ length: totalPages }, (_, i) => i + 1).map((pageNum) => (
                <div
                    key={pageNum}
                    className={`pdf-thumbnail-item ${pageNum === currentPage ? 'active' : ''}`}
                    onClick={() => onPageClick(pageNum)}
                >
                    {thumbnails.get(pageNum) ? (
                        <img src={thumbnails.get(pageNum)} alt={`Page ${pageNum}`} />
                    ) : (
                        <div className="pdf-page-placeholder" style={{ width: 150, height: 200 }} />
                    )}
                    <div className="pdf-thumbnail-label">Page {pageNum}</div>
                </div>
            ))}
        </div>
    );
};

// Outline Panel Component
interface OutlinePanelProps {
    outline: OutlineItem[];
    pdfService: PdfService;
    onNavigate: (page: number) => void;
}

const OutlinePanel: React.FC<OutlinePanelProps> = ({ outline, pdfService, onNavigate }) => {
    const handleClick = async (item: OutlineItem) => {
        const pageIndex = await pdfService.getPageIndexFromDest(item.dest);
        onNavigate(pageIndex + 1);
    };

    const renderItems = (items: OutlineItem[], level = 1): React.ReactNode => {
        return items.map((item, index) => (
            <React.Fragment key={index}>
                <div
                    className={`pdf-outline-item level-${Math.min(level, 3)}`}
                    onClick={() => handleClick(item)}
                >
                    {item.title}
                </div>
                {item.items && renderItems(item.items, level + 1)}
            </React.Fragment>
        ));
    };

    if (outline.length === 0) {
        return <div style={{ padding: 16, color: 'var(--pdf-text-muted)', textAlign: 'center' }}>No outline available</div>;
    }

    return <>{renderItems(outline)}</>;
};

// PDF Canvas Component - wrapped in React.memo for performance
interface PdfCanvasProps {
    pdfService: PdfService;
    pageViewports: PageViewport[];
    scale: number;
    rotation: number;
    currentPage: number;
    onPageChange: (page: number) => void;
    renderedPages: Map<number, 'rendering' | 'done'>;
    setRenderedPages: React.Dispatch<React.SetStateAction<Map<number, 'rendering' | 'done'>>>;
    textContent: Map<number, TextContent>;
    setTextContent: React.Dispatch<React.SetStateAction<Map<number, TextContent>>>;
    searchText: string;
    currentMatchIndex: number;
    searchMatches: { pageNum: number; itemIndex: number; startOffset: number; length: number }[];
}

const PdfCanvas: React.FC<PdfCanvasProps> = React.memo(({
    pdfService,
    pageViewports,
    scale,
    rotation,
    currentPage,
    onPageChange,
    renderedPages,
    setRenderedPages,
    textContent,
    setTextContent,
    searchText,
    currentMatchIndex,
    searchMatches
}) => {
    const containerRef = useRef<HTMLDivElement>(null);
    const pageRefs = useRef<Map<number, HTMLDivElement>>(new Map());
    // Track active render tasks for cancellation
    const renderTasksRef = useRef<Map<number, RenderTask>>(new Map());

    // Render visible pages
    const renderVisiblePages = useCallback(async () => {
        if (!containerRef.current) return;

        const container = containerRef.current;
        const scrollTop = container.scrollTop;
        const viewportTop = scrollTop - CONFIG.RENDER_BUFFER_PX;
        const viewportBottom = scrollTop + container.clientHeight + CONFIG.RENDER_BUFFER_PX;

        for (let i = 1; i <= pageViewports.length; i++) {
            const pageRef = pageRefs.current.get(i);
            if (!pageRef) continue;

            const top = pageRef.offsetTop;
            const bottom = top + pageRef.offsetHeight;

            if (bottom >= viewportTop && top <= viewportBottom) {
                if (!renderedPages.has(i)) {
                    setRenderedPages(prev => new Map(prev).set(i, 'rendering'));

                    try {
                        const canvas = pageRef.querySelector('canvas');
                        if (canvas) {
                            // Cancel any existing render task for this page
                            const existingTask = renderTasksRef.current.get(i);
                            if (existingTask) {
                                existingTask.cancel();
                                renderTasksRef.current.delete(i);
                            }

                            // Start new render and store the task
                            const renderTask = await pdfService.renderPage(i, canvas, scale, rotation);
                            renderTasksRef.current.set(i, renderTask);

                            // Wait for render to complete
                            await renderTask.promise;

                            // Clean up completed task
                            renderTasksRef.current.delete(i);

                            // Get text content
                            const text = await pdfService.getTextContent(i);
                            setTextContent(prev => new Map(prev).set(i, text));

                            setRenderedPages(prev => new Map(prev).set(i, 'done'));
                        }
                    } catch (e) {
                        // Handle RenderingCancelledException - expected when cancelled
                        if (e instanceof Error && e.message.includes('Rendering cancelled')) {
                            // This is expected when a render was cancelled, just return
                            return;
                        }
                        console.error(`Error rendering page ${i}:`, e);
                        setRenderedPages(prev => {
                            const next = new Map(prev);
                            next.delete(i);
                            return next;
                        });
                    }
                }
            }
        }
    }, [pdfService, pageViewports, scale, rotation, renderedPages, setRenderedPages, setTextContent]);

    // Cancel all render tasks when scale/rotation changes or on unmount
    useEffect(() => {
        return () => {
            // Cancel all active render tasks on cleanup
            renderTasksRef.current.forEach((task) => {
                task.cancel();
            });
            renderTasksRef.current.clear();
        };
    }, [scale, rotation]);

    // Scroll to top when document loads (pageViewports changes)
    useEffect(() => {
        if (containerRef.current && pageViewports.length > 0) {
            containerRef.current.scrollTo(0, 0);
        }
    }, [pageViewports.length]);

    // Initial render and scroll handler
    useEffect(() => {
        renderVisiblePages();

        const handleScroll = () => {
            renderVisiblePages();

            // Update current page based on scroll position
            if (!containerRef.current) return;
            const container = containerRef.current;
            const scrollTop = container.scrollTop;
            const containerHeight = container.clientHeight;

            let mostVisible = 1;
            let maxVisible = 0;

            pageRefs.current.forEach((pageRef, pageNum) => {
                const top = pageRef.offsetTop;
                const bottom = top + pageRef.offsetHeight;
                const visibleTop = Math.max(scrollTop, top);
                const visibleBottom = Math.min(scrollTop + containerHeight, bottom);
                const visible = Math.max(0, visibleBottom - visibleTop);

                if (visible > maxVisible) {
                    maxVisible = visible;
                    mostVisible = pageNum;
                }
            });

            if (mostVisible !== currentPage) {
                onPageChange(mostVisible);
            }
        };

        const container = containerRef.current;
        container?.addEventListener('scroll', handleScroll);
        return () => container?.removeEventListener('scroll', handleScroll);
    }, [renderVisiblePages, currentPage, onPageChange]);

    // Get matches for a specific page
    const getPageMatches = useCallback((pageNum: number) => {
        return searchMatches.filter(m => m.pageNum === pageNum);
    }, [searchMatches]);

    // Check if a match is the current match
    const isCurrentMatch = useCallback((pageNum: number, itemIndex: number) => {
        if (currentMatchIndex < 0 || currentMatchIndex >= searchMatches.length) return false;
        const current = searchMatches[currentMatchIndex];
        return current.pageNum === pageNum && current.itemIndex === itemIndex;
    }, [searchMatches, currentMatchIndex]);

    return (
        <div className="pdf-page-wrapper-container" ref={containerRef}>
            {pageViewports.map((viewport, index) => {
                const pageNum = index + 1;
                const isRotated = rotation % 180 !== 0;
                const width = (isRotated ? viewport.height : viewport.width) * scale;
                const height = (isRotated ? viewport.width : viewport.height) * scale;
                const pageTextContent = textContent.get(pageNum);
                const pageMatches = getPageMatches(pageNum);

                return (
                    <div
                        key={pageNum}
                        className="pdf-page-wrapper"
                        data-page={pageNum}
                        ref={(el) => {
                            if (el) pageRefs.current.set(pageNum, el);
                        }}
                        style={{ width, height }}
                    >
                        {/* Canvas must always exist for rendering, placeholder shows while loading */}
                        <canvas
                            className="pdf-page-canvas"
                            style={{
                                display: renderedPages.get(pageNum) === 'done' ? 'block' : 'none'
                            }}
                        />

                        {/* Text Layer with Search Highlights */}
                        {renderedPages.get(pageNum) === 'done' && pageTextContent && searchText && (
                            <div className="pdf-text-layer" style={{ width, height }}>
                                {pageTextContent.items.map((item, itemIndex) => {
                                    const matchesInItem = pageMatches.filter(m => m.itemIndex === itemIndex);
                                    if (matchesInItem.length === 0) return null;

                                    // Get transform from PDF coordinates
                                    const transform = item.transform;
                                    const fontSize = Math.sqrt(transform[0] * transform[0] + transform[1] * transform[1]);
                                    const left = transform[4] * scale;
                                    const bottom = transform[5] * scale;
                                    const top = height - bottom - (fontSize * scale);

                                    return matchesInItem.map((match, matchIdx) => {
                                        const isCurrent = isCurrentMatch(pageNum, itemIndex);
                                        return (
                                            <span
                                                key={`${itemIndex}-${matchIdx}`}
                                                className={`pdf-text-highlight ${isCurrent ? 'current' : ''}`}
                                                data-item={itemIndex}
                                                style={{
                                                    left: `${left}px`,
                                                    top: `${top}px`,
                                                    fontSize: `${fontSize * scale}px`,
                                                    padding: '2px 4px',
                                                }}
                                            >
                                                {item.str.substring(match.startOffset, match.startOffset + match.length)}
                                            </span>
                                        );
                                    });
                                })}
                            </div>
                        )}

                        {renderedPages.get(pageNum) !== 'done' && (
                            <div className="pdf-page-placeholder" style={{ width, height }}>
                                {renderedPages.get(pageNum) === 'rendering' && (
                                    <div className="pdf-page-loading">Loading...</div>
                                )}
                            </div>
                        )}
                    </div>
                );
            })}
        </div>
    );
});

export default PdfViewer;
