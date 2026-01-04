/**
 * PdfViewer - Main React component for the PDF Viewer PCF control
 */

import * as React from 'react';
import { useState, useEffect, useRef, useCallback, useMemo } from 'react';
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
    // Thumbnail rendering
    THUMBNAIL_PREFETCH_BEFORE: 3,
    THUMBNAIL_PREFETCH_AFTER: 10,
    THUMBNAIL_MAX_PAGES: 30,
    THUMBNAIL_BATCH_DELAY_MS: 50,
    THUMBNAIL_INITIAL_DELAY_MS: 300,
    THUMBNAIL_DEFERRED_DELAY_MS: 500,
    // UI delays
    FOCUS_DELAY_MS: 100,
    // Render task management
    RENDER_TASK_TIMEOUT_MS: 10000,
    // Virtual scrolling - reduces DOM nodes for large documents
    VIRTUAL_SCROLL_BUFFER_BEFORE: 5,  // Pages to render before current
    VIRTUAL_SCROLL_BUFFER_AFTER: 8,   // Pages to render after current
    VIRTUAL_SCROLL_THRESHOLD: 30,     // Only use virtual scrolling for documents with more pages
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

    // Page input state (separate from currentPage to prevent glitches while typing)
    const [pageInputValue, setPageInputValue] = useState('');

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

    // Navigation synchronization refs - prevent race conditions between state updates and scroll
    const isNavigatingRef = useRef(false);
    const pendingScrollRef = useRef<number | null>(null);

    // Memoized computed values
    const columnsWithFiles = useMemo(() =>
        viewerConfig?.fileColumns.filter(c => c.hasFile) || [],
        [viewerConfig?.fileColumns]
    );

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
            dataverseServiceRef.current.abort();
        };
    }, [defaultFileColumn]);

    // Notify parent of page changes
    useEffect(() => {
        onPageChange?.(currentPage, totalPages);
    }, [currentPage, totalPages, onPageChange]);

    // Sync page input value when currentPage changes (from scroll, thumbnail click, etc.)
    useEffect(() => {
        setPageInputValue(currentPage.toString());
    }, [currentPage]);

    // Handle pending scroll after currentPage updates and virtual window re-renders
    // This useEffect runs AFTER React re-renders with the new currentPage,
    // ensuring the target page element exists in the DOM
    useEffect(() => {
        if (pendingScrollRef.current === null) return;

        const targetPage = pendingScrollRef.current;

        // Use requestAnimationFrame to ensure DOM has updated
        requestAnimationFrame(() => {
            const container = canvasContainerRef.current?.querySelector('.pdf-page-wrapper-container');
            const pageElement = canvasContainerRef.current?.querySelector<HTMLElement>(`[data-page="${targetPage}"]`);

            if (container && pageElement) {
                // Use instant scroll for programmatic navigation (no smooth animation)
                container.scrollTo({
                    top: pageElement.offsetTop - 20,
                    behavior: 'instant'
                });
            }

            pendingScrollRef.current = null;

            // Clear navigation flag after scroll settles
            setTimeout(() => {
                isNavigatingRef.current = false;
            }, 100);
        });
    }, [currentPage]);

    // Cleanup blob URLs to prevent memory leaks
    useEffect(() => {
        return () => {
            if (imageUrl && imageUrl.startsWith('blob:')) {
                URL.revokeObjectURL(imageUrl);
            }
        };
    }, [imageUrl]);

    // Load ID ref to prevent stale callbacks from updating state during rapid document switches
    const loadIdRef = useRef(0);

    // Load document from selected column
    const loadDocument = useCallback(async (columnName: string) => {
        // Increment load ID to prevent stale callbacks from updating state
        const thisLoadId = ++loadIdRef.current;

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

        // Abort if a newer load has started
        if (loadIdRef.current !== thisLoadId) return;

        try {
            const fileUrl = dataverseService.getFileUrl(columnName);

            // Skip HEAD request - detect file type from column config instead (200-500ms faster)
            // Check if this is an image column based on metadata
            if (isImageFile(columnName)) {
                // Load as image
                if (loadIdRef.current !== thisLoadId) return;
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
                    // Only update progress if this is still the current load
                    if (loadIdRef.current === thisLoadId && total > 0) {
                        setLoadingProgress((loaded / total) * 100);
                    }
                });

                // Abort if a newer load has started
                if (loadIdRef.current !== thisLoadId) return;

                const pageCount = pdfService.getPageCount();
                setTotalPages(pageCount);
                setCurrentPage(1);

                // Get first page viewport and show immediately
                const firstViewport = await pdfService.getPageViewport(1, 1);
                if (loadIdRef.current !== thisLoadId) return;
                setPageViewports([firstViewport]);

                // Show viewer NOW - don't wait for anything else
                setViewerState('viewing');

                // Load ALL remaining viewports in a single batch after render
                requestAnimationFrame(() => {
                    // Check if still current load before starting async work
                    if (loadIdRef.current !== thisLoadId) return;
                    (async () => {
                        const allViewports: PageViewport[] = [firstViewport];
                        for (let i = 2; i <= pageCount; i++) {
                            if (loadIdRef.current !== thisLoadId) return;
                            const vp = await pdfService.getPageViewport(i, 1);
                            allViewports.push(vp);
                        }
                        if (loadIdRef.current !== thisLoadId) return;
                        setPageViewports(allViewports);

                        // Load metadata after viewports
                        if (loadIdRef.current === thisLoadId) {
                            pdfService.getMetadata().then(meta => {
                                if (loadIdRef.current === thisLoadId) setMetadata(meta);
                                return meta;
                            }).catch(() => { /* ignore */ });
                            pdfService.getOutline().then(out => {
                                if (loadIdRef.current === thisLoadId) setOutline(out);
                                return out;
                            }).catch(() => { /* ignore */ });
                        }
                    })();
                });
            }
        } catch (error) {
            // Only show error if this is still the current load
            if (loadIdRef.current !== thisLoadId) return;
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
        } catch {
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
    const zoomIn = useCallback(() => {
        applyZoom(currentScale * CONFIG.ZOOM_STEP);
    }, [applyZoom, currentScale]);

    const zoomOut = useCallback(() => {
        applyZoom(currentScale / CONFIG.ZOOM_STEP);
    }, [applyZoom, currentScale]);

    const setZoom = useCallback((value: string) => {
        if (['auto', 'page-fit', 'page-width'].includes(value)) {
            applyZoom(calculateScale(value), value);
        } else {
            applyZoom(parseFloat(value));
        }
    }, [applyZoom, calculateScale]);

    // Navigation handlers - use pending scroll to wait for virtual window re-render
    const goToPage = useCallback((page: number) => {
        const targetPage = Math.max(1, Math.min(totalPages, page));

        // Set navigation flag to prevent IntersectionObserver from overriding
        isNavigatingRef.current = true;
        pendingScrollRef.current = targetPage;

        // Set state - this will trigger virtual window recalculation
        setCurrentPage(targetPage);

        // Note: Actual scroll happens in useEffect after re-render ensures page is in DOM
    }, [totalPages]);

    const previousPage = useCallback(() => goToPage(currentPage - 1), [goToPage, currentPage]);
    const nextPage = useCallback(() => goToPage(currentPage + 1), [goToPage, currentPage]);

    // Handle page input change - only allow valid numbers up to totalPages digit length
    const handlePageInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
        const value = e.target.value;
        // Only allow digits, limit to reasonable length (max digits in totalPages + 1)
        const maxDigits = Math.max(totalPages.toString().length + 1, 4);
        if (value === '' || (/^\d+$/.test(value) && value.length <= maxDigits)) {
            setPageInputValue(value);
        }
    }, [totalPages]);

    // Handle page input blur - navigate to the entered page
    const handlePageInputBlur = useCallback(() => {
        const page = parseInt(pageInputValue, 10);
        if (!isNaN(page) && page >= 1 && page <= totalPages) {
            goToPage(page);
        } else {
            // Reset to current page if invalid
            setPageInputValue(currentPage.toString());
        }
    }, [pageInputValue, totalPages, goToPage, currentPage]);

    // Handle page input key press - navigate on Enter
    const handlePageInputKeyDown = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            (e.target as HTMLInputElement).blur();
        } else if (e.key === 'Escape') {
            e.preventDefault();
            setPageInputValue(currentPage.toString());
            (e.target as HTMLInputElement).blur();
        }
    }, [currentPage]);

    // Rotation handlers
    const rotateLeft = useCallback(() => {
        setCurrentRotation((prev) => (prev - 90 + 360) % 360);
        setRenderedPages(new Map());
        setTextContent(new Map());
    }, []);

    const rotateRight = useCallback(() => {
        setCurrentRotation((prev) => (prev + 90) % 360);
        setRenderedPages(new Map());
        setTextContent(new Map());
    }, []);

    // Toggle handlers
    const toggleSidebar = useCallback(() => setSidebarVisible(prev => !prev), []);
    const toggleDarkMode = useCallback(() => setDarkMode(prev => !prev), []);
    const toggleFullscreen = useCallback(() => {
        if (!document.fullscreenElement) {
            containerRef.current?.requestFullscreen();
        } else {
            document.exitFullscreen();
        }
    }, []);

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
        setTimeout(() => searchInputRef.current?.focus(), CONFIG.FOCUS_DELAY_MS);
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
        } catch {
            // Download failed silently - user will notice
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
                                type="text"
                                inputMode="numeric"
                                className="pdf-input page-input"
                                value={pageInputValue}
                                onChange={handlePageInputChange}
                                onBlur={handlePageInputBlur}
                                onKeyDown={handlePageInputKeyDown}
                                onFocus={(e) => e.target.select()}
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
                            isNavigatingRef={isNavigatingRef}
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

const ThumbnailPanel: React.FC<ThumbnailPanelProps> = React.memo(({ pdfService, totalPages, currentPage, onPageClick }) => {
    const [thumbnails, setThumbnails] = useState<Map<number, string>>(new Map());
    const thumbnailTasksRef = useRef<Map<number, RenderTask>>(new Map());
    const mountedRef = useRef(true);
    const idleCallbackRef = useRef<number | null>(null);
    const containerRef = useRef<HTMLDivElement>(null);
    const [visibleRange, setVisibleRange] = useState({ start: 1, end: 20 });

    // Helper to schedule work during idle time (with fallback for older browsers)
    const scheduleIdleWork = useCallback((callback: () => void): number => {
        if (typeof window.requestIdleCallback === 'function') {
            return window.requestIdleCallback(callback, { timeout: 2000 });
        }
        // Fallback for browsers without requestIdleCallback (Safari, older browsers)
        return setTimeout(callback, CONFIG.THUMBNAIL_BATCH_DELAY_MS) as unknown as number;
    }, []);

    const cancelIdleWork = useCallback((id: number) => {
        if (typeof window.cancelIdleCallback === 'function') {
            window.cancelIdleCallback(id);
        } else {
            clearTimeout(id);
        }
    }, []);

    // Track visible range for thumbnail virtualization
    useEffect(() => {
        const container = containerRef.current;
        if (!container || totalPages <= 30) return; // Only virtualize for large documents

        const updateVisibleRange = () => {
            const scrollTop = container.scrollTop;
            const containerHeight = container.clientHeight;
            const itemHeight = 220; // Approximate height of thumbnail item (200px + margin)

            const startIndex = Math.floor(scrollTop / itemHeight);
            const visibleCount = Math.ceil(containerHeight / itemHeight);

            // Add buffer pages before and after
            const start = Math.max(1, startIndex - 5);
            const end = Math.min(totalPages, startIndex + visibleCount + 10);

            setVisibleRange({ start, end });
        };

        // Throttled scroll handler
        let ticking = false;
        const handleScroll = () => {
            if (!ticking) {
                requestAnimationFrame(() => {
                    updateVisibleRange();
                    ticking = false;
                });
                ticking = true;
            }
        };

        container.addEventListener('scroll', handleScroll, { passive: true });
        updateVisibleRange(); // Initial calculation

        return () => container.removeEventListener('scroll', handleScroll);
    }, [totalPages]);

    useEffect(() => {
        mountedRef.current = true;

        const renderSingleThumbnail = async (pageNum: number): Promise<boolean> => {
            if (!mountedRef.current || thumbnails.has(pageNum)) return false;
            try {
                const canvas = document.createElement('canvas');
                const renderTask = await pdfService.renderPage(pageNum, canvas, CONFIG.THUMBNAIL_SCALE, 0);
                thumbnailTasksRef.current.set(pageNum, renderTask);
                await renderTask.promise;
                thumbnailTasksRef.current.delete(pageNum);
                if (mountedRef.current) {
                    setThumbnails(prev => new Map(prev).set(pageNum, canvas.toDataURL()));
                }
                return true;
            } catch {
                return false;
            }
        };

        const renderThumbnails = async () => {
            // Render visible thumbnails first (around current page) - high priority
            const start = Math.max(1, currentPage - CONFIG.THUMBNAIL_PREFETCH_BEFORE);
            const end = Math.min(totalPages, currentPage + CONFIG.THUMBNAIL_PREFETCH_AFTER);

            // First batch: pages near current view (immediate)
            for (let i = start; i <= end && mountedRef.current; i++) {
                await renderSingleThumbnail(i);
            }

            // Second batch: remaining pages using requestIdleCallback for non-blocking rendering
            const remainingPages: number[] = [];
            for (let i = 1; i <= Math.min(totalPages, CONFIG.THUMBNAIL_MAX_PAGES); i++) {
                if (!thumbnails.has(i) && (i < start || i > end)) {
                    remainingPages.push(i);
                }
            }

            // Process remaining thumbnails during browser idle time
            const processNextThumbnail = (index: number) => {
                if (!mountedRef.current || index >= remainingPages.length) return;

                idleCallbackRef.current = scheduleIdleWork(() => {
                    renderSingleThumbnail(remainingPages[index]).then(() => {
                        if (mountedRef.current) {
                            processNextThumbnail(index + 1);
                        }
                        return;
                    }).catch(() => {
                        // Continue with next thumbnail on error
                        if (mountedRef.current) {
                            processNextThumbnail(index + 1);
                        }
                    });
                });
            };

            // Start idle processing after a delay
            await new Promise(resolve => setTimeout(resolve, CONFIG.THUMBNAIL_DEFERRED_DELAY_MS));
            if (mountedRef.current && remainingPages.length > 0) {
                processNextThumbnail(0);
            }
        };

        if (totalPages > 0) {
            // Defer thumbnail rendering to not block initial PDF render
            const timer = setTimeout(renderThumbnails, CONFIG.THUMBNAIL_INITIAL_DELAY_MS);
            return () => {
                clearTimeout(timer);
                mountedRef.current = false;
                if (idleCallbackRef.current !== null) {
                    cancelIdleWork(idleCallbackRef.current);
                }
                thumbnailTasksRef.current.forEach((task) => task.cancel());
                thumbnailTasksRef.current.clear();
            };
        }

        return () => {
            mountedRef.current = false;
            if (idleCallbackRef.current !== null) {
                cancelIdleWork(idleCallbackRef.current);
            }
            thumbnailTasksRef.current.forEach((task) => task.cancel());
            thumbnailTasksRef.current.clear();
        };
    }, [pdfService, totalPages, currentPage, scheduleIdleWork, cancelIdleWork]);

    // For large documents (30+ pages), use virtualization
    const useVirtualization = totalPages > 30;
    const itemHeight = 220; // Approximate height per thumbnail

    return (
        <div className="pdf-thumbnail-container" ref={containerRef}>
            {useVirtualization && (
                /* Top spacer for virtualized thumbnails */
                <div style={{ height: (visibleRange.start - 1) * itemHeight }} />
            )}

            {Array.from({ length: totalPages }, (_, i) => i + 1)
                .filter(pageNum => !useVirtualization || (pageNum >= visibleRange.start && pageNum <= visibleRange.end))
                .map((pageNum) => (
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

            {useVirtualization && (
                /* Bottom spacer for virtualized thumbnails */
                <div style={{ height: Math.max(0, totalPages - visibleRange.end) * itemHeight }} />
            )}
        </div>
    );
});

// Outline Panel Component
interface OutlinePanelProps {
    outline: OutlineItem[];
    pdfService: PdfService;
    onNavigate: (page: number) => void;
}

const OutlinePanel: React.FC<OutlinePanelProps> = React.memo(({ outline, pdfService, onNavigate }) => {
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
});

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
    isNavigatingRef: React.MutableRefObject<boolean>;
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
    searchMatches,
    isNavigatingRef
}) => {
    const containerRef = useRef<HTMLDivElement>(null);
    const pageRefs = useRef<Map<number, HTMLDivElement>>(new Map());
    // Track active render tasks for cancellation
    const renderTasksRef = useRef<Map<number, RenderTask>>(new Map());
    // Track render task start times for timeout cleanup
    const renderTaskStartTimesRef = useRef<Map<number, number>>(new Map());
    // Track pages currently being rendered (ref to avoid stale closure issues)
    const renderingPagesRef = useRef<Set<number>>(new Set());
    // IntersectionObserver for efficient visibility detection
    const observerRef = useRef<IntersectionObserver | null>(null);
    // Track visible pages and their intersection ratios for current page detection
    const visiblePagesRef = useRef<Map<number, number>>(new Map());
    // Debounce timer for current page updates
    const pageChangeDebounceRef = useRef<NodeJS.Timeout | null>(null);
    // Track pages that have been successfully rendered at least once (prevents flicker)
    const everRenderedRef = useRef<Set<number>>(new Set());

    // Render a single page - separated to avoid race conditions
    const renderPage = useCallback(async (pageNum: number, canvas: HTMLCanvasElement) => {
        // Check if already rendering this page (use ref, not state)
        if (renderingPagesRef.current.has(pageNum)) {
            return;
        }

        // Mark as rendering
        renderingPagesRef.current.add(pageNum);
        setRenderedPages(prev => new Map(prev).set(pageNum, 'rendering'));

        try {
            // Cancel any existing render task for this page
            const existingTask = renderTasksRef.current.get(pageNum);
            if (existingTask) {
                existingTask.cancel();
                renderTasksRef.current.delete(pageNum);
            }

            // Clear the canvas before rendering to prevent ghosting
            const ctx = canvas.getContext('2d');
            if (ctx) {
                ctx.clearRect(0, 0, canvas.width, canvas.height);
            }

            // Start new render and store the task with start time
            const renderTask = await pdfService.renderPage(pageNum, canvas, scale, rotation);
            renderTasksRef.current.set(pageNum, renderTask);
            renderTaskStartTimesRef.current.set(pageNum, Date.now());

            // Wait for render to complete
            await renderTask.promise;

            // Clean up completed task and start time
            renderTasksRef.current.delete(pageNum);
            renderTaskStartTimesRef.current.delete(pageNum);

            // Get text content
            const text = await pdfService.getTextContent(pageNum);
            setTextContent(prev => new Map(prev).set(pageNum, text));

            // Mark as done and track that this page has been rendered at least once
            setRenderedPages(prev => new Map(prev).set(pageNum, 'done'));
            everRenderedRef.current.add(pageNum);
        } catch (e) {
            // Handle RenderingCancelledException - expected when cancelled
            if (e instanceof Error && e.message.includes('Rendering cancelled')) {
                // Remove from rendering set so it can be retried
                renderingPagesRef.current.delete(pageNum);
                return;
            }
            // Page render failed - remove from rendering map to allow retry
            setRenderedPages(prev => {
                const next = new Map(prev);
                next.delete(pageNum);
                return next;
            });
        } finally {
            // Always remove from rendering set and start time when done
            renderingPagesRef.current.delete(pageNum);
            renderTaskStartTimesRef.current.delete(pageNum);
        }
    }, [pdfService, scale, rotation, setRenderedPages, setTextContent]);

    // Handle page visibility changes from IntersectionObserver
    const handlePageVisible = useCallback((pageNum: number, isIntersecting: boolean, ratio: number) => {
        if (isIntersecting) {
            visiblePagesRef.current.set(pageNum, ratio);

            // Trigger render if not already rendered or rendering
            if (!renderedPages.has(pageNum) && !renderingPagesRef.current.has(pageNum)) {
                const pageRef = pageRefs.current.get(pageNum);
                const canvas = pageRef?.querySelector('canvas');
                if (canvas) {
                    renderPage(pageNum, canvas as HTMLCanvasElement);
                }
            }
        } else {
            visiblePagesRef.current.delete(pageNum);
        }

        // Debounce current page updates to prevent rapid flickering
        if (pageChangeDebounceRef.current) {
            clearTimeout(pageChangeDebounceRef.current);
        }

        pageChangeDebounceRef.current = setTimeout(() => {
            // Skip page updates during programmatic navigation to prevent race conditions
            if (isNavigatingRef.current) {
                return;
            }

            // Update current page based on most visible page
            let mostVisible = 1;
            let maxRatio = 0;
            visiblePagesRef.current.forEach((pageRatio, num) => {
                if (pageRatio > maxRatio) {
                    maxRatio = pageRatio;
                    mostVisible = num;
                }
            });

            if (visiblePagesRef.current.size > 0 && mostVisible !== currentPage) {
                onPageChange(mostVisible);
            }
        }, 50); // 50ms debounce
    }, [renderedPages, renderPage, currentPage, onPageChange, isNavigatingRef]);

    // Cancel all render tasks when scale/rotation changes or on unmount
    useEffect(() => {
        return () => {
            // Cancel all active render tasks on cleanup
            renderTasksRef.current.forEach((task) => {
                task.cancel();
            });
            renderTasksRef.current.clear();
            renderTaskStartTimesRef.current.clear();
            renderingPagesRef.current.clear();
            visiblePagesRef.current.clear();
            if (pageChangeDebounceRef.current) {
                clearTimeout(pageChangeDebounceRef.current);
            }
        };
    }, [scale, rotation]);

    // Periodic cleanup of stale render tasks (tasks running longer than timeout)
    useEffect(() => {
        const cleanupInterval = setInterval(() => {
            const now = Date.now();
            renderTaskStartTimesRef.current.forEach((startTime, pageNum) => {
                if (now - startTime > CONFIG.RENDER_TASK_TIMEOUT_MS) {
                    const task = renderTasksRef.current.get(pageNum);
                    if (task) {
                        task.cancel();
                        renderTasksRef.current.delete(pageNum);
                    }
                    renderTaskStartTimesRef.current.delete(pageNum);
                    renderingPagesRef.current.delete(pageNum);
                }
            });
        }, CONFIG.RENDER_TASK_TIMEOUT_MS / 2); // Check every half the timeout period

        return () => clearInterval(cleanupInterval);
    }, []);

    // Scroll to top when document loads (pageViewports changes)
    useEffect(() => {
        if (containerRef.current && pageViewports.length > 0) {
            containerRef.current.scrollTo(0, 0);
        }
    }, [pageViewports.length]);

    // Set up IntersectionObserver for efficient visibility detection
    useEffect(() => {
        if (!containerRef.current || pageViewports.length === 0) return;

        // Clean up previous observer
        if (observerRef.current) {
            observerRef.current.disconnect();
        }

        // Create observer with buffer margin for pre-loading
        observerRef.current = new IntersectionObserver(
            (entries) => {
                entries.forEach((entry) => {
                    const pageNum = parseInt(entry.target.getAttribute('data-page') || '0', 10);
                    if (pageNum > 0) {
                        handlePageVisible(pageNum, entry.isIntersecting, entry.intersectionRatio);
                    }
                });
            },
            {
                root: containerRef.current,
                rootMargin: `${CONFIG.RENDER_BUFFER_PX}px 0px ${CONFIG.RENDER_BUFFER_PX}px 0px`,
                threshold: [0, 0.1, 0.25, 0.5, 0.75, 1.0] // Multiple thresholds for accurate current page detection
            }
        );

        // Observe all page elements
        pageRefs.current.forEach((pageRef) => {
            observerRef.current?.observe(pageRef);
        });

        return () => {
            observerRef.current?.disconnect();
            visiblePagesRef.current.clear();
        };
    }, [pageViewports.length, handlePageVisible]);

    // Pre-compute matches grouped by page for O(1) lookup instead of O(n) filtering per render
    const matchesByPage = useMemo(() => {
        const map = new Map<number, typeof searchMatches>();
        searchMatches.forEach(match => {
            const arr = map.get(match.pageNum) || [];
            arr.push(match);
            map.set(match.pageNum, arr);
        });
        return map;
    }, [searchMatches]);

    // Get matches for a specific page - O(1) lookup from pre-computed map
    const getPageMatches = useCallback((pageNum: number) => {
        return matchesByPage.get(pageNum) || [];
    }, [matchesByPage]);

    // Pre-compute current match for O(1) lookup
    const currentMatch = useMemo(() => {
        if (currentMatchIndex < 0 || currentMatchIndex >= searchMatches.length) return null;
        return searchMatches[currentMatchIndex];
    }, [searchMatches, currentMatchIndex]);

    // Check if a match is the current match - O(1) comparison
    const isCurrentMatch = useCallback((pageNum: number, itemIndex: number) => {
        if (!currentMatch) return false;
        return currentMatch.pageNum === pageNum && currentMatch.itemIndex === itemIndex;
    }, [currentMatch]);

    // Virtual scrolling: calculate which pages to render
    // For large documents, only render pages near the current view
    const virtualWindow = useMemo(() => {
        const totalPages = pageViewports.length;
        // For small/medium documents, render all pages (threshold: 30 pages)
        if (totalPages <= CONFIG.VIRTUAL_SCROLL_THRESHOLD) {
            return { startPage: 1, endPage: totalPages, useVirtual: false };
        }
        // For large documents, use virtual scrolling with expanded buffer
        const startPage = Math.max(1, currentPage - CONFIG.VIRTUAL_SCROLL_BUFFER_BEFORE);
        const endPage = Math.min(totalPages, currentPage + CONFIG.VIRTUAL_SCROLL_BUFFER_AFTER);
        return { startPage, endPage, useVirtual: true };
    }, [pageViewports.length, currentPage]);

    // Calculate cumulative heights for spacers
    const pageHeights = useMemo(() => {
        return pageViewports.map((vp, index) => {
            const isRotated = rotation % 180 !== 0;
            return (isRotated ? vp.width : vp.height) * scale + 20; // 20px margin
        });
    }, [pageViewports, scale, rotation]);

    // Calculate spacer heights for virtual scrolling
    const { topSpacerHeight, bottomSpacerHeight } = useMemo(() => {
        if (!virtualWindow.useVirtual) {
            return { topSpacerHeight: 0, bottomSpacerHeight: 0 };
        }
        let top = 0;
        let bottom = 0;
        for (let i = 0; i < pageHeights.length; i++) {
            const pageNum = i + 1;
            if (pageNum < virtualWindow.startPage) {
                top += pageHeights[i];
            } else if (pageNum > virtualWindow.endPage) {
                bottom += pageHeights[i];
            }
        }
        return { topSpacerHeight: top, bottomSpacerHeight: bottom };
    }, [virtualWindow, pageHeights]);

    return (
        <div className="pdf-page-wrapper-container" ref={containerRef}>
            {/* Top spacer for virtual scrolling */}
            {virtualWindow.useVirtual && topSpacerHeight > 0 && (
                <div className="pdf-virtual-spacer" style={{ height: topSpacerHeight }} />
            )}

            {pageViewports.map((viewport, index) => {
                const pageNum = index + 1;

                // Skip pages outside virtual window for large documents
                if (virtualWindow.useVirtual &&
                    (pageNum < virtualWindow.startPage || pageNum > virtualWindow.endPage)) {
                    return null;
                }

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
                        {/* Keep canvas visible if it was ever rendered to prevent flicker during re-rendering */}
                        <canvas
                            className="pdf-page-canvas"
                            style={{
                                display: (renderedPages.get(pageNum) === 'done' || everRenderedRef.current.has(pageNum)) ? 'block' : 'none'
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

                        {/* Only show placeholder if page was never rendered */}
                        {renderedPages.get(pageNum) !== 'done' && !everRenderedRef.current.has(pageNum) && (
                            <div className="pdf-page-placeholder" style={{ width, height }}>
                                {renderedPages.get(pageNum) === 'rendering' && (
                                    <div className="pdf-page-loading">Loading...</div>
                                )}
                            </div>
                        )}
                    </div>
                );
            })}

            {/* Bottom spacer for virtual scrolling */}
            {virtualWindow.useVirtual && bottomSpacerHeight > 0 && (
                <div className="pdf-virtual-spacer" style={{ height: bottomSpacerHeight }} />
            )}
        </div>
    );
});

export default PdfViewer;
