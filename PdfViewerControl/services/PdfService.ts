/**
 * PdfService - Wrapper for PDF.js library
 * Handles PDF loading, rendering, and text extraction
 */

import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy, PDFPageProxy, RenderTask } from 'pdfjs-dist';

// PDF.js configuration - must match installed pdfjs-dist version
const PDF_WORKER_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs';
const PDF_CMAP_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/cmaps/';

export interface PdfMetadata {
    title: string;
    author: string;
    subject: string;
    creator: string;
    pageCount: number;
}

export interface PageViewport {
    width: number;
    height: number;
    scale: number;
    rotation: number;
}

export interface TextItem {
    str: string;
    transform: number[];
    fontName?: string;
}

export interface TextContent {
    items: TextItem[];
}

export interface OutlineItem {
    title: string;
    dest: unknown;
    items?: OutlineItem[];
}

// Re-export RenderTask type from pdfjs-dist
export type { RenderTask } from 'pdfjs-dist';

export class PdfService {
    private document: PDFDocumentProxy | null = null;
    private isInitialized = false;
    // Page cache for performance - eliminates redundant page fetches
    private pageCache = new Map<number, PDFPageProxy>();

    constructor() {
        this.initializeWorker();
    }

    /**
     * Initialize PDF.js worker
     */
    private initializeWorker(): void {
        if (this.isInitialized) return;

        try {
            // Try to use bundled worker first, fall back to CDN
            pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_URL;
            this.isInitialized = true;
        } catch (e) {
            console.warn('Failed to initialize PDF.js worker:', e);
        }
    }

    /**
     * Load a PDF document from URL
     */
    async loadFromUrl(
        url: string,
        onProgress?: (loaded: number, total: number) => void
    ): Promise<PDFDocumentProxy> {
        this.initializeWorker();

        // Add cache buster to avoid stale content
        const cacheBustUrl = url.includes('?') ? `${url}&_=${Date.now()}` : `${url}?_=${Date.now()}`;

        const loadingTask = pdfjsLib.getDocument({
            url: cacheBustUrl,
            cMapUrl: PDF_CMAP_URL,
            cMapPacked: true,
        });

        if (onProgress) {
            loadingTask.onProgress = (progress: { loaded: number; total: number }) => {
                if (progress.total > 0) {
                    onProgress(progress.loaded, progress.total);
                }
            };
        }

        this.document = await loadingTask.promise;
        return this.document;
    }

    /**
     * Load a PDF document from Blob
     */
    async loadFromBlob(
        blob: Blob,
        onProgress?: (loaded: number, total: number) => void
    ): Promise<PDFDocumentProxy> {
        const arrayBuffer = await blob.arrayBuffer();
        return this.loadFromArrayBuffer(arrayBuffer, onProgress);
    }

    /**
     * Load a PDF document from ArrayBuffer
     */
    async loadFromArrayBuffer(
        arrayBuffer: ArrayBuffer,
        onProgress?: (loaded: number, total: number) => void
    ): Promise<PDFDocumentProxy> {
        this.initializeWorker();

        try {
            // Validate that we have data
            if (!arrayBuffer || arrayBuffer.byteLength === 0) {
                throw new Error('File is empty');
            }

            // Check PDF magic bytes (should start with %PDF)
            const headerView = new Uint8Array(arrayBuffer.slice(0, 5));
            const header = String.fromCharCode(...headerView);
            if (!header.startsWith('%PDF')) {
                throw new Error('Invalid PDF file format');
            }

            const loadingTask = pdfjsLib.getDocument({
                data: arrayBuffer,
                cMapUrl: PDF_CMAP_URL,
                cMapPacked: true,
            });

            if (onProgress) {
                loadingTask.onProgress = (progress: { loaded: number; total: number }) => {
                    if (progress.total > 0) {
                        onProgress(progress.loaded, progress.total);
                    }
                };
            }

            this.document = await loadingTask.promise;
            return this.document;
        } catch (error) {
            // Re-throw with more context
            if (error instanceof Error) {
                throw new Error(`PDF loading failed: ${error.message}`);
            }
            throw new Error('PDF loading failed: Unknown error');
        }
    }

    /**
     * Get the current document
     */
    getDocument(): PDFDocumentProxy | null {
        return this.document;
    }

    /**
     * Get total number of pages
     */
    getPageCount(): number {
        return this.document?.numPages || 0;
    }

    /**
     * Get a specific page - uses cache to avoid redundant fetches
     */
    async getPage(pageNumber: number): Promise<PDFPageProxy> {
        if (!this.document) {
            throw new Error('No document loaded');
        }
        if (pageNumber < 1 || pageNumber > this.document.numPages) {
            throw new Error(`Invalid page number: ${pageNumber}`);
        }

        // Check cache first
        const cached = this.pageCache.get(pageNumber);
        if (cached) {
            return cached;
        }

        // Fetch and cache the page
        const page = await this.document.getPage(pageNumber);
        this.pageCache.set(pageNumber, page);
        return page;
    }

    /**
     * Get viewport for a page at given scale
     */
    async getPageViewport(pageNumber: number, scale: number, rotation = 0): Promise<PageViewport> {
        const page = await this.getPage(pageNumber);
        const viewport = page.getViewport({ scale, rotation });
        return {
            width: viewport.width,
            height: viewport.height,
            scale,
            rotation
        };
    }

    /**
     * Get all page viewports at scale 1 - parallelized for performance
     */
    async getAllPageViewports(): Promise<PageViewport[]> {
        if (!this.document) return [];

        const pageCount = this.document.numPages;
        const chunkSize = 10; // Process 10 pages in parallel at a time
        const viewports: PageViewport[] = new Array(pageCount);

        for (let i = 0; i < pageCount; i += chunkSize) {
            const chunkEnd = Math.min(i + chunkSize, pageCount);
            const chunkPromises: Promise<PageViewport>[] = [];

            for (let j = i; j < chunkEnd; j++) {
                const pageNum = j + 1;
                chunkPromises.push(this.getPageViewport(pageNum, 1));
            }

            const chunkResults = await Promise.all(chunkPromises);
            chunkResults.forEach((vp, idx) => {
                viewports[i + idx] = vp;
            });
        }

        return viewports;
    }

    /**
     * Render a page to a canvas
     * Returns the RenderTask so it can be cancelled if needed
     */
    async renderPage(
        pageNumber: number,
        canvas: HTMLCanvasElement,
        scale: number,
        rotation = 0
    ): Promise<RenderTask> {
        const page = await this.getPage(pageNumber);
        const dpr = window.devicePixelRatio || 1;
        const viewport = page.getViewport({ scale: scale * dpr, rotation });

        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.style.width = `${viewport.width / dpr}px`;
        canvas.style.height = `${viewport.height / dpr}px`;

        const context = canvas.getContext('2d');
        if (!context) {
            throw new Error('Could not get canvas context');
        }

        const renderTask = page.render({
            canvasContext: context,
            viewport
        });

        return renderTask;
    }

    /**
     * Get text content of a page
     */
    async getTextContent(pageNumber: number): Promise<TextContent> {
        const page = await this.getPage(pageNumber);
        const textContent = await page.getTextContent();
        return {
            items: textContent.items.map((item) => {
                if ('str' in item) {
                    return {
                        str: item.str,
                        transform: item.transform,
                        fontName: item.fontName
                    };
                }
                return { str: '', transform: [], fontName: undefined };
            })
        };
    }

    /**
     * Get document metadata
     */
    async getMetadata(): Promise<PdfMetadata> {
        if (!this.document) {
            return { title: '', author: '', subject: '', creator: '', pageCount: 0 };
        }

        try {
            const metadata = await this.document.getMetadata();
            const info = metadata.info as Record<string, unknown>;
            return {
                title: (info?.Title as string) || '',
                author: (info?.Author as string) || '',
                subject: (info?.Subject as string) || '',
                creator: (info?.Creator as string) || '',
                pageCount: this.document.numPages
            };
        } catch {
            return {
                title: '',
                author: '',
                subject: '',
                creator: '',
                pageCount: this.document.numPages
            };
        }
    }

    /**
     * Get document outline (bookmarks)
     */
    async getOutline(): Promise<OutlineItem[]> {
        if (!this.document) return [];

        try {
            const outline = await this.document.getOutline();
            if (!outline) return [];

            const mapOutline = (items: {title: string; dest: unknown; items?: unknown[]}[]): OutlineItem[] => {
                return items.map(item => ({
                    title: item.title,
                    dest: item.dest,
                    items: item.items ? mapOutline(item.items as {title: string; dest: unknown; items?: unknown[]}[]) : undefined
                }));
            };

            return mapOutline(outline as {title: string; dest: unknown; items?: unknown[]}[]);
        } catch {
            return [];
        }
    }

    /**
     * Get page index from a destination
     */
    async getPageIndexFromDest(dest: unknown): Promise<number> {
        if (!this.document) return 0;

        try {
            const destination = typeof dest === 'string'
                ? await this.document.getDestination(dest)
                : dest as unknown[];

            if (destination && Array.isArray(destination) && destination.length > 0) {
                const ref = destination[0];
                if (ref && typeof ref === 'object' && 'num' in ref) {
                    const pageIndex = await this.document.getPageIndex(ref as { num: number; gen: number });
                    return pageIndex;
                }
            }
        } catch {
            // Error getting page index
        }
        return 0;
    }

    /**
     * Destroy the document and clean up resources
     */
    async destroy(): Promise<void> {
        // Clear page cache
        this.pageCache.clear();

        if (this.document) {
            await this.document.destroy();
            this.document = null;
        }
    }
}

// Export a singleton instance
export const pdfService = new PdfService();
