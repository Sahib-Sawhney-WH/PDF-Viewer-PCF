/**
 * DataverseService - Handles all Dataverse Web API interactions
 * Includes form context detection, file column discovery, and file fetching
 */

export interface FileColumn {
    logicalName: string;
    displayName: string;
    type: 'File' | 'Image';
    hasFile?: boolean;
}

export interface ViewerConfig {
    tableName: string | null;
    tableSetName: string | null;
    recordId: string | null;
    fileColumns: FileColumn[];
    selectedColumn: string | null;
    defaultColumn: string | null;
    baseUrl: string;
    isDataverseContext: boolean;
}

// Declare global Xrm type for D365
declare global {
    interface Window {
        Xrm?: {
            Page?: {
                data?: {
                    entity?: {
                        getEntityName: () => string;
                        getId: () => string;
                    };
                };
            };
        };
    }
}

export class DataverseService {
    private config: ViewerConfig;
    private context: ComponentFramework.Context<unknown> | null = null;

    constructor() {
        this.config = {
            tableName: null,
            tableSetName: null,
            recordId: null,
            fileColumns: [],
            selectedColumn: null,
            defaultColumn: null,
            baseUrl: window.location.origin,
            isDataverseContext: false
        };
    }

    /**
     * Initialize with PCF context
     */
    setContext(context: ComponentFramework.Context<unknown>): void {
        this.context = context;
    }

    /**
     * Set the default file column from manifest property
     */
    setDefaultColumn(columnName: string | null): void {
        this.config.defaultColumn = columnName;
    }

    /**
     * Get the current configuration
     */
    getConfig(): ViewerConfig {
        return { ...this.config };
    }

    /**
     * Try to get Xrm context from various sources
     */
    private getXrmContext(): typeof window.Xrm | null {
        try {
            if (typeof window.Xrm !== 'undefined') return window.Xrm;
            if (typeof parent !== 'undefined' && (parent as Window & typeof globalThis).Xrm) {
                return (parent as Window & typeof globalThis).Xrm;
            }
            if (typeof window.parent !== 'undefined' && (window.parent as Window & typeof globalThis).Xrm) {
                return (window.parent as Window & typeof globalThis).Xrm;
            }
        } catch {
            // Cross-origin access denied
        }
        return null;
    }

    /**
     * Get form context from Xrm
     */
    private getFormContext(): NonNullable<typeof window.Xrm>['Page'] | null {
        const xrm = this.getXrmContext();
        if (!xrm) return null;

        try {
            if (xrm.Page?.data) return xrm.Page;

            // Try parent frames
            if (typeof parent !== 'undefined') {
                for (let i = 0; i < 5; i++) {
                    try {
                        let p = parent as Window & typeof globalThis;
                        for (let j = 0; j < i; j++) {
                            p = p.parent as Window & typeof globalThis;
                        }
                        if (p.Xrm?.Page?.data) return p.Xrm.Page;
                    } catch {
                        // Cross-origin
                    }
                }
            }
        } catch {
            // Error accessing form context
        }
        return null;
    }

    /**
     * Detect form context - table name and record ID
     */
    async detectFormContext(): Promise<boolean> {
        // Try Xrm.Page first
        const formContext = this.getFormContext();
        if (formContext?.data?.entity) {
            try {
                this.config.tableName = formContext.data.entity.getEntityName();
                this.config.recordId = formContext.data.entity.getId().replace(/[{}]/g, '');
                return true;
            } catch {
                // Fall through to URL params
            }
        }

        // Try URL parameters
        const urlParams = new URLSearchParams(window.location.search);

        this.config.recordId = urlParams.get('id')?.replace(/[{}]/g, '') || null;
        this.config.tableName = urlParams.get('typename') || urlParams.get('etn') || null;

        // Try data parameter (D365 custom parameters)
        const dataParam = urlParams.get('data');
        if (dataParam) {
            try {
                const dataParams = new URLSearchParams(decodeURIComponent(dataParam));
                this.config.recordId = this.config.recordId || dataParams.get('id')?.replace(/[{}]/g, '') || null;
                this.config.tableName = this.config.tableName || dataParams.get('typename') || dataParams.get('table') || null;
            } catch {
                // Invalid data param
            }
        }

        // Try GUID extraction from URL
        if (!this.config.recordId) {
            const guidMatch = /[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}/i.exec(window.location.href);
            if (guidMatch) {
                this.config.recordId = guidMatch[0];
            }
        }

        return !!(this.config.tableName && this.config.recordId);
    }

    /**
     * Get the entity set name (plural) for Web API calls
     */
    async getEntitySetName(logicalName: string): Promise<string> {
        try {
            const response = await fetch(
                `${this.config.baseUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${logicalName}')?$select=EntitySetName`,
                {
                    headers: {
                        'Accept': 'application/json',
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0'
                    }
                }
            );

            if (response.ok) {
                const data = await response.json();
                return data.EntitySetName;
            }
        } catch {
            // Fall back to guessing
        }

        // Fallback: guess the plural form
        if (logicalName.endsWith('y')) {
            return logicalName.slice(0, -1) + 'ies';
        } else if (logicalName.endsWith('s') || logicalName.endsWith('x') ||
                   logicalName.endsWith('ch') || logicalName.endsWith('sh')) {
            return logicalName + 'es';
        }
        return logicalName + 's';
    }

    /**
     * Discover file columns on the entity
     */
    async discoverFileColumns(): Promise<FileColumn[]> {
        if (!this.config.tableName) return [];

        this.config.fileColumns = [];

        try {
            // Get File attributes
            const fileResponse = await fetch(
                `${this.config.baseUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${this.config.tableName}')/Attributes/Microsoft.Dynamics.CRM.FileAttributeMetadata?$select=LogicalName,DisplayName,SchemaName`,
                {
                    headers: {
                        'Accept': 'application/json',
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0'
                    }
                }
            );

            if (fileResponse.ok) {
                const fileData = await fileResponse.json();
                for (const attr of fileData.value) {
                    this.config.fileColumns.push({
                        logicalName: attr.LogicalName,
                        displayName: attr.DisplayName?.UserLocalizedLabel?.Label || attr.SchemaName || attr.LogicalName,
                        type: 'File'
                    });
                }
            }

            // Get Image attributes
            const imageResponse = await fetch(
                `${this.config.baseUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${this.config.tableName}')/Attributes/Microsoft.Dynamics.CRM.ImageAttributeMetadata?$select=LogicalName,DisplayName,SchemaName`,
                {
                    headers: {
                        'Accept': 'application/json',
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0'
                    }
                }
            );

            if (imageResponse.ok) {
                const imageData = await imageResponse.json();
                for (const attr of imageData.value) {
                    // Skip the default entity image
                    if (attr.LogicalName === 'entityimage') continue;

                    this.config.fileColumns.push({
                        logicalName: attr.LogicalName,
                        displayName: attr.DisplayName?.UserLocalizedLabel?.Label || attr.SchemaName || attr.LogicalName,
                        type: 'Image'
                    });
                }
            }
        } catch {
            // Error discovering columns
        }

        return this.config.fileColumns;
    }

    /**
     * Check which columns actually have files attached
     */
    async checkWhichColumnsHaveFiles(): Promise<void> {
        if (!this.config.fileColumns.length || !this.config.recordId || !this.config.tableSetName) return;

        try {
            const columns = this.config.fileColumns.map(c => c.logicalName).join(',');
            const response = await fetch(
                `${this.config.baseUrl}/api/data/v9.2/${this.config.tableSetName}(${this.config.recordId})?$select=${columns}`,
                {
                    headers: {
                        'Accept': 'application/json',
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0'
                    }
                }
            );

            if (response.ok) {
                const record = await response.json();
                for (const col of this.config.fileColumns) {
                    col.hasFile = !!record[col.logicalName];
                }
            }
        } catch {
            // Error checking files
        }
    }

    /**
     * Initialize the service - detect context and discover columns
     */
    async initialize(): Promise<boolean> {
        const hasContext = await this.detectFormContext();
        if (!hasContext) {
            this.config.isDataverseContext = false;
            return false;
        }

        this.config.isDataverseContext = true;
        this.config.tableSetName = await this.getEntitySetName(this.config.tableName!);
        await this.discoverFileColumns();
        await this.checkWhichColumnsHaveFiles();

        return true;
    }

    /**
     * Get the URL for a file column's content
     */
    getFileUrl(columnName: string): string {
        if (!this.config.tableSetName || !this.config.recordId) {
            throw new Error('Dataverse context not initialized');
        }
        return `${this.config.baseUrl}/api/data/v9.2/${this.config.tableSetName}(${this.config.recordId})/${columnName}/$value`;
    }

    /**
     * Fetch file content as a blob
     */
    async fetchFile(columnName: string): Promise<{ blob: Blob; contentType: string }> {
        const url = this.getFileUrl(columnName);

        const response = await fetch(url, {
            headers: {
                'Accept': '*/*'
            }
        });

        if (!response.ok) {
            throw new Error(`Failed to fetch file: ${response.statusText}`);
        }

        const blob = await response.blob();
        const contentType = response.headers.get('Content-Type') || 'application/octet-stream';

        return { blob, contentType };
    }

    /**
     * Get columns that have files attached
     */
    getColumnsWithFiles(): FileColumn[] {
        return this.config.fileColumns.filter(c => c.hasFile);
    }

    /**
     * Select a column to display
     */
    selectColumn(columnName: string): void {
        this.config.selectedColumn = columnName;
    }

    /**
     * Get the best column to auto-select
     * Priority: 1. defaultColumn, 2. only column with file, 3. null (show picker)
     */
    getAutoSelectColumn(): string | null {
        const columnsWithFiles = this.getColumnsWithFiles();

        // Priority 1: Default column from manifest (if it has a file)
        if (this.config.defaultColumn) {
            const defaultCol = columnsWithFiles.find(c => c.logicalName === this.config.defaultColumn);
            if (defaultCol) return this.config.defaultColumn;
        }

        // Priority 2: Auto-select if only one column has a file
        if (columnsWithFiles.length === 1) {
            return columnsWithFiles[0].logicalName;
        }

        // Priority 3: Multiple columns or none - return null to show picker
        return null;
    }
}

export const dataverseService = new DataverseService();
