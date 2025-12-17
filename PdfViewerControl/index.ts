/**
 * PDF Viewer PCF Control
 * Displays PDF and image files from Dataverse file columns
 */

import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { PdfViewer } from "./components/PdfViewer";

export class PdfViewerControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _notifyOutputChanged: () => void;

    // Output values
    private _currentPage = 1;
    private _totalPages = 0;
    private _selectedColumn = '';

    constructor() {
        // Empty
    }

    /**
     * Initializes the control instance.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;

        // Register for container resize events
        context.mode.trackContainerResize(true);

        // Render the React component
        this.renderComponent(context);
    }

    /**
     * Called when any value in the property bag has changed.
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Re-render component with updated props
        this.renderComponent(context);
    }

    /**
     * Renders the React component
     */
    private renderComponent(context: ComponentFramework.Context<IInputs>): void {
        // Get properties from manifest
        const defaultFileColumn = context.parameters.defaultFileColumn?.raw || undefined;
        const showToolbar = context.parameters.showToolbar?.raw !== false;
        const showSidebar = context.parameters.showSidebar?.raw === true; // Default to false for performance
        const defaultZoom = context.parameters.defaultZoom?.raw || 'auto';
        const theme = context.parameters.theme?.raw || 'light';

        // Get rows parameter (like markdown editor)
        const rowsParam = context.parameters.rows?.raw;
        const rows = rowsParam || 10;

        // Get allocated dimensions from context
        const allocatedWidth = context.mode.allocatedWidth;

        // Calculate height from rows setting
        // Each row is approximately 54px (same as markdown editor)
        const height = rows * 54 + 50;
        const width = allocatedWidth > 0 ? allocatedWidth : undefined;

        // Render the component using React 16 API
        ReactDOM.render(
            React.createElement(PdfViewer, {
                defaultFileColumn,
                showToolbar,
                showSidebar,
                defaultZoom,
                theme,
                width,
                height,
                onPageChange: (currentPage: number, totalPages: number) => {
                    if (this._currentPage !== currentPage || this._totalPages !== totalPages) {
                        this._currentPage = currentPage;
                        this._totalPages = totalPages;
                        this._notifyOutputChanged();
                    }
                },
                onColumnChange: (columnName: string) => {
                    if (this._selectedColumn !== columnName) {
                        this._selectedColumn = columnName;
                        this._notifyOutputChanged();
                    }
                }
            }),
            this._container
        );
    }

    /**
     * Returns outputs from the control
     */
    public getOutputs(): IOutputs {
        return {
            currentPage: this._currentPage,
            totalPages: this._totalPages,
            selectedColumn: this._selectedColumn
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree.
     */
    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this._container);
    }
}
