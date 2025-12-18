# RSM PDF Viewer PCF Control

A feature-rich PDF and image viewer control for Power Apps Model-Driven Apps. View, navigate, search, and interact with PDF documents and images stored in Dataverse file columns.

## Features

- **PDF Rendering** - High-quality PDF display using PDF.js
- **Image Support** - View JPG, PNG, GIF, BMP, WebP, and other image formats
- **Zoom Controls** - Zoom in/out, fit to page, fit to width, custom percentages (25%-500%)
- **Page Navigation** - Previous/next, go to page, scroll-based tracking
- **Search** - Find text in PDFs with highlighted matches (Ctrl+F)
- **Thumbnails Sidebar** - Visual page thumbnails for quick navigation
- **Outline/Bookmarks** - Navigate PDF structure via bookmarks
- **Rotation** - Rotate documents left/right
- **Dark Mode** - Toggle between light and dark themes
- **Fullscreen** - Immersive viewing experience
- **Print & Download** - Native print and file download
- **Keyboard Shortcuts** - Full keyboard navigation support
- **Auto-Discovery** - Automatically finds file columns on the current record
- **Responsive** - Adapts to container size

## Installation

### Import Solution

1. Download the solution file:
   - **Managed**: `Solutions/RSMPdfViewer/bin/Release/RSMPdfViewer.zip`
   - **Unmanaged**: `Solutions/RSMPdfViewer/bin/Debug/RSMPdfViewer.zip`

2. Go to [Power Apps](https://make.powerapps.com)

3. Navigate to **Solutions** > **Import solution**

4. Select the downloaded `.zip` file and follow the import wizard

### Add to Form

1. Open your Model-Driven App in the form designer

2. Add a new section or use an existing one

3. Click **+ Component** > **Get more components**

4. Search for "PDF Viewer" and add it

5. Configure the control properties:
   - **Default File Column**: Select which file column to display by default
   - **Show Toolbar**: Toggle toolbar visibility
   - **Show Sidebar**: Toggle sidebar visibility
   - **Default Zoom**: Set initial zoom (auto, page-fit, page-width, or percentage)
   - **Theme**: light, dark, or auto

## Configuration Properties

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `defaultFileColumn` | Text | - | Logical name of the file column to display by default |
| `showToolbar` | Boolean | true | Show/hide the toolbar |
| `showSidebar` | Boolean | true | Show/hide the sidebar |
| `defaultZoom` | Text | auto | Initial zoom: `auto`, `page-fit`, `page-width`, or percentage |
| `theme` | Text | light | Theme: `light`, `dark`, or `auto` |
| `rows` | Number | 10 | Number of rows to display (controls height) |

## Output Properties

| Property | Type | Description |
|----------|------|-------------|
| `currentPage` | Number | Current page number being viewed |
| `totalPages` | Number | Total number of pages in the document |
| `selectedColumn` | Text | Currently selected file column name |

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+F` | Open find panel |
| `Ctrl+P` | Print document |
| `Ctrl+G` | Next search match |
| `Ctrl+Shift+G` | Previous search match |
| `Enter` | Next search match (in find panel) |
| `Shift+Enter` | Previous search match (in find panel) |
| `Escape` | Close find panel / dialogs |
| `←` / `→` | Previous / Next page |
| `Page Up` / `Page Down` | Previous / Next page |
| `Home` / `End` | First / Last page |

*Note: Keyboard shortcuts only work when the PDF viewer has focus.*

## Supported File Types

### PDF
- Standard PDF documents
- Text-based PDFs (searchable)
- Note: Scanned/image PDFs display but are not searchable

### Images
- JPEG (.jpg, .jpeg)
- PNG (.png)
- GIF (.gif)
- BMP (.bmp)
- WebP (.webp)
- SVG (.svg)
- TIFF (.tif, .tiff)

## Development

### Prerequisites

- Node.js 18+
- Power Platform CLI (`pac`)
- .NET 6.0 SDK

### Setup

```bash
# Install dependencies
npm install

# Start development server
npm start

# Build control
npm run build

# Build solution (managed)
cd Solutions/RSMPdfViewer
dotnet build --configuration Release

# Build solution (unmanaged)
dotnet build --configuration Debug
```

### Project Structure

```
PDF Viewer/
├── PdfViewerControl/
│   ├── components/
│   │   └── PdfViewer.tsx       # Main React component
│   ├── services/
│   │   ├── PdfService.ts       # PDF.js wrapper
│   │   └── DataverseService.ts # Dataverse Web API
│   ├── css/
│   │   └── PdfViewer.css       # Styles with theme support
│   ├── index.ts                # PCF lifecycle
│   └── ControlManifest.Input.xml
├── Solutions/
│   └── RSMPdfViewer/           # Solution project
├── package.json
└── README.md
```

## Dependencies

- [PDF.js](https://mozilla.github.io/pdf.js/) v4.10.38 - PDF rendering
- React 16+ (provided by PCF framework)

## Browser Support

- Microsoft Edge (Chromium)
- Google Chrome
- Firefox
- Safari

## Known Limitations

- Search only works on text-based PDFs (not scanned/image PDFs)
- Very large PDFs (500+ pages) may have slower initial load
- Annotations and form fields are read-only

## License

MIT License - See [LICENSE](LICENSE) file

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## Support

For issues and feature requests, please use the GitHub issue tracker.
