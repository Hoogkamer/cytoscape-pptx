# Cytoscape PPTX Export Demo

This demo page shows how to use the cytoscape-pptx library in a browser environment.

## Running the Demo

### Option 1: Using Python's built-in server

```bash
# From the project root
python -m http.server 8000

# Then open in browser:
# http://localhost:8000/demo/
```

### Option 2: Using Node.js http-server

```bash
# Install http-server globally if you haven't
npm install -g http-server

# From the project root
http-server -p 8000

# Then open in browser:
# http://localhost:8000/demo/
```

### Option 3: Using VS Code Live Server

1. Install the "Live Server" extension in VS Code
2. Right-click on `demo/index.html`
3. Select "Open with Live Server"

### Option 4: Direct file access (may have limitations)

You can also open the file directly in your browser:
```bash
# From the demo directory
xdg-open index.html  # Linux
open index.html      # macOS
start index.html     # Windows
```

Note: Some browsers may have CORS restrictions when opening files directly.

## Features

- **Interactive Graph Visualization**: See your Cytoscape graph rendered in real-time
- **Multiple Layout Presets**: Choose from 16x9, 16x10, 4x3, Wide, A3, A4, or Auto
- **Custom Sizing**: Set custom width and height for your presentation
- **Segmented Edges Toggle**: Enable or disable segmented edges rendering
- **One-Click Export**: Download your graph as a PowerPoint presentation

## How It Works

1. The page loads Cytoscape.js, PptxGenJS, and cytoscape-pptx from CDN/local builds
2. Creates a sample graph with 6 nodes and 8 edges
3. Renders the graph in the browser
4. When you click "Export to PowerPoint", it:
   - Creates a new PowerPoint presentation
   - Converts the Cytoscape graph to PowerPoint shapes
   - Downloads the .pptx file to your computer

## Customization

You can modify the graph in the `<script>` section of `index.html`:

- Add/remove nodes and edges in the `elements` array
- Change styling in the `style` array
- Modify the layout algorithm (currently using 'circle')
