# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an **npm library package** that other applications install and use to export Cytoscape.js graphs to PowerPoint presentations using PptxGenJS. It converts Cytoscape graph elements (nodes and edges) into PowerPoint shapes with proper styling, positioning, and layout.

Users install this package via `npm install cytoscape-pptx` and import it into their applications. The package provides two main exports:
- `pptxAddSlide(presentation, cy, options)`: Adds a Cytoscape graph to a PowerPoint presentation
- `pptxGetLayouts()`: Returns available PowerPoint layout presets

## Build System

The project is written in **TypeScript** and uses Rollup with `@rollup/plugin-typescript` to build three distribution formats:

- **UMD** (browser): `dist/pptx.umd.js`
- **CommonJS** (Node): `dist/pptx.cjs.js`
- **ES Module** (bundlers): `dist/pptx.esm.js`
- **Type Definitions**: `dist/main.d.ts`

Build commands:
```bash
npm run build      # Build all formats and generate type definitions
npm run dev        # Build in watch mode
npm test           # Run tests (runs pretest build automatically)
```

The TypeScript configuration (`tsconfig.json`) is set to strict mode for maximum type safety.

## Testing

The `test/test.js` script verifies the package works correctly by:
1. Creating a Cytoscape graph with nodes and edges
2. Exporting it to PowerPoint using various options (default, preset layouts, custom settings)
3. Verifying the output file is created successfully

The test uses the built ES module (`dist/pptx.esm.js`) to simulate how consumers would use the package. It requires `cytoscape` and `pptxgenjs` as dev dependencies.

To add new tests:
- The test script runs in Node.js with Cytoscape in headless mode
- Import the library from `../dist/pptx.esm.js` (not from src)
- Test different graph configurations, layouts, and export options
- Verify edge cases like different arrow configurations, segmented edges, and custom shapes

## Architecture

### Core Components

**src/main.ts** - Single TypeScript file containing all functionality with these key sections:

1. **Main API Functions** (lines 1-90):
   - `pptxAddSlide()`: Main entry point that orchestrates the entire export process
   - `pptxGetLayouts()`: Returns standard PowerPoint layout sizes (16x9, 16x10, 4x3, WIDE, A3, A4, AUTO)
   - `defaultOptions()`: Default export options (width, height, margins, segmentedEdges)

2. **Slide Creation & Scaling** (lines 41-107, 450-485):
   - `createSlide()`: Defines custom layout and creates PowerPoint slide
   - `calcScale()`: Calculates scaling factor to fit graph within slide margins
   - Coordinate transformation functions: `calcX()`, `calcY()`, `calcW()`, `calcH()`

3. **Node Rendering** (lines 196-224, 277-326):
   - `drawNodes()`: Renders Cytoscape nodes as PowerPoint shapes
   - `getNodeLocation()`: Transforms node coordinates to slide coordinates
   - `getShape()`: Maps Cytoscape shapes to PowerPoint shapes
   - `getShapePoints()`: Custom shape definitions for unsupported PowerPoint shapes (triangle, pentagon, hexagon, star, etc.)

4. **Edge Rendering** (lines 109-158, 169-194, 237-276):
   - `drawEdges()`: Renders edges as lines or custom segmented shapes
   - `getEdgeSegments()`: Handles edges with bend points (segmented edges)
   - `getEdgeLocation()`: Transforms straight edge coordinates with flip handling
   - Arrow support: target and source arrows controlled by `targetArrowShape` and `sourceArrowShape`

5. **Styling & Utilities** (lines 462-500):
   - `rgb2Hex()`: Converts RGB color strings to hex
   - `px2Num()`: Extracts numeric value from pixel strings
   - `calcFontSize()`: Scales font sizes appropriately

### Rendering Order

The library renders elements in a specific order to ensure proper layering (lines 20-38):

1. Ultimate parent nodes (parent nodes with no parent)
2. Non-ultimate parent nodes
3. Childless nodes
4. All edges

This ensures parent/group nodes appear behind their children.

### Shape Mapping

Cytoscape shapes are mapped to PowerPoint equivalents. When PowerPoint doesn't have a native equivalent, custom shapes are created using point arrays (see `customShapes` object at lines 330-442):

- Native PowerPoint: `ellipse`, `rectangle`, `diamond`
- Custom shapes (prefixed with `_`): `_triangle`, `_pentagon`, `_hexagon`, `_octagon`, `_star`, `_vee`, etc.

### Edge Handling

Edges can be exported in two modes:

- **Segmented** (`segmentedEdges: true`): Preserves bend points as custom shapes with multiple segments
- **Straight** (`segmentedEdges: false`): Draws all edges as straight lines between endpoints

Control points (curved edges) are not supported and will be rendered as straight lines. Edge labels are positioned at the calculated midpoint.

### Arrow Support

Recent changes added bidirectional arrow support (commit 7ef47f2):
- Edges can have arrows on target end, source end, both, or neither
- Controlled via `targetArrowShape` and `sourceArrowShape` style properties
- Set to "none" to disable arrows on either end

## Dependencies

- **Peer dependencies** (not bundled): `cytoscape`, `pptxgenjs` - Applications using this library must install these separately
- **Dev dependencies for building**:
  - `rollup` - Module bundler
  - `@rollup/plugin-typescript` - TypeScript plugin for Rollup
  - `@rollup/plugin-commonjs` - CommonJS module conversion
  - `@rollup/plugin-node-resolve` - Node module resolution
  - `typescript` - TypeScript compiler
  - `tslib` - TypeScript runtime helpers
  - `@types/node` - Node.js type definitions
  - `@babel/eslint-parser` - ESLint parser
- **Dev dependencies for testing**: `cytoscape`, `pptxgenjs` - Installed as dev dependencies to enable local testing

## Package Configuration

The `package.json` defines:
- **Entry points**: Three distribution formats for different consumption methods
  - `main`: CommonJS build (`dist/pptx.cjs.js`) for Node.js
  - `module`: ES module build (`dist/pptx.esm.js`) for bundlers
  - `browser`: UMD build (`dist/pptx.umd.js`) for browsers
  - `types`: TypeScript type definitions (`dist/main.d.ts`) for IDE autocomplete and type checking
- **Published files**: Only the `dist/` directory is published to npm (see `files` field)
- **Type**: Package uses `"type": "module"` for ES module support in development files (rollup.config.js)

The `.cjs.js` extension ensures the CommonJS build works correctly regardless of the package type setting.

## TypeScript Support

The library is fully typed with comprehensive TypeScript definitions exported for consumer applications. The main exported types include:

- `PptxPresentation` - PptxGenJS presentation interface
- `PptxSlide` - PptxGenJS slide interface
- `CytoscapeInstance` - Cytoscape graph instance
- `CytoscapeCollection` - Collection of Cytoscape elements
- `CytoscapeElement` - Individual node or edge
- `ExportOptions` - Configuration options for export
- `LayoutPreset` - Predefined slide layouts
- `BoundingBox` - Coordinate boundaries
- `Point` - X/Y coordinate pair

TypeScript consumers will get full IntelliSense support, type checking, and inline documentation when using this library.
