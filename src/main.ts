// Type definitions
interface PptxPresentation {
  defineLayout(layout: { name: string; width: number; height: number }): void;
  layout: string;
  addSlide(): PptxSlide;
}

interface PptxSlide {
  addShape(
    type: string,
    options: Record<string, any>
  ): void;
  addText(
    text: string,
    options: Record<string, any>
  ): void;
}

interface CytoscapeInstance {
  elements(): CytoscapeCollection;
  nodes(selector?: string): CytoscapeCollection;
  edges(): CytoscapeCollection;
}

interface CytoscapeCollection {
  boundingBox(): BoundingBox;
  nodes(selector?: string): CytoscapeCollection;
  difference(other: CytoscapeCollection): CytoscapeCollection;
  forEach(callback: (element: CytoscapeElement, index: number) => void): void;
  length: number;
}

interface CytoscapeElement {
  id(): string;
  boundingBox(): BoundingBox;
  style(): ElementStyle;
  segmentPoints?(): Point[];
  controlPoints?(): Point[];
  sourceEndpoint?(): Point;
  targetEndpoint?(): Point;
  midpoint?(): Point;
  source(): CytoscapeElement;
  target(): CytoscapeElement;
  position(): Point;
}

interface BoundingBox {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  w: number;
  h: number;
}

interface Point {
  x: number;
  y: number;
}

interface ElementStyle {
  label?: string;
  shape?: string;
  color?: string;
  backgroundColor?: string;
  backgroundOpacity?: number;
  borderColor?: string;
  borderWidth?: string | number;
  fontSize?: string | number;
  textValign?: string;
  lineColor?: string;
  width?: string | number;
  targetArrowShape?: string;
  sourceArrowShape?: string;
  lineStyle?: string;
}

interface ExportOptions {
  width?: number;
  height?: number;
  marginTop?: number;
  marginLeft?: number;
  segmentedEdges?: boolean;
}

interface SlideSize {
  scale: number;
  graphSize: BoundingBox;
  options: Required<ExportOptions>;
}

interface ScaleResult {
  scale: number;
  centerY: number;
  centerX: number;
}

interface LayoutPreset {
  name: string;
  width: number;
  height: number;
}

interface EdgeLocation {
  x: number;
  y: number;
  w: number;
  h: number;
  flipH: boolean;
  flipV: boolean;
}

interface NodeLocation {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface ShapeResult {
  shape: string;
  points?: Point[];
  rectRadius?: number;
}

// Main API function
function pptxAddSlide(
  presentation: PptxPresentation,
  cy: CytoscapeInstance,
  { options = {} }: { options?: ExportOptions } = {}
): PptxSlide {
  // calculate sizes and scale
  const graphSize = cy.elements().boundingBox();
  let thisOptions: Required<ExportOptions> = {
    ...defaultOptions(),
    ...options,
  } as Required<ExportOptions>;

  if (!(thisOptions.width && thisOptions.height)) {
    thisOptions = addSizeToOptions({ thisOptions, graphSize });
  }

  const scale = calcScale(graphSize, thisOptions);
  thisOptions.marginTop = scale.centerY; //center graph
  thisOptions.marginLeft = scale.centerX; //center graph

  const slideSize: SlideSize = { scale: scale.scale, graphSize, options: thisOptions };
  const slide = createSlide({ presentation, options: thisOptions });

  //draw parents first so they come under the rest of the nodes
  const parents = cy.nodes(":parent");
  const ultimoParents = parents.nodes(":orphan");
  const nonUltimoParents = parents.difference(ultimoParents);
  drawNodes({ slide, nodes: ultimoParents, slideSize });
  drawNodes({ slide, nodes: nonUltimoParents, slideSize });

  //draw non group nodes
  drawNodes({ slide, nodes: cy.nodes(":childless"), slideSize });

  //draw edges
  const edges = cy.edges();
  drawEdges({
    slide,
    edges,
    slideSize,
    segmentedEdges: thisOptions.segmentedEdges,
  });

  return slide;
}

function createSlide({
  presentation,
  options,
}: {
  presentation: PptxPresentation;
  options: Required<ExportOptions>;
}): PptxSlide {
  presentation.defineLayout({
    name: "LAYOUT",
    width: options.width,
    height: options.height,
  });
  presentation.layout = "LAYOUT";
  return presentation.addSlide();
}

function pptxGetLayouts(): LayoutPreset[] {
  const standardLayouts: LayoutPreset[] = [
    {
      name: "16x9",
      width: 10,
      height: 5.625,
    },
    {
      name: "16x10",
      width: 10,
      height: 6.25,
    },
    {
      name: "4x3",
      width: 10,
      height: 7.5,
    },
    {
      name: "WIDE",
      width: 13.3,
      height: 7.5,
    },
    {
      name: "A3",
      width: 16.5,
      height: 11.7,
    },
    {
      name: "A4",
      width: 11.7,
      height: 8.3,
    },
    {
      name: "AUTO",
      width: 0,
      height: 0,
    },
  ];
  return standardLayouts;
}

function defaultOptions(): Required<ExportOptions> {
  return {
    width: pptxGetLayouts()[0].width,
    height: pptxGetLayouts()[0].height,
    marginTop: 1,
    marginLeft: 0.2,
    segmentedEdges: true,
  };
}

function addSizeToOptions({
  thisOptions,
  graphSize,
}: {
  thisOptions: Required<ExportOptions>;
  graphSize: BoundingBox;
}): Required<ExportOptions> {
  return {
    ...thisOptions,
    width: graphSize.w / 100,
    height: graphSize.h / 100,
  };
}

function drawEdges({
  slide,
  edges,
  slideSize,
  segmentedEdges,
}: {
  slide: PptxSlide;
  edges: CytoscapeCollection;
  slideSize: SlideSize;
  segmentedEdges: boolean;
}): void {
  edges.forEach((e: CytoscapeElement) => {
    const edgeStyle = e.style();
    const lineprop = {
      color: rgb2Hex(edgeStyle.lineColor),
      width: 100 * slideSize.scale * px2Num(edgeStyle.width),
      endArrowType: edgeStyle.targetArrowShape === "none" ? "none" : "triangle",
      beginArrowType:
        edgeStyle.sourceArrowShape === "none" ? "none" : "triangle",
      dashType:
        edgeStyle.lineStyle === "solid"
          ? "solid"
          : edgeStyle.lineStyle === "dashed"
          ? "lgDash"
          : "lgDashDotDot",
    };
    // if it is a segmented edge, then draw a custom shape, otherwise a normal line

    // Safely get segment points (may not work in headless mode)
    let segmentPoints: Point[] | null = null;
    try {
      segmentPoints = e.segmentPoints?.() || null;
    } catch (err) {
      // Ignore error in headless mode
    }

    if (segmentedEdges && segmentPoints) {
      slide.addShape("custGeom", {
        ...getEdgeSegments({ e, slideSize }),
        line: lineprop,
      });
    } else {
      slide.addShape("line", {
        ...getEdgeLocation({ e, slideSize }),
        line: lineprop,
      });
    }

    // if edge contains a name, add a textbox for it
    if (edgeStyle.label) {
      let midpoint = getMidpoint(e);
      // if it is a segmented edge, but we draw it as a straight line, recalculate the midpoint for the label
      // control points (curved edges) are not supported yet, and drawn as straight lines

      // Safely get control points (may not work in headless mode)
      let controlPoints: Point[] | null = null;
      try {
        controlPoints = e.controlPoints?.() || null;
      } catch (err) {
        // Ignore error in headless mode
      }

      if ((!segmentedEdges && segmentPoints) || controlPoints) {
        const src = getSourceEndpoint(e);
        const tgt = getTargetEndpoint(e);
        midpoint = {
          x: (src.x + tgt.x) / 2,
          y: (src.y + tgt.y) / 2,
        };
      }
      slide.addText(edgeStyle.label, {
        ...getLabelLocation({ slideSize, midpoint }),
        align: "center",
        margin: 0,
        fontSize: calcFontSize(edgeStyle.fontSize, slideSize.scale),
      });
    }
  });
}

function updateBbx({
  bbx,
  x,
  y,
}: {
  bbx: Partial<BoundingBox>;
  x: number;
  y: number;
}): Partial<BoundingBox> {
  if (Object.keys(bbx).length === 0) return { x1: x, x2: x, y1: y, y2: y };
  else
    return {
      x1: (bbx.x1 ?? x) < x ? bbx.x1 : x,
      x2: (bbx.x2 ?? x) > x ? bbx.x2 : x,
      y1: (bbx.y1 ?? y) < y ? bbx.y1 : y,
      y2: (bbx.y2 ?? y) > y ? bbx.y2 : y,
    };
}

function getEdgeSegments({
  e,
  slideSize,
}: {
  e: CytoscapeElement;
  slideSize: SlideSize;
}): {
  points: Point[];
  x: number;
  y: number;
  w: number;
  h: number;
} {
  const edgeSegments: Point[] = [];
  edgeSegments.push({ ...getSourceEndpoint(e) });
  // segmentPoints will exist if we're calling this function
  try {
    const segments = e.segmentPoints?.();
    if (segments) {
      segments.forEach((sp: Point) => edgeSegments.push({ ...sp }));
    }
  } catch (err) {
    // If segmentPoints fails, segments array will only have source and target
  }
  edgeSegments.push({ ...getTargetEndpoint(e) });

  //calculate the bounding box
  let bbx: Partial<BoundingBox> = {};
  edgeSegments.forEach((pp: Point) => {
    bbx = updateBbx({ bbx, ...pp });
  });

  // calculate the relative segment positions, relative to the start of the bounding box
  edgeSegments.forEach((pp: Point) => {
    pp.x = (pp.x - (bbx.x1 ?? 0)) * slideSize.scale;
    pp.y = (pp.y - (bbx.y1 ?? 0)) * slideSize.scale;
  });

  return {
    points: edgeSegments,
    x: calcX({ slideSize, elementSize: { x1: bbx.x1 ?? 0 } }),
    y: calcY({ slideSize, elementSize: { y1: bbx.y1 ?? 0 } }),
    w: calcW({ elementSize: { w: (bbx.x2 ?? 0) - (bbx.x1 ?? 0) }, slideSize }),
    h: calcH({ elementSize: { h: (bbx.y2 ?? 0) - (bbx.y1 ?? 0) }, slideSize }),
  };
}

function drawNodes({
  slide,
  nodes,
  slideSize,
}: {
  slide: PptxSlide;
  nodes: CytoscapeCollection;
  slideSize: SlideSize;
}): void {
  nodes.forEach((n: CytoscapeElement) => {
    const nodeSize = n.boundingBox();
    const nodeStyle = n.style();
    const nodeLocation = getNodeLocation({ nodeSize, slideSize });

    const shapeparams: Record<string, any> = {
      ...getShape(nodeStyle, nodeLocation),
      ...nodeLocation,
      color: rgb2Hex(nodeStyle.color),
      fill: {
        color: rgb2Hex(nodeStyle.backgroundColor),
        transparency: 100 - (nodeStyle.backgroundOpacity ?? 1) * 100,
      },
      line: {
        color: rgb2Hex(nodeStyle.borderColor),
        width: 100 * slideSize.scale * px2Num(nodeStyle.borderWidth),
      },
      align: "center",
      valign: nodeStyle.textValign,
      fontSize: calcFontSize(nodeStyle.fontSize, slideSize.scale),
      margin: 0,
      name: `node-${n.id()}`, // Add unique identifier for better PowerPoint management
    };
    if (nodeStyle.shape === "round-rectangle") {
      shapeparams.rectRadius = slideSize.scale * 10;
    }
    slide.addText(nodeStyle.label ?? "", shapeparams);
  });
}

// Helper functions to safely get endpoints (with fallback for Node.js/headless mode)
function getSourceEndpoint(edge: CytoscapeElement): Point {
  try {
    const endpoint = edge.sourceEndpoint?.();
    if (endpoint) return endpoint;
  } catch (err) {
    // Fall through to fallback
  }
  // Fallback to source node position
  const source = edge.source();
  return source.position();
}

function getTargetEndpoint(edge: CytoscapeElement): Point {
  try {
    const endpoint = edge.targetEndpoint?.();
    if (endpoint) return endpoint;
  } catch (err) {
    // Fall through to fallback
  }
  // Fallback to target node position
  const target = edge.target();
  return target.position();
}

function getMidpoint(edge: CytoscapeElement): Point {
  try {
    const midpoint = edge.midpoint?.();
    if (midpoint) return midpoint;
  } catch (err) {
    // Fall through to fallback
  }
  // Fallback to calculating midpoint from endpoints
  const src = getSourceEndpoint(edge);
  const tgt = getTargetEndpoint(edge);
  return {
    x: (src.x + tgt.x) / 2,
    y: (src.y + tgt.y) / 2,
  };
}

function getLabelLocation({
  slideSize,
  midpoint,
}: {
  slideSize: SlideSize;
  midpoint: Point;
}): {
  x: number;
  y: number;
  w: number;
  h: number;
} {
  const x1 = midpoint.x;
  const y1 = midpoint.y;

  return {
    x: calcX({ slideSize, elementSize: { x1 } }) - 0.5,
    y: calcY({ slideSize, elementSize: { y1 } }),
    w: 1,
    h: 0.1,
  };
}

function getEdgeLocation({
  e,
  slideSize,
}: {
  e: CytoscapeElement;
  slideSize: SlideSize;
}): EdgeLocation {
  const sourceEndpoint = getSourceEndpoint(e);
  const targetEndpoint = getTargetEndpoint(e);
  const edgeSize = {
    x1: sourceEndpoint.x,
    y1: sourceEndpoint.y,
    x2: targetEndpoint.x,
    y2: targetEndpoint.y,
    h: targetEndpoint.y - sourceEndpoint.y,
    w: targetEndpoint.x - sourceEndpoint.x,
  };
  let x = calcX({ elementSize: edgeSize, slideSize });
  let y = calcY({ elementSize: edgeSize, slideSize });
  let w = calcW({ elementSize: edgeSize, slideSize });
  let h = calcH({ elementSize: edgeSize, slideSize });
  let flipV = false;
  let flipH = false;

  // height and width cannot be negative, so correct and rotate to make them positive
  if (w >= 0 && h >= 0) {
    flipV = false;
    flipH = false;
  } else if (w < 0 && h >= 0) {
    flipV = false;
    flipH = true;
    x = x + w;
    w = -w;
  } else if (w < 0 && h < 0) {
    flipV = true;
    flipH = true;
    x = x + w;
    w = -w;
    y = y + h;
    h = -h;
  } else if (w >= 0 && h < 0) {
    flipV = true;
    flipH = false;
    y = y + h;
    h = -h;
  }
  return { x, y, w, h, flipH, flipV };
}

function getNodeLocation({
  nodeSize,
  slideSize,
}: {
  nodeSize: BoundingBox;
  slideSize: SlideSize;
}): NodeLocation {
  const x = calcX({ elementSize: nodeSize, slideSize });
  const y = calcY({ elementSize: nodeSize, slideSize });
  const w = calcW({ elementSize: nodeSize, slideSize });
  const h = calcH({ elementSize: nodeSize, slideSize });

  return { x, y, w, h };
}

function getShape(nodeStyle: ElementStyle, nodeLocation: NodeLocation): ShapeResult {
  // translate cytoscape shapes to powerpoint shapes
  const shapesMapping: Record<string, string> = {
    ellipse: "ellipse",
    triangle: "_triangle",
    "round-triangle": "_triangle",
    rectangle: "rect",
    "round-rectangle": "roundRect",
    "bottom-round-rectangle": "_bottomcutrectange",
    "cut-rectangle": "octagon",
    barrel: "_barrel",
    rhomboid: "_rhomboid",
    diamond: "diamond",
    "round-diamond": "trapezoid",
    pentagon: "_pentagon",
    "round-pentagon": "_pentagon",
    hexagon: "_hexagon",
    "round-hexagon": "_hexagon",
    "concave-hexagon": "_concavehexagon",
    heptagon: "_heptagon",
    "round-heptagon": "_heptagon",
    octagon: "_octagon",
    "round-octagon": "_octagon",
    star: "_star",
    tag: "_tag",
    "round-tag": "rightArrow",
    vee: "_vee",
  };

  const shape = shapesMapping[nodeStyle.shape ?? "ellipse"] || "ellipse";

  if (shape[0] !== "_") {
    // shape is available in powerpoint
    return { shape };
  } else {
    // shape is not available in powerpoint, so create a custom shape
    return {
      shape: "custGeom",
      points: getShapePoints(shape, nodeLocation),
    };
  }
}

function getShapePoints(shape: string, nodeLocation: NodeLocation): Point[] {
  const width = nodeLocation.w;
  const height = nodeLocation.h;
  const customShapes: Record<string, Point[]> = {
    _triangle: [
      { x: 0.0, y: 1.0 },
      { x: 0.5, y: 0.0 },
      { x: 1.0, y: 1.0 },
      { x: 0.0, y: 1.0 },
    ],
    _tag: [
      { x: 0.0, y: 0.0 },
      { x: 0.66, y: 0.0 },
      { x: 1.0, y: 0.5 },
      { x: 0.66, y: 1.0 },
      { x: 0.0, y: 1.0 },
      { x: 0.0, y: 0.0 },
    ],
    _vee: [
      { x: 0.0, y: 0.0 },
      { x: 0.5, y: 0.34 },
      { x: 1.0, y: 0.0 },
      { x: 0.5, y: 1.0 },
      { x: 0.0, y: 0.0 },
    ],
    _rhomboid: [
      { x: 0.0, y: 0.0 },
      { x: 0.66, y: 0.0 },
      { x: 1.0, y: 1.0 },
      { x: 0.33, y: 1.0 },
      { x: 0.0, y: 0.0 },
    ],
    _pentagon: [
      { x: 0.0, y: 0.4 },
      { x: 0.5, y: 0.0 },
      { x: 1.0, y: 0.4 },
      { x: 0.8, y: 1.0 },
      { x: 0.2, y: 1.0 },
      { x: 0.0, y: 0.4 },
    ],
    _hexagon: [
      { x: 0.0, y: 0.5 },
      { x: 0.2, y: 0.0 },
      { x: 0.8, y: 0.0 },
      { x: 1.0, y: 0.5 },
      { x: 0.8, y: 1.0 },
      { x: 0.2, y: 1.0 },
      { x: 0.0, y: 0.5 },
    ],
    _heptagon: [
      { x: 0.0, y: 0.6 },
      { x: 0.15, y: 0.2 },
      { x: 0.5, y: 0.0 },
      { x: 0.85, y: 0.2 },
      { x: 1.0, y: 0.6 },
      { x: 0.7, y: 1.0 },
      { x: 0.3, y: 1.0 },
      { x: 0.0, y: 0.6 },
    ],
    _concavehexagon: [
      { x: 0.0, y: 0.0 },
      { x: 1.0, y: 0.0 },
      { x: 0.85, y: 0.5 },
      { x: 1.0, y: 1.0 },
      { x: 0.0, y: 1.0 },
      { x: 0.15, y: 0.5 },
      { x: 0.0, y: 0.0 },
    ],
    _octagon: [
      { x: 0.0, y: 0.3 },
      { x: 0.3, y: 0.0 },
      { x: 0.7, y: 0.0 },
      { x: 1.0, y: 0.3 },
      { x: 1.0, y: 0.7 },
      { x: 0.7, y: 1.0 },
      { x: 0.3, y: 1.0 },
      { x: 0.0, y: 0.7 },
      { x: 0.0, y: 0.3 },
    ],
    _star: [
      { x: 0.0, y: 0.4 },
      { x: 0.33, y: 0.27 },
      { x: 0.5, y: 0.0 },
      { x: 0.67, y: 0.27 },
      { x: 1.0, y: 0.38 },
      { x: 0.8, y: 0.67 },
      { x: 0.8, y: 1.0 },
      { x: 0.5, y: 0.9 },
      { x: 0.2, y: 1.0 },
      { x: 0.2, y: 0.67 },
      { x: 0.0, y: 0.4 },
    ],
    _barrel: [
      { x: 0.0, y: 0.1 },
      { x: 0.2, y: 0.0 },
      { x: 0.8, y: 0.0 },
      { x: 1.0, y: 0.1 },
      { x: 1.0, y: 0.9 },
      { x: 0.8, y: 1.0 },
      { x: 0.2, y: 1.0 },
      { x: 0.0, y: 0.9 },
      { x: 0.0, y: 0.1 },
    ],

    _bottomcutrectange: [
      { x: 0.0, y: 0.0 },
      { x: 1.0, y: 0.0 },
      { x: 1.0, y: 0.8 },
      { x: 0.95, y: 0.95 },
      { x: 0.8, y: 1.0 },
      { x: 0.2, y: 1.0 },
      { x: 0.05, y: 0.95 },
      { x: 0.0, y: 0.8 },
      { x: 0.0, y: 0.0 },
    ],
  };
  const thisShape = customShapes[shape];
  thisShape.forEach((tp: Point) => {
    tp.x = tp.x * width;
    tp.y = tp.y * height;
  });
  return thisShape;
}

function calcScale(bbx: BoundingBox, options: Required<ExportOptions>): ScaleResult {
  const heightInch = options.height - 2 * options.marginTop;
  const widthInch = options.width - 2 * options.marginLeft;
  const scaleH = heightInch / bbx.h;
  const scaleW = widthInch / bbx.w;
  const scale = Math.min(scaleH, scaleW, 0.01);

  // calculate margin to center graph in slide
  const centerY = (options.height - scale * bbx.h) / 2;
  const centerX = (options.width - scale * bbx.w) / 2;
  return { scale, centerY, centerX };
}

function calcFontSize(fontSize: string | number | undefined, scale: number): number {
  return (px2Num(fontSize) - 5) * scale * 100;
}

function calcX({
  elementSize,
  slideSize,
}: {
  elementSize: { x1: number };
  slideSize: SlideSize;
}): number {
  const res =
    (elementSize.x1 - slideSize.graphSize.x1) * slideSize.scale +
    slideSize.options.marginLeft;
  return res;
}

function calcY({
  elementSize,
  slideSize,
}: {
  elementSize: { y1: number };
  slideSize: SlideSize;
}): number {
  const res =
    (elementSize.y1 - slideSize.graphSize.y1) * slideSize.scale +
    slideSize.options.marginTop;
  return res;
}

function calcW({
  elementSize,
  slideSize,
}: {
  elementSize: { w: number };
  slideSize: SlideSize;
}): number {
  const res = elementSize.w * slideSize.scale;
  return res;
}

function calcH({
  elementSize,
  slideSize,
}: {
  elementSize: { h: number };
  slideSize: SlideSize;
}): number {
  const res = elementSize.h * slideSize.scale;
  return res;
}

function rgb2Hex(color: string | undefined): string {
  if (!color) return "000000"; // Default to black if color is undefined
  const arr: number[] = [];
  color.replace(/[\d+\.]+/g, function (v: string): string {
    arr.push(parseFloat(v));
    return v;
  });
  return arr.slice(0, 3).map(toHex).join("").toUpperCase();
}

function px2Num(px: string | number | undefined): number {
  if (!px && px !== 0) return 0; // Default to 0 if px is undefined
  if (typeof px === "number") return px; // Already a number
  return parseFloat(String(px).replace("px", ""));
}

function toHex(int: number): string {
  const hex = int.toString(16);
  return hex.length === 1 ? "0" + hex : hex;
}

export { pptxAddSlide, pptxGetLayouts };
export type {
  PptxPresentation,
  PptxSlide,
  CytoscapeInstance,
  CytoscapeCollection,
  CytoscapeElement,
  ExportOptions,
  LayoutPreset,
  BoundingBox,
  Point,
};
