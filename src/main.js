function pptxAddSlide(pres, cy, { options }) {
  let thisOptions = { ...defaultOptions(), ...options };

  // calculate sizes and scale
  let graphSize = cy.elements().boundingBox();
  thisOptions = {
    ...thisOptions,
    ...calculateSlideSize({ options: thisOptions, graphSize }),
  };

  let scale = calcScale(graphSize, thisOptions);
  thisOptions.marginTop = scale.centerY; //center graph
  thisOptions.marginLeft = scale.centerX; //center graph
  let slideSize = { scale: scale.scale, graphSize, layout: thisOptions };

  //define presentation size, and add slide
  pres.defineLayout({
    name: "LAYOUT",
    width: thisOptions.width,
    height: thisOptions.height,
  });
  pres.layout = "LAYOUT";
  const slide = pres.addSlide();

  //draw parents first so they come under the rest of the nodes
  let parents = cy.nodes(":parent");
  let ultimoParents = parents.nodes(":orphan");
  let nonUltimoParents = parents.difference(ultimoParents);
  drawNodes({ slide, nodes: ultimoParents, slideSize });
  drawNodes({ slide, nodes: nonUltimoParents, slideSize });

  //draw non group nodes
  let nonParents = cy.nodes(":childless");
  drawNodes({ slide, nodes: nonParents, slideSize });

  //draw edges
  let edges = cy.edges();
  drawEdges({ slide, edges, slideSize });
}

function pptxGetLayouts() {
  const standardLayouts = [
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

function defaultOptions() {
  return {
    width: pptxGetLayouts()[0].width,
    height: pptxGetLayouts()[0].height,
    marginTop: 1,
    marginLeft: 0.2,
  };
}
function calculateSlideSize({ options, graphSize }) {
  if (options.width && options.height) return {};
  else {
    return {
      width: graphSize.w / 100,
      height: graphSize.h / 100,
    };
  }
}
function drawEdges({ slide, edges, slideSize }) {
  edges.forEach((e, i) => {
    let edgeSize = {
      x1: e.sourceEndpoint().x,
      y1: e.sourceEndpoint().y,
      x2: e.targetEndpoint().x,
      y2: e.targetEndpoint().y,
      h: e.targetEndpoint().y - e.sourceEndpoint().y,
      w: e.targetEndpoint().x - e.sourceEndpoint().x,
    };
    let edgeStyle = e.style();
    let eprop = getEdgeLocation({ edgeSize, slideSize });
    slide.addShape("line", {
      ...eprop.location,
      flipH: eprop.flipH,
      flipV: eprop.flipV,
      line: {
        color: rgb2Hex(edgeStyle.lineColor),
        width: 100 * slideSize.scale * px2Num(edgeStyle.width),
        endArrowType: "triangle",
        dashType:
          edgeStyle.lineStyle === "solid"
            ? "solid"
            : edgeStyle.lineStyle === "dashed"
            ? "lgDash"
            : "lgDashDotDot",
      },
    });
    // if edge contains a name, add a textbox for it
    if (edgeStyle.label) {
      slide.addText(edgeStyle.label, {
        ...getLabelLocation({ slideSize, midpoint: e.midpoint() }),
        //fill: { color: "#FFFFFF" },
        align: "center",
        margin: 0,
        fontSize: calcFontSize(edgeStyle.fontSize, slideSize.scale),
      });
    }
  });
}

function drawNodes({ slide, nodes, slideSize }) {
  nodes.forEach((n, i) => {
    let nodeSize = n.boundingBox();
    let nodeStyle = n.style();

    let shapeparams = {
      shape: getShape(nodeStyle),
      ...getNodeLocation({ nodeSize, slideSize }),
      color: rgb2Hex(nodeStyle.color),
      fill: {
        color: rgb2Hex(nodeStyle.backgroundColor),
        transparency: 100 - nodeStyle.backgroundOpacity * 100,
      },
      line: {
        color: rgb2Hex(nodeStyle.borderColor),
        width: 100 * slideSize.scale * px2Num(nodeStyle.borderWidth),
      },
      align: "center",
      valign: nodeStyle.textValign,
      fontSize: calcFontSize(nodeStyle.fontSize, slideSize.scale),
      margin: 0,
      rectRadius: slideSize.scale * 10,
    };
    slide.addText(nodeStyle.label, shapeparams);
  });
}

function getLabelLocation({ slideSize, midpoint }) {
  let x1 = midpoint.x;
  let y1 = midpoint.y;

  return {
    x: calcX({ slideSize, elementSize: { x1 } }) - 0.5,
    y: calcY({ slideSize, elementSize: { y1 } }),
    w: 1,
    h: 0.1,
  };
}
function getEdgeLocation({ edgeSize, slideSize }) {
  let x = calcX({ elementSize: edgeSize, slideSize });
  let y = calcY({ elementSize: edgeSize, slideSize });
  let w = calcW({ elementSize: edgeSize, slideSize });
  let h = calcH({ elementSize: edgeSize, slideSize });
  let flipV = false;
  let flipH = false;

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
  return { location: { x: x, y: y, w: w, h: h }, flipH, flipV };
}
function getNodeLocation({ nodeSize, slideSize }) {
  let x = calcX({ elementSize: nodeSize, slideSize });
  let y = calcY({ elementSize: nodeSize, slideSize });
  let w = calcW({ elementSize: nodeSize, slideSize });
  let h = calcH({ elementSize: nodeSize, slideSize });

  return { x: x, y: y, w: w, h: h };
}

function getShape(nodeStyle) {
  let shapesMapping = {
    ellipse: "ellipse",
    "round-rectangle": "roundRect",
    rectangle: "rect",
  };
  return shapesMapping[nodeStyle.shape];
}
function calcScale(bbx, layout) {
  let heightInch = layout.height - 2 * layout.marginTop;
  let widthInch = layout.width - 2 * layout.marginLeft;
  let scaleH = heightInch / bbx.h;
  let scaleW = widthInch / bbx.w;
  let scale = Math.min(scaleH, scaleW, 0.01);

  // calculate margin to center graph in slide
  let centerY = (layout.height - scale * bbx.h) / 2;
  let centerX = (layout.width - scale * bbx.w) / 2;
  return { scale, centerY, centerX };
}
function calcFontSize(fontSize, scale) {
  return (px2Num(fontSize) - 5) * scale * 100;
}

function calcX({ elementSize, slideSize }) {
  let res =
    (elementSize.x1 - slideSize.graphSize.x1) * slideSize.scale +
    slideSize.layout.marginLeft;
  return res;
}
function calcY({ elementSize, slideSize }) {
  let res =
    (elementSize.y1 - slideSize.graphSize.y1) * slideSize.scale +
    slideSize.layout.marginTop;
  return res;
}
function calcW({ elementSize, slideSize }) {
  let res = elementSize.w * slideSize.scale;
  return res;
}
function calcH({ elementSize, slideSize }) {
  let res = elementSize.h * slideSize.scale;
  return res;
}

function rgb2Hex(color) {
  var arr = [];
  color.replace(/[\d+\.]+/g, function (v) {
    arr.push(parseFloat(v));
  });
  return "#" + arr.slice(0, 3).map(toHex).join("");
}
function px2Num(px) {
  return parseFloat(px.replace("px", ""));
}
function toHex(int) {
  var hex = int.toString(16);
  return hex.length == 1 ? "0" + hex : hex;
}

export { pptxAddSlide, pptxGetLayouts };
