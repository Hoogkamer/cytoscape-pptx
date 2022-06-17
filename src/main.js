export default function pptxAddSlide(pres, cy) {
  const slide = pres.addSlide();

  let terms = cy.nodes(":childless");

  let relations = window.CYTOSCAPE.edges();

  let bbx = window.CYTOSCAPE.elements().boundingBox();
  let size = {
    width: bbx.w / 100,
    height: bbx.h / 100,
  };
  pres.defineLayout({ name: "A3", width: size.width, height: size.height });
  let scale = calcScale(bbx, size);
  scale = 0.01;
  pres.layout = "A3";

  console.log("sizescale", size, scale);

  //draw groups first so they come under the rest of the nodes
  let groups = window.CYTOSCAPE.nodes(":parent");
  let ultimoParents = groups.nodes(":orphan");
  let rest = groups.difference(ultimoParents);
  drawNodes(pres, slide, ultimoParents, scale, bbx);
  drawNodes(pres, slide, rest, scale, bbx);

  drawNodes(pres, slide, terms, scale, bbx);
  drawEdges(pres, slide, relations, scale, bbx);
}
function drawEdges(pres, slide, edges, scale, bbx) {
  // for testing connection lines
  edges.forEach((e, i) => {
    let bbx1 = {
      x1: e.sourceEndpoint().x,
      y1: e.sourceEndpoint().y,
      x2: e.targetEndpoint().x,
      y2: e.targetEndpoint().y,
      h: e.targetEndpoint().y - e.sourceEndpoint().y,
      w: e.targetEndpoint().x - e.sourceEndpoint().x,
    };
    console.log(">>>", e.data("name"), getEdgeLocation(bbx, bbx1, scale));
    let edgeStyle = e.style();
    console.log(e.data("name"), e.style());
    let eprop = getEdgeLocation(bbx, bbx1, scale);
    slide.addShape(pres.shapes.LINE, {
      ...eprop.location,
      flipH: eprop.flipH,
      flipV: eprop.flipV,
      line: {
        color: rgb2Hex(edgeStyle.lineColor),
        width: px2Num(edgeStyle.width),
        ...arrowType(),
        dashType:
          edgeStyle.lineStyle === "solid"
            ? "solid"
            : edgeStyle.lineStyle === "dashed"
            ? "lgDash"
            : "lgDashDotDot",
      },
    });
    console.log(e.data("name"), e.midpoint());
    if (edgeStyle.label) {
      slide.addText(edgeStyle.label, {
        shape: pres.shapes.RECTANGLE,
        ...getLabelLocation(bbx, e.midpoint(), scale),
        fill: { color: "#FFFFFF" },
        align: "center",
        fontSize: px2Num(edgeStyle.fontSize) - 4,
      });
    }
  });
}

function drawNodes(pres, slide, nodes, scale, bbx) {
  nodes.forEach((n, i) => {
    let bbx1 = n.boundingBox();
    console.log(n.data("name"), n.style());
    let nodeStyle = n.style();

    let shapeparams = {
      shape: getShape(nodeStyle),
      ...getNodeLocation(bbx, bbx1, scale),
      color: rgb2Hex(nodeStyle.color),
      fill: {
        color: rgb2Hex(nodeStyle.backgroundColor),
        transparency: 100 - nodeStyle.backgroundOpacity * 100,
      },
      line: {
        color: rgb2Hex(nodeStyle.borderColor),
        width: 100 * scale * px2Num(nodeStyle.borderWidth),
      },
      align: "center",
      valign: nodeStyle.textValign,
      fontSize: px2Num(nodeStyle.fontSize) - 4,
      rectRadius: scale * 10,
    };
    console.log(shapeparams);
    slide.addText(nodeStyle.label, shapeparams);
  });
}

// reverse relation if it points to the left, otherwise the text is upside down
function arrowType() {
  return { endArrowType: "triangle" };
}
function getLabelLocation(bbx, midpoint, scale) {
  let x1 = midpoint.x;
  let y1 = midpoint.y;

  return {
    x: calcX(bbx, { x1 }, scale) - 0.5,
    y: calcY(bbx, { y1 }, scale),
    w: 1,
    h: 0.1,
  };
}
function getEdgeLocation(bbx, bbx1, scale) {
  let x = calcX(bbx, bbx1, scale);
  let y = calcY(bbx, bbx1, scale);
  let w = calcW(bbx, bbx1, scale);
  let h = calcH(bbx, bbx1, scale);
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
function getNodeLocation(bbx, bbx1, scale) {
  let x = calcX(bbx, bbx1, scale);
  let y = calcY(bbx, bbx1, scale);
  let w = calcW(bbx, bbx1, scale);
  let h = calcH(bbx, bbx1, scale);

  return { x: x, y: y, w: w, h: h };
}

function getShape(nodeStyle) {
  let shapesMapping = {
    ellipse: "ellipse",
    "round-rectangle": "roundRect",
    rectangle: "rect",
  };

  let shape = shapesMapping[nodeStyle.shape];
  console.log(shape);
  return shape;
}
function calcScale(bbx, size) {
  let scaleH = (size.height - 0.3) / bbx.h;
  let scaleW = (size.width - 0.3) / bbx.w;
  console.log(bbx, size, scaleH, scaleW);
  return Math.min(scaleH, scaleW, 1);
}

function calcX(bbx, bbx1, scale) {
  let res = (bbx1.x1 - bbx.x1) * scale;

  return res;
}
function calcY(bbx, bbx1, scale) {
  let res = (bbx1.y1 - bbx.y1) * scale;

  return res;
}
function calcW(bbx, bbx1, scale) {
  let res = bbx1.w * scale;

  return res;
}
function calcH(bbx, bbx1, scale) {
  let res = bbx1.h * scale;

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
