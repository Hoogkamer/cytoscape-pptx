export default function pptxAddSlide(pres, cy) {
  const slide = pres.addSlide();

  let bbx = cy.elements().boundingBox();
  let size = {
    width: bbx.w / 100,
    height: bbx.h / 100,
  };
  pres.defineLayout({ name: "A3", width: size.width, height: size.height });
  //let scale = calcScale(bbx, size);

  let scale = 0.01;
  pres.layout = "A3";

  //draw groups first so they come under the rest of the nodes
  let groups = cy.nodes(":parent");
  let ultimoParents = groups.nodes(":orphan");
  let rest = groups.difference(ultimoParents);
  drawNodes(slide, ultimoParents, scale, bbx);
  drawNodes(slide, rest, scale, bbx);

  let terms = cy.nodes(":childless");
  let relations = cy.edges();
  drawNodes(slide, terms, scale, bbx);
  drawEdges(slide, relations, scale, bbx);
}
function drawEdges(slide, edges, scale, bbx) {
  edges.forEach((e, i) => {
    let bbx1 = {
      x1: e.sourceEndpoint().x,
      y1: e.sourceEndpoint().y,
      x2: e.targetEndpoint().x,
      y2: e.targetEndpoint().y,
      h: e.targetEndpoint().y - e.sourceEndpoint().y,
      w: e.targetEndpoint().x - e.sourceEndpoint().x,
    };
    let edgeStyle = e.style();
    let eprop = getEdgeLocation(bbx, bbx1, scale);
    slide.addShape("line", {
      ...eprop.location,
      flipH: eprop.flipH,
      flipV: eprop.flipV,
      line: {
        color: rgb2Hex(edgeStyle.lineColor),
        width: px2Num(edgeStyle.width),
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
        shape: "rect",
        ...getLabelLocation(bbx, e.midpoint(), scale),
        fill: { color: "#FFFFFF" },
        align: "center",
        margin: 0,
        fontSize: px2Num(edgeStyle.fontSize) - 5,
      });
    }
  });
}

function drawNodes(slide, nodes, scale, bbx) {
  nodes.forEach((n, i) => {
    let bbx1 = n.boundingBox();
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
      fontSize: px2Num(nodeStyle.fontSize) - 5,
      margin: 0,
      rectRadius: scale * 10,
    };
    slide.addText(nodeStyle.label, shapeparams);
  });
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
  return shapesMapping[nodeStyle.shape];
}
function calcScale(bbx, size) {
  let scaleH = (size.height - 0.3) / bbx.h;
  let scaleW = (size.width - 0.3) / bbx.w;
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
