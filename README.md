# cytoscape-pptx

pptx export for cytoscape

install cytoscape https://js.cytoscape.org

install pptxgenjs https://github.com/gitbrent/PptxGenJS

install cytoscape-pptx:

```
npm install cytoscape-pptx
```

```javascript
import pptxgen from "pptxgenjs";
import pptxAddSlide from "cytoscape-pptx";
import cytoscape from "cytoscape";

var cy = cytoscape({
  container: document.getElementById("cy"), // container to render in

  elements: [
    // list of graph elements to start with
    {
      // node a
      data: { id: "a" },
    },
    {
      // node b
      data: { id: "b" },
    },
    {
      // edge ab
      data: { id: "ab", source: "a", target: "b" },
    },
  ],

  style: [
    // the stylesheet for the graph
    {
      selector: "node",
      style: {
        "background-color": "#666",
        label: "data(id)",
      },
    },

    {
      selector: "edge",
      style: {
        width: 3,
        "line-color": "#ccc",
        "target-arrow-color": "#ccc",
        "target-arrow-shape": "triangle",
        "curve-style": "bezier",
      },
    },
  ],

  layout: {
    name: "grid",
    rows: 1,
  },
});

const pres = new pptxgen();
pptxAddSlide(pres, cy);
pres.writeFile();
```
