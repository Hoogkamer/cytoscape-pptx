import cytoscape from "cytoscape";
import pptxgen from "pptxgenjs";
import { pptxAddSlide, pptxGetLayouts } from "../dist/pptx.esm.js";
import { existsSync } from "fs";

console.log("🧪 Testing cytoscape-pptx package...\n");

// Create a test Cytoscape graph
const cy = cytoscape({
  elements: [
    // Nodes
    { data: { id: "a" } },
    { data: { id: "b" } },
    { data: { id: "c" } },
    { data: { id: "d" } },
    { data: { id: "e" } },

    // Edges with different arrow configurations
    { data: { id: "ab", source: "a", target: "b" } },
    { data: { id: "bc", source: "b", target: "c" } },
    { data: { id: "cd", source: "c", target: "d" } },
    { data: { id: "de", source: "d", target: "e" } },
    { data: { id: "ea", source: "e", target: "a" } },
  ],
  style: [
    {
      selector: "node",
      style: {
        "background-color": "#0074D9",
        "label": "data(id)",
        "color": "#fff",
        "text-valign": "center",
        "text-halign": "center",
        "width": 60,
        "height": 60,
        "font-size": 14,
      },
    },
    {
      selector: "edge",
      style: {
        "width": 3,
        "line-color": "#FF4136",
        "target-arrow-color": "#FF4136",
        "target-arrow-shape": "triangle",
        "curve-style": "bezier",
      },
    },
  ],
  layout: {
    name: "circle",
  },
});

console.log("✓ Created Cytoscape graph with 5 nodes and 5 edges");

// Test getting layouts
const layouts = pptxGetLayouts();
console.log(`✓ Retrieved ${layouts.length} built-in layouts`);
console.log(`  Available layouts: ${layouts.map(l => l.name).join(", ")}\n`);

// Create PowerPoint presentation
const pres = new pptxgen();

// Test with default options
console.log("📄 Creating PowerPoint with default options...");
pptxAddSlide(pres, cy);

// Test with a specific layout
console.log("📄 Creating PowerPoint with 16x9 layout...");
const layout16x9 = layouts.find(l => l.name === "16x9");
pptxAddSlide(pres, cy, { options: layout16x9 });

// Test with custom options
console.log("📄 Creating PowerPoint with custom options...");
pptxAddSlide(pres, cy, {
  options: {
    width: 10,
    height: 7.5,
    marginTop: 0.5,
    marginLeft: 0.5,
    segmentedEdges: true,
  },
});

// Save the presentation
const outputFile = "test/test-output.pptx";
console.log(`\n💾 Writing PowerPoint to ${outputFile}...`);

pres.writeFile({ fileName: outputFile }).then(() => {
  // Verify the file was created
  if (existsSync(outputFile)) {
    console.log("✅ SUCCESS: PowerPoint file created successfully!");
    console.log(`   File: ${outputFile}`);
    console.log(`   Open this file in PowerPoint to inspect the output.\n`);

    console.log("✨ All tests passed!");
    console.log(`\n📎 Output file saved at: ${outputFile}`);
    console.log("   (Delete manually when done inspecting)\n");
    process.exit(0);
  } else {
    console.error("❌ FAILED: PowerPoint file was not created");
    process.exit(1);
  }
}).catch((error) => {
  console.error("❌ FAILED: Error writing PowerPoint file");
  console.error(error);
  process.exit(1);
});
