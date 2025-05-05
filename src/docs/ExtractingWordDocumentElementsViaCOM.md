# Extracting Information from Microsoft Word Documents via COM

This document provides a comprehensive guide to extracting information from Microsoft Word documents using COM (Component Object Model) automation, with a focus on shapes, charts, groupings, and canvases. It includes suggestions for mapping these elements to SVG (Scalable Vector Graphics) equivalents, including embedded text. The information is generic and applicable to any COM-compatible programming environment, though the user has indicated they will use the TypeScript `winax` library for COM interaction.

## 1. Overview of COM Automation for Word

COM is a Microsoft technology that allows programmatic control of applications like Microsoft Word. The Word Object Model exposes objects such as `Document`, `Shapes`, `InlineShapes`, `GroupShapes`, and `CanvasShapes`, enabling access to textual and graphical elements in a Word document.

### Key Objects in the Word Object Model
- **Application**: Represents the Word application.
- **Document**: Represents an open Word document.
- **Shapes**: A collection of floating shapes (e.g., rectangles, lines, pictures).
- **InlineShapes**: A collection of shapes embedded in the text flow (e.g., inline images).
- **GroupShapes**: A collection of shapes grouped together.
- **CanvasShapes**: Shapes within a drawing canvas.
- **Chart**: Represents a chart object.
- **TextFrame**: Contains text within shapes or canvases.

### Prerequisites
- A COM-compatible programming environment (e.g., TypeScript with `winax`, Python with `pywin32`, C#, VB.NET).
- Microsoft Word installed on the system.
- Basic understanding of the Word Object Model (refer to [Microsoft's Word VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/word)).

## 2. Extracting Document Elements

Below are the steps and techniques for extracting various elements from a Word document, focusing on shapes, charts, groupings, and canvases.

### 2.1. Opening a Word Document
To extract information, you first need to open the document using the Word Application object.

**Generic COM Steps**:
1. Instantiate the Word Application: Create an instance of `Word.Application`.
2. Open the Document: Use `Documents.Open` to load the `.docx` or `.doc` file.
3. Access Elements: Navigate the object model to extract desired elements.
4. Close and Clean Up: Save or discard changes, close the document, and quit the application.

**Pseudocode**:
```pseudocode
app = CreateObject("Word.Application")
app.Visible = false  // Run in background
doc = app.Documents.Open("path/to/document.docx")
// Extract elements (see below)
doc.Close()
app.Quit()
```

### 2.2. Extracting Shapes
Shapes in Word include geometric shapes, text boxes, images, and other floating objects. They are stored in the `Shapes` collection.

**Steps**:
1. Iterate through `doc.Shapes` to access each shape.
2. Extract properties like `Type`, `Width`, `Height`, `Left`, `Top`, `Fill`, `Line`, and `TextFrame` (for text-containing shapes).
3. Handle specific shape types (e.g., `msoPicture`, `msoTextBox`, `msoAutoShape`).

**Key Properties**:
- `Type`: Indicates the shape type (e.g., `msoShapeRectangle`, `msoPicture`).
- `TextFrame.TextRange.Text`: Extracts text within a shape.
- `ZOrder`: Determines the stacking order.
- `Rotation`: Indicates rotation angle.

**Pseudocode**:
```pseudocode
for each shape in doc.Shapes
    print shape.Type
    print shape.Width, shape.Height, shape.Left, shape.Top
    if shape.TextFrame.HasText then
        print shape.TextFrame.TextRange.Text
    print shape.Fill.ForeColor, shape.Line.Weight
```

### 2.3. Extracting Inline Shapes
Inline shapes are embedded in the text flow and stored in the `InlineShapes` collection.

**Steps**:
1. Iterate through `doc.InlineShapes`.
2. Extract properties like `Type`, `Width`, `Height`, and `Hyperlink`.
3. Convert to floating shapes if needed using `ConvertToShape`.

**Key Properties**:
- `Type`: E.g., `msoPicture`, `msoMedia`.
- `Range`: The text range containing the inline shape.

**Pseudocode**:
```pseudocode
for each inlineShape in doc.InlineShapes
    print inlineShape.Type
    print inlineShape.Width, inlineShape.Height
```

### 2.4. Extracting Charts
Charts are special shapes (`msoChart`) that contain data and formatting.

**Steps**:
1. Access charts via `doc.Shapes` or `doc.InlineShapes` where `Type` is `msoChart`.
2. Extract chart data using `Chart.ChartData.Workbook`.
3. Extract formatting (e.g., title, axis labels, series colors).

**Key Properties**:
- `Chart.Title.Text`: Chart title.
- `Chart.ChartData.Workbook`: Access to the underlying Excel data.
- `Chart.SeriesCollection`: Data series and their properties.

**Pseudocode**:
```pseudocode
for each shape in doc.Shapes
    if shape.Type = msoChart then
        chart = shape.Chart
        print chart.Title.Text
        workbook = chart.ChartData.Workbook
        // Extract data from workbook
```

### 2.5. Extracting Grouped Shapes
Grouped shapes are collections of shapes treated as a single object, accessible via `GroupItems`.

**Steps**:
1. Check if a shape is a group (`Type = msoGroup`).
2. Iterate through `shape.GroupItems` to access individual shapes.
3. Extract properties of each sub-shape recursively.

**Key Properties**:
- `GroupItems`: Collection of shapes in the group.
- `Count`: Number of shapes in the group.

**Pseudocode**:
```pseudocode
for each shape in doc.Shapes
    if shape.Type = msoGroup then
        for each subShape in shape.GroupItems
            print subShape.Type, subShape.TextFrame.Text
```

### 2.6. Extracting Canvases
Drawing canvases are containers for multiple shapes, accessible as shapes with `Type = msoCanvas`.

**Steps**:
1. Identify canvas shapes (`msoCanvas`).
2. Access `shape.CanvasItems` to iterate through contained shapes.
3. Extract properties of each shape within the canvas.

**Key Properties**:
- `CanvasItems`: Collection of shapes in the canvas.
- `TextFrame`: Text within the canvas, if any.

**Pseudocode**:
```pseudocode
for each shape in doc.Shapes
    if shape.Type = msoCanvas then
        for each canvasShape in shape.CanvasItems
            print canvasShape.Type, canvasShape.TextFrame.Text
```

## 3. Mapping to SVG Equivalents

SVG is a vector graphics format that can represent shapes, text, and images. Below are suggestions for mapping Word elements to SVG, including embedded text.

### 3.1. General SVG Structure
An SVG document is an XML-based format with elements like `<rect>`, `<circle>`, `<text>`, `<image>`, and `<g>` (for grouping).

**Basic SVG Template**:
```xml
<svg width="width" height="height" xmlns="http://www.w3.org/2000/svg">
    <!-- Shapes, text, and groups -->
</svg>
```

### 3.2. Mapping Shapes
- **Rectangles**: Use `<rect>` with `x`, `y`, `width`, `height`, `fill`, and `stroke`.
- **Circles/Ellipses**: Use `<circle>` or `<ellipse>`.
- **Lines**: Use `<line>` with `x1`, `y1`, `x2`, `y2`.
- **Text Boxes**: Use `<text>` with `x`, `y`, and `font` attributes.
- **Images**: Use `<image>` with `xlink:href` pointing to a base64-encoded image.

**Example (Rectangle with Text)**:
```xml
<svg width="200" height="100">
    <rect x="10" y="10" width="180" height="80" fill="blue" stroke="black" stroke-width="2"/>
    <text x="20" y="50" font-family="Arial" font-size="12" fill="white">Sample Text</text>
</svg>
```

### 3.3. Mapping Charts
Charts can be approximated in SVG using `<rect>` (for bars), `<path>` (for lines), and `<text>` (for labels).

**Steps**:
1. Extract chart data (e.g., series values, categories).
2. Normalize data to SVG coordinates.
3. Draw bars, lines, or points using SVG elements.
4. Add text labels with `<text>`.

**Example (Simple Bar Chart)**:
```xml
<svg width="400" height="200">
    <rect x="50" y="100" width="50" height="80" fill="red"/>
    <rect x="120" y="50" width="50" height="130" fill="blue"/>
    <text x="50" y="190">Category 1</text>
    <text x="120" y="190">Category 2</text>
</svg>
```

### 3.4. Mapping Grouped Shapes
Use the `<g>` element to group SVG shapes, preserving their relative positions and properties.

**Example**:
```xml
<svg width="200" height="200">
    <g transform="translate(10,10)">
        <rect x="0" y="0" width="50" height="50" fill="green"/>
        <circle cx="60" cy="60" r="30" fill="yellow"/>
    </g>
</svg>
```

### 3.5. Mapping Canvases
Canvases can be mapped to an SVG `<g>` element or a nested `<svg>` for isolation.

**Example**:
```xml
<svg width="300" height="300">
    <g id="canvas">
        <rect x="10" y="10" width="100" height="100" fill="gray"/>
        <text x="20" y="50" font-family="Arial" font-size="14">Canvas Text</text>
    </g>
</svg>
```

### 3.6. Handling Text
- Use `<text>` for text in shapes, canvases, or charts.
- Map Word font properties (e.g., `Font.Name`, `Font.Size`, `Font.Color`) to SVG attributes (`font-family`, `font-size`, `fill`).
- Handle text alignment using `text-anchor` (`start`, `middle`, `end`).

**Example**:
```xml
<text x="50" y="50" font-family="Arial" font-size="16" fill="black" text-anchor="middle">Centered Text</text>
```

## 4. Practical Considerations

### 4.1. Performance
- Large documents with many shapes can be slow to process. Cache properties and avoid redundant COM calls.
- Run Word in the background (`app.Visible = false`) to improve performance.

### 4.2. Error Handling
- Handle COM errors (e.g., missing Word installation, corrupt documents).
- Check for null or invalid properties (e.g., `TextFrame` may not exist).

**Pseudocode**:
```pseudocode
try
    app = CreateObject("Word.Application")
    doc = app.Documents.Open("path/to/document.docx")
catch error
    print "Failed to open document: ", error
finally
    if doc then doc.Close()
    if app then app.Quit()
```

### 4.3. SVG Conversion Challenges
- **Coordinate Systems**: Word uses points (1/72 inch); SVG uses pixels. Convert using a scaling factor (e.g., 1 point ≈ 1.333 pixels at 96 DPI).
- **Complex Shapes**: Some Word shapes (e.g., freeform shapes) require `<path>` elements with Bézier curves.
- **Images**: Extract images as binary data and encode to base64 for `<image>` elements.
- **Fonts**: Ensure SVG font families are web-safe or embed fonts using `@font-face`.

### 4.4. Tools for SVG Generation
- Use libraries like `svg.js` or `d3.js` to programmatically generate SVG in TypeScript.
- Validate SVG output using tools like the W3C SVG Validator.

## 5. Example Workflow

**Scenario**: Extract a rectangle shape with text and a chart from a Word document and convert them to SVG.

**Steps**:
1. Open the document via COM.
2. Iterate through `doc.Shapes` to find a rectangle (`msoShapeRectangle`) and a chart (`msoChart`).
3. Extract rectangle properties (position, size, fill, text) and chart data (series, labels).
4. Generate an SVG with a `<rect>`, `<text>`, and chart elements (e.g., `<rect>` for bars).

**Pseudocode**:
```pseudocode
app = CreateObject("Word.Application")
doc = app.Documents.Open("document.docx")
svg = "<svg width='800' height='600'>"

for each shape in doc.Shapes
    if shape.Type = msoShapeRectangle then
        svg += "<rect x='" + shape.Left + "' y='" + shape.Top + "' width='" + shape.Width + "' height='" + shape.Height + "' fill='" + shape.Fill.ForeColor + "'/>"
        if shape.TextFrame.HasText then
            svg += "<text x='" + (shape.Left + shape.Width/2) + "' y='" + (shape.Top + shape.Height/2) + "' font-family='Arial' font-size='12'>" + shape.TextFrame.TextRange.Text + "</text>"
    if shape.Type = msoChart then
        chart = shape.Chart
        // Extract data and generate bars
        svg += "<rect x='100' y='100' width='50' height='100' fill='blue'/>"
        svg += "<text x='100' y='200'>" + chart.Title.Text + "</text>"

svg += "</svg>"
doc.Close()
app.Quit()
// Save svg to file
```

**Resulting SVG**:
```xml
<svg width="800" height="600">
    <rect x="100" y="100" width="200" height="100" fill="blue"/>
    <text x="200" y="150" font-family="Arial" font-size="12">Sample Text</text>
    <rect x="100" y="300" width="50" height="100" fill="blue"/>
    <text x="100" y="420">Chart Title</text>
</svg>
```

## 6. Additional Resources
- **Microsoft Word Object Model Reference**: [docs.microsoft.com](https://docs.microsoft.com/en-us/office/vba/api/overview/word)
- **SVG Specification**: [w3.org/TR/SVG2/](https://www.w3.org/TR/SVG2/)
- **TypeScript winax Documentation**: Check the `winax` GitHub repository or npm package for COM-specific guidance.
- **SVG Libraries**: Explore `svg.js` or `d3.js` for dynamic SVG generation.