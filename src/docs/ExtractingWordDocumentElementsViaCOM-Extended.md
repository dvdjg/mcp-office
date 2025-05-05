# Detailed Extraction of Microsoft Word Document Elements via COM

This document provides an exhaustive guide to extracting low-level information from Microsoft Word documents using COM (Component Object Model) automation. It emphasizes shapes, charts, groupings, and canvases, with a deep dive into converting native Word charts to SVG (Scalable Vector Graphics), including embedded text and data. The guide is generic, applicable to any COM-compatible programming environment (e.g., TypeScript with `winax`, Python with `pywin32`, C#), and includes enumerated values, detailed pseudocode, and practical considerations for low-level extraction and SVG conversion.

## 1. COM Automation and the Word Object Model

COM enables programmatic control of Microsoft Word through its object model, exposing objects like `Application`, `Document`, `Shapes`, `InlineShapes`, `GroupShapes`, `CanvasShapes`, and `Chart`. This guide focuses on low-level access to these objects to extract detailed properties and map them to SVG.

### 1.1. Key Objects and Their Roles
- **Application**: The Word application instance (`Word.Application`).
- **Document**: An open Word document (`Documents.Open`).
- **Shapes**: Floating shapes (e.g., rectangles, images) in the `Shapes` collection.
- **InlineShapes**: Shapes embedded in text (e.g., inline images) in the `InlineShapes` collection.
- **GroupShapes**: Shapes within a group (`GroupItems` collection).
- **CanvasShapes**: Shapes within a drawing canvas (`CanvasItems` collection).
- **Chart**: Chart objects (`msoChart`) with data and formatting.
- **TextFrame**: Text container within shapes or canvases.

### 1.2. Enumerated Values
Below are key enumerated values from the Word Object Model (based on `Mso` constants in VBA/Office). These are critical for identifying and processing elements.

#### Shape Types (`MsoShapeType`)
- `msoAutoShape` (1): Geometric shapes (e.g., rectangles, circles).
- `msoCallout` (2): Callout shapes.
- `msoChart` (3): Chart objects.
- `msoComment` (4): Comments/annotations.
- `msoFreeform` (7): Custom-drawn shapes.
- `msoGroup` (6): Grouped shapes.
- `msoEmbeddedOLEObject` (8): Embedded objects (e.g., Excel tables).
- `msoFormControl` (9): Form controls (e.g., buttons).
- `msoLine` (10): Lines or arrows.
- `msoPicture` (13): Images.
- `msoPlaceholder` (14): Placeholders in templates.
- `msoTextBox` (17): Text boxes.
- `msoCanvas` (20): Drawing canvases.

#### AutoShape Types (`MsoAutoShapeType`)
- `msoShapeRectangle` (1): Rectangle.
- `msoShapeOval` (9): Circle or ellipse.
- `msoShapeLine` (20): Straight line.
- `msoShapeArc` (25): Arc.
- `msoShapeFreeform` (5): Custom polygon.
- (Full list: ~150 types; see [MsoAutoShapeType](https://docs.microsoft.com/en-us/office/vba/api/office.msoautoshapetype)).

#### Fill Types (`MsoFillType`)
- `msoFillSolid` (0): Solid color.
- `msoFillPatterned` (1): Pattern (e.g., stripes).
- `msoFillGradient` (2): Gradient.
- `msoFillTextured` (3): Texture.
- `msoFillBackground` (4): Matches document background.
- `msoFillPicture` (5): Image fill.

#### Line Styles (`MsoLineStyle`)
- `msoLineSingle` (1): Solid line.
- `msoLineSimple` (2): Dashed line.
- `msoLineThickThin` (4): Thick-thin double line.
- (See [MsoLineStyle](https://docs.microsoft.com/en-us/office/vba/api/office.msolinestyle)).

#### Chart Types (`XlChartType`)
- `xlColumnClustered` (51): Clustered column.
- `xlColumnStacked` (52): Stacked column.
- `xlLine` (4): Line chart.
- `xlPie` (5): Pie chart.
- `xlArea` (1): Area chart.
- (See [XlChartType](https://docs.microsoft.com/en-us/office/vba/api/excel.xlcharttype)).

### 1.3. Prerequisites
- Microsoft Word installed.
- COM-compatible environment (e.g., TypeScript with `winax`).
- Familiarity with Word Object Model (see [Microsoft Word VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/word)).
- SVG knowledge for conversion (see [SVG 2 Specification](https://www.w3.org/TR/SVG2/)).

## 2. Low-Level Extraction Process

This section details the steps to extract shapes, charts, groupings, and canvases, including all relevant properties and enumerated values.

### 2.1. Opening and Accessing a Document
**Steps**:
1. Create a `Word.Application` instance.
2. Set `Visible = false` for background processing.
3. Open the document using `Documents.Open(FileName, ReadOnly: true)`.
4. Access collections (`Shapes`, `InlineShapes`, etc.).
5. Close the document and quit the application.

**Pseudocode**:
```pseudocode
app = CreateObject("Word.Application")
app.Visible = false
doc = app.Documents.Open("path/to/document.docx", ReadOnly: true)
// Process elements (see below)
doc.Close(SaveChanges: false)
app.Quit()
```

**Properties to Extract**:
- `Document.Name`: Document filename.
- `Document.Path`: File path.
- `Document.PageSetup`: Page dimensions (`Width`, `Height`).

### 2.2. Extracting Shapes
Shapes are floating objects in the `Shapes` collection. Each shape has properties defining its type, position, appearance, and content.

**Steps**:
1. Iterate through `doc.Shapes`.
2. Check `Type` to identify the shape (e.g., `msoAutoShape`, `msoChart`).
3. Extract geometric properties (`Left`, `Top`, `Width`, `Height`, `Rotation`).
4. Extract appearance (`Fill`, `Line`).
5. Extract text via `TextFrame.TextRange` if `HasText = true`.
6. Handle special cases (e.g., images, freeforms).

**Key Properties**:
- `Type`: See `MsoShapeType` above.
- `AutoShapeType`: For `msoAutoShape` shapes (e.g., `msoShapeRectangle`).
- `Left`, `Top`: Position in points (relative to page top-left).
- `Width`, `Height`: Size in points.
- `Rotation`: Angle in degrees.
- `Fill.Type`: See `MsoFillType` (e.g., `msoFillSolid`).
- `Fill.ForeColor.RGB`: Fill color (RGB value).
- `Line.Style`: See `MsoLineStyle`.
- `Line.Weight`: Line thickness in points.
- `TextFrame.TextRange.Text`: Text content.
- `TextFrame.TextRange.Font`: Font properties (`Name`, `Size`, `Color`).
- `ZOrderPosition`: Stacking order.

**Pseudocode**:
```pseudocode
for each shape in doc.Shapes
    print "Type:", shape.Type
    if shape.Type = msoAutoShape then
        print "AutoShapeType:", shape.AutoShapeType
    print "Position:", shape.Left, shape.Top
    print "Size:", shape.Width, shape.Height
    print "Rotation:", shape.Rotation
    print "Fill:", shape.Fill.Type, shape.Fill.ForeColor.RGB
    print "Line:", shape.Line.Style, shape.Line.Weight
    if shape.TextFrame and shape.TextFrame.HasText then
        print "Text:", shape.TextFrame.TextRange.Text
        print "Font:", shape.TextFrame.TextRange.Font.Name, shape.TextFrame.TextRange.Font.Size
```

### 2.3. Extracting Inline Shapes
Inline shapes are embedded in the text flow and stored in `InlineShapes`.

**Steps**:
1. Iterate through `doc.InlineShapes`.
2. Check `Type` (e.g., `msoPicture`, `msoChart`).
3. Extract properties (`Width`, `Height`, `Range`).
4. Convert to floating shape if needed (`ConvertToShape`).

**Key Properties**:
- `Type`: See `MsoShapeType`.
- `Width`, `Height`: Size in points.
- `Range`: Text range containing the shape.
- `Hyperlink.Address`: If the shape is a hyperlink.

**Pseudocode**:
```pseudocode
for each inlineShape in doc.InlineShapes
    print "Type:", inlineShape.Type
    print "Size:", inlineShape.Width, inlineShape.Height
    print "Range:", inlineShape.Range.Text
```

### 2.4. Extracting Grouped Shapes
Grouped shapes (`msoGroup`) contain multiple shapes in `GroupItems`.

**Steps**:
1. Identify groups (`Type = msoGroup`).
2. Iterate through `shape.GroupItems`.
3. Extract properties recursively for each sub-shape.

**Key Properties**:
- `GroupItems.Count`: Number of shapes.
- `GroupItems.Item(Index)`: Access individual shapes.
- Same properties as regular shapes for sub-shapes.

**Pseudocode**:
```pseudocode
for each shape in doc.Shapes
    if shape.Type = msoGroup then
        print "Group with", shape.GroupItems.Count, "items"
        for each subShape in shape.GroupItems
            print "SubShape Type:", subShape.Type
            print "SubShape Position:", subShape.Left, subShape.Top
            if subShape.TextFrame and subShape.TextFrame.HasText then
                print "SubShape Text:", subShape.TextFrame.TextRange.Text
```

### 2.5. Extracting Canvases
Drawing canvases (`msoCanvas`) are containers for shapes, stored in `CanvasItems`.

**Steps**:
1. Identify canvases (`Type = msoCanvas`).
2. Iterate through `shape.CanvasItems`.
3. Extract properties for each contained shape.

**Key Properties**:
- `CanvasItems.Count`: Number of shapes.
- `CanvasItems.Item(Index)`: Access individual shapes.
- Same properties as regular shapes.

**Pseudocode**:
```pseudocode
for each shape in doc.Shapes
    if shape.Type = msoCanvas then
        print "Canvas with", shape.CanvasItems.Count, "items"
        for each canvasShape in shape.CanvasItems
            print "CanvasShape Type:", canvasShape.Type
            print "CanvasShape Position:", canvasShape.Left, canvasShape.Top
            if canvasShape.TextFrame and canvasShape.TextFrame.HasText then
                print "CanvasShape Text:", canvasShape.TextRange.Text
```

### 2.6. Extracting Charts (Low-Level)
Charts (`msoChart`) are complex objects with data stored in an embedded Excel workbook and visual elements defined by properties.

**Steps**:
1. Identify charts (`Type = msoChart` in `Shapes` or `InlineShapes`).
2. Access `Chart` object (`shape.Chart`).
3. Extract chart type (`Chart.ChartType`).
4. Extract data via `Chart.ChartData.Workbook`.
5. Extract visual properties (e.g., title, legend, axes).
6. Extract series data and formatting.

**Key Properties**:
- `Chart.ChartType`: See `XlChartType` (e.g., `xlColumnClustered`).
- `Chart.Title.Text`: Chart title.
- `Chart.Legend`: Legend properties (`Position`, `Font`).
- `Chart.Axes`: Axis properties (`Title`, `Scale`).
- `Chart.SeriesCollection`: Data series (`Name`, `Values`, `XValues`).
- `Chart.ChartData.Workbook`: Embedded Excel workbook.
- `Chart.ChartData.Workbook.Worksheets(1).Range`: Data range.

**Data Extraction**:
- Activate the workbook: `Chart.ChartData.Activate`.
- Access the worksheet: `Chart.ChartData.Workbook.Worksheets(1)`.
- Read data ranges (e.g., `Range("A1:B10").Value`).

**Pseudocode**:
```pseudocode
for each shape in doc.Shapes
    if shape.Type = msoChart then
        chart = shape.Chart
        print "Chart Type:", chart.ChartType
        print "Title:", chart.Title.Text
        if chart.HasLegend then
            print "Legend Position:", chart.Legend.Position
        // Extract data
        chart.ChartData.Activate
        worksheet = chart.ChartData.Workbook.Worksheets(1)
        // Assume data in A1:B10 (adjust dynamically)
        data = worksheet.Range("A1:B10").Value
        for each row in data
            print "Data Row:", row
        // Extract series
        for each series in chart.SeriesCollection
            print "Series Name:", series.Name
            print "Values:", series.Values
            print "XValues:", series.XValues
```

## 3. Converting Native Charts to SVG

Converting Word charts to SVG requires extracting both data and visual properties, then rendering them using SVG elements (`<rect>`, `<path>`, `<text>`, etc.). Below is a detailed process for this conversion.

### 3.1. SVG Coordinate System
- **Word Units**: Points (1 point = 1/72 inch).
- **SVG Units**: Pixels (typically 1 point ≈ 1.333 pixels at 96 DPI).
- **Conversion**: Multiply Word coordinates by 1.333 (adjust for target DPI).

**Example**:
```pseudocode
svgX = wordLeft * 1.333
svgY = wordTop * 1.333
svgWidth = wordWidth * 1.333
svgHeight = wordHeight * 1.333
```

### 3.2. Chart Components
Charts have the following components to map to SVG:
- **Data Series**: Values plotted (e.g., bars, lines).
- **Axes**: X and Y axes with labels and scales.
- **Title**: Chart title.
- **Legend**: Series labels.
- **Plot Area**: The area containing the data visualization.

### 3.3. Conversion Process
1. **Extract Chart Data**:
   - Use `Chart.ChartData.Workbook` to get raw data.
   - Identify categories (X-axis) and values (Y-axis).
   - Example: For a column chart, categories in column A, values in column B.

2. **Extract Visual Properties**:
   - Chart type (`Chart.ChartType`).
   - Title (`Chart.Title.Text`, `Chart.Title.Font`).
   - Axis labels (`Chart.Axes(xlCategory).Title`, `Chart.Axes(xlValue).Title`).
   - Series colors (`Series.Format.Fill.ForeColor`).
   - Plot area size (`Chart.PlotArea.Width`, `Chart.PlotArea.Height`).

3. **Normalize Data**:
   - Determine min/max values for Y-axis scaling.
   - Map data to SVG coordinates within the plot area.

4. **Render SVG Elements**:
   - Use `<rect>` for column/bar charts.
   - Use `<path>` for line/area charts.
   - Use `<circle>` for scatter charts.
   - Use `<text>` for labels, title, and legend.
   - Use `<g>` to group elements (e.g., plot area, axes).

### 3.4. Example: Column Chart Conversion
**Scenario**: Convert a clustered column chart with two series to SVG.

**Word Chart Properties**:
- `ChartType`: `xlColumnClustered`.
- `Title`: "Sales Data".
- `Categories`: ["Q1", "Q2", "Q3"] (X-axis).
- `Series 1`: "Product A", values [100, 150, 200].
- `Series 2`: "Product B", values [80, 120, 180].
- `PlotArea`: 300x200 points.
- `Series Colors`: Blue (RGB: 0,0,255), Red (RGB: 255,0,0).

**Steps**:
1. **Extract Data**:
   ```pseudocode
   chart.ChartData.Activate
   worksheet = chart.ChartData.Workbook.Worksheets(1)
   categories = worksheet.Range("A2:A4").Value  // ["Q1", "Q2", "Q3"]
   series1 = worksheet.Range("B2:B4").Value    // [100, 150, 200]
   series2 = worksheet.Range("C2:C4").Value    // [80, 120, 180]
   ```

2. **Normalize Data**:
   - Y-axis range: Min = 0, Max = 200 (round up to 250 for scale).
   - Plot area: 300x200 points → 400x266.6 pixels (1.333 scaling).
   - Bar width: Divide plot width by number of categories (400 / 3 ≈ 133 pixels per category).
   - Bar spacing: Allocate 50% for bars (66.5 pixels per category, ~33 pixels per bar).

3. **Map to SVG**:
   - Plot area: `<g>` at (50,50) for padding.
   - Bars: `<rect>` for each value, scaled to Y-axis.
   - X-axis labels: `<text>` below bars.
   - Y-axis labels: `<text>` along left side.
   - Title: `<text>` above plot area.

**SVG Output**:
```xml
<svg width="500" height="400" xmlns="http://www.w3.org/2000/svg">
    <!-- Title -->
    <text x="250" y="30" font-family="Arial" font-size="16" text-anchor="middle">Sales Data</text>
    <!-- Plot Area -->
    <g transform="translate(50,50)">
        <!-- Y-axis labels -->
        <text x="-30" y="0" font-family="Arial" font-size="12">250</text>
        <text x="-30" y="66.6" font-family="Arial" font-size="12">167</text>
        <text x="-30" y="133.2" font-family="Arial" font-size="12">83</text>
        <text x="-30" y="199.8" font-family="Arial" font-size="12">0</text>
        <!-- Bars -->
        <!-- Series 1: Product A -->
        <rect x="33" y="99.9" width="33" height="99.9" fill="rgb(0,0,255)"/>  <!-- Q1: 100 -->
        <rect x="166" y="49.95" width="33" height="149.85" fill="rgb(0,0,255)"/>  <!-- Q2: 150 -->
        <rect x="299" y="0" width="33" height="199.8" fill="rgb(0,0,255)"/>  <!-- Q3: 200 -->
        <!-- Series 2: Product B -->
        <rect x="66" y="119.88" width="33" height="79.92" fill="rgb(255,0,0)"/>  <!-- Q1: 80 -->
        <rect x="199" y="79.92" width="33" height="119.88" fill="rgb(255,0,0)"/>  <!-- Q2: 120 -->
        <rect x="332" y="19.98" width="33" height="179.82" fill="rgb(255,0,0)"/>  <!-- Q3: 180 -->
        <!-- X-axis labels -->
        <text x="66" y="220" font-family="Arial" font-size="12" text-anchor="middle">Q1</text>
        <text x="199" y="220" font-family="Arial" font-size="12" text-anchor="middle">Q2</text>
        <text x="332" y="220" font-family="Arial" font-size="12" text-anchor="middle">Q3</text>
    </g>
    <!-- Legend -->
    <g transform="translate(400,50)">
        <rect x="0" y="0" width="15" height="15" fill="rgb(0,0,255)"/>
        <text x="20" y="12" font-family="Arial" font-size="12">Product A</text>
        <rect x="0" y="20" width="15" height="15" fill="rgb(255,0,0)"/>
        <text x="20" y="32" font-family="Arial" font-size="12">Product B</text>
    </g>
</svg>
```

**Calculations**:
- Y-axis scaling: 200 points height → 199.8 pixels (max value 250 → 0.7992 pixels per unit).
  - Q1 Product A: 100 * 0.7992 = 79.92 pixels.
  - Q2 Product A: 150 * 0.7992 = 119.88 pixels.
- X-axis: 400 pixels / 3 categories = 133 pixels per category.
  - Bar width: 33 pixels (half of 66.5 pixels per series).
  - Positions: Q1 at x=33, Q2 at x=166, Q3 at x=299.

### 3.5. Other Chart Types
- **Line Chart** (`xlLine`):
  - Use `<path>` with `d` attribute for lines.
  - Example: `<path d="M33,99.9 L166,49.95 L299,0" stroke="blue" fill="none"/>`.
- **Pie Chart** (`xlPie`):
  - Use `<path>` with arc commands (`A`) to draw wedges.
  - Calculate angles: Total value = sum of series; each slice = (value/total) * 360°.
- **Scatter Chart**:
  - Use `<circle>` for points.
  - Map X/Y values to SVG coordinates.

### 3.6. Handling Text in Charts
- **Title**: Map `Chart.Title.Font` to SVG `font-family`, `font-size`, `fill`.
- **Axis Labels**: Use `<text>` with `text-anchor` for alignment.
- **Data Labels**: Extract `Series.DataLabels.Text` and position near bars/points.

**Example**:
```xml
<text x="50" y="30" font-family="Arial" font-size="16" fill="black">Chart Title</text>
```

## 4. Mapping Other Elements to SVG

### 4.1. Shapes
- **Rectangle** (`msoShapeRectangle`):
  ```xml
  <rect x="left*1.333" y="top*1.333" width="width*1.333" height="height*1.333" fill="rgb(r,g,b)" stroke="rgb(r,g,b)" stroke-width="lineWeight*1.333"/>
  ```
- **Text Box** (`msoTextBox`):
  ```xml
  <text x="left*1.333" y="top*1.333" font-family="fontName" font-size="fontSize*1.333" fill="rgb(r,g,b)">textContent</text>
  ```
- **Image** (`msoPicture`):
  - Extract image data via `shape.Copy` to clipboard or save as file.
  - Convert to base64 for `<image>`: `<image x="left*1.333" y="top*1.333" width="width*1.333" height="height*1.333" xlink:href="data:image/png;base64,..."/>`.

### 4.2. Grouped Shapes
- Use `<g>` to group shapes.
- Apply `transform="translate(left*1.333,top*1.333)"` for group position.
- Example:
  ```xml
  <g transform="translate(left*1.333,top*1.333)">
      <rect x="0" y="0" width="50" height="50" fill="green"/>
      <circle cx="60" cy="60" r="30" fill="yellow"/>
  </g>
  ```

### 4.3. Canvases
- Map to `<g>` or nested `<svg>` for isolation.
- Example:
  ```xml
  <g id="canvas">
      <rect x="10" y="10" width="100" height="100" fill="gray"/>
      <text x="20" y="50" font-family="Arial" font-size="14">Canvas Text</text>
  </g>
  ```

## 5. Practical Considerations

### 5.1. Performance Optimization
- **Minimize COM Calls**: Cache properties (e.g., `shape.TextFrame`) to avoid repeated queries.
- **Batch Processing**: Process shapes in chunks for large documents.
- **Background Execution**: Set `app.Visible = false` and `app.ScreenUpdating = false`.

### 5.2. Error Handling
- **COM Errors**: Handle exceptions for missing Word, corrupt files, or invalid properties.
- **Null Checks**: Verify `TextFrame`, `ChartData`, etc., before access.
- **Cleanup**: Ensure `doc.Close()` and `app.Quit()` in `finally` block.

**Pseudocode**:
```pseudocode
try
    app = CreateObject("Word.Application")
    app.Visible = false
    doc = app.Documents.Open("path/to/document.docx")
    // Process elements
catch error
    print "Error:", error
finally
    if doc then doc.Close(SaveChanges: false)
    if app then app.Quit()
```

### 5.3. SVG Conversion Challenges
- **Coordinate Precision**: Round SVG coordinates to 2 decimals to avoid rendering issues.
- **Complex Shapes**: Freeform shapes (`msoFreeform`) require parsing `Nodes` to generate `<path>` elements.
- **Fonts**: Map Word fonts to web-safe fonts (e.g., Arial, Times New Roman) or use `@font-face`.
- **Gradients**: Word gradients (`msoFillGradient`) need SVG `<linearGradient>` or `<radialGradient>`.

### 5.4. Tools and Libraries
- **SVG Generation**: Use `svg.js` or `d3.js` in TypeScript for dynamic SVG creation.
- **Image Processing**: Convert Word images to base64 using external libraries (e.g., `image-to-base64`).
- **Validation**: Test SVG with W3C SVG Validator.

## 6. Comprehensive Example Workflow

**Scenario**: Extract a rectangle with text, a grouped shape, and a column chart from a Word document, converting them to SVG.

**Pseudocode**:
```pseudocode
app = CreateObject("Word.Application")
app.Visible = false
doc = app.Documents.Open("document.docx")
svg = "<svg width='800' height='600' xmlns='http://www.w3.org/2000/svg'>"

for each shape in doc.Shapes
    if shape.Type = msoShapeRectangle then
        svg += "<rect x='" + (shape.Left*1.333) + "' y='" + (shape.Top*1.333) + "' width='" + (shape.Width*1.333) + "' height='" + (shape.Height*1.333) + "' fill='rgb(" + shape.Fill.ForeColor.RGB + ")' stroke='rgb(" + shape.Line.ForeColor.RGB + ")' stroke-width='" + (shape.Line.Weight*1.333) + "'/>"
        if shape.TextFrame and shape.TextFrame.HasText then
            svg += "<text x='" + (shape.Left*1.333 + shape.Width*1.333/2) + "' y='" + (shape.Top*1.333 + shape.Height*1.333/2) + "' font-family='" + shape.TextFrame.TextRange.Font.Name + "' font-size='" + (shape.TextFrame.TextRange.Font.Size*1.333) + "' text-anchor='middle'>" + shape.TextFrame.TextRange.Text + "</text>"
    if shape.Type = msoGroup then
        svg += "<g transform='translate(" + (shape.Left*1.333) + "," + (shape.Top*1.333) + ")'>"
        for each subShape in shape.GroupItems
            if subShape.Type = msoAutoShape and subShape.AutoShapeType = msoShapeOval then
                svg += "<circle cx='" + (subShape.Left*1.333 + subShape.Width*1.333/2) + "' cy='" + (subShape.Top*1.333 + subShape.Height*1.333/2) + "' r='" + (subShape.Width*1.333/2) + "' fill='rgb(" + subShape.Fill.ForeColor.RGB + ")'/>"
        svg += "</g>"
    if shape.Type = msoChart then
        chart = shape.Chart
        plotWidth = chart.PlotArea.Width * 1.333
        plotHeight = chart.PlotArea.Height * 1.333
        svg += "<g transform='translate(" + (shape.Left*1.333) + "," + (shape.Top*1.333) + ")'>"
        // Extract data
        chart.ChartData.Activate
        worksheet = chart.ChartData.Workbook.Worksheets(1)
        categories = worksheet.Range("A2:A4").Value
        series1 = worksheet.Range("B2:B4").Value
        maxValue = max(series1)
        for i = 1 to length(categories)
            x = (i-1) * (plotWidth / length(categories))
            height = (series1[i] / maxValue) * plotHeight
            svg += "<rect x='" + x + "' y='" + (plotHeight-height) + "' width='" + (plotWidth/length(categories)/2) + "' height='" + height + "' fill='rgb(0,0,255)'/>"
            svg += "<text x='" + (x + plotWidth/length(categories)/2) + "' y='" + (plotHeight+20) + "' font-family='Arial' font-size='12' text-anchor='middle'>" + categories[i] + "</text>"
        svg += "<text x='" + (plotWidth/2) + "' y='-10' font-family='Arial' font-size='16' text-anchor='middle'>" + chart.Title.Text + "</text>"
        svg += "</g>"

svg += "</svg>"
doc.Close(SaveChanges: false)
app.Quit()
// Save svg to file
```

**Resulting SVG**:
```xml
<svg width="800" height="600" xmlns="http://www.w3.org/2000/svg">
    <!-- Rectangle with Text -->
    <rect x="133.3" y="133.3" width="266.6" height="133.3" fill="rgb(0,0,255)" stroke="rgb(0,0,0)" stroke-width="1.333"/>
    <text x="266.6" y="199.95" font-family="Arial" font-size="16" text-anchor="middle">Sample Text</text>
    <!-- Grouped Shapes -->
    <g transform="translate(266.6,266.6)">
        <circle cx="66.65" cy="66.65" r="33.325" fill="rgb(255,255,0)"/>
    </g>
    <!-- Column Chart -->
    <g transform="translate(133.3,399.9)">
        <rect x="0" y="133.2" width="66.6" height="66.6" fill="rgb(0,0,255)"/>  <!-- Q1 -->
        <rect x="133.2" y="66.6" width="66.6" height="133.2" fill="rgb(0,0,255)"/>  <!-- Q2 -->
        <rect x="266.4" y="0" width="66.6" height="199.8" fill="rgb(0,0,255)"/>  <!-- Q3 -->
        <text x="66.6" y="219.8" font-family="Arial" font-size="12" text-anchor="middle">Q1</text>
        <text x="199.8" y="219.8" font-family="Arial" font-size="12" text-anchor="middle">Q2</text>
        <text x="333" y="219.8" font-family="Arial" font-size="12" text-anchor="middle">Q3</text>
        <text x="199.8" y="-10" font-family="Arial" font-size="16" text-anchor="middle">Sales Data</text>
    </g>
</svg>
```

## 7. Additional Resources
- **Word Object Model**: [Microsoft VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/word).
- **Chart Object**: [Excel VBA Chart Reference](https://docs.microsoft.com/en-us/office/vba/api/excel.chart).
- **SVG Specification**: [W3C SVG 2](https://www.w3.org/TR/SVG2/).
- **Winax Documentation**: Check npm or GitHub for TypeScript COM specifics.
- **SVG Tools**: `svg.js`, `d3.js` for SVG generation.

This guide provides a low-level, detailed approach to extracting and converting Word document elements, with a focus on charts and SVG mapping, suitable for implementation in any COM-compatible environment.