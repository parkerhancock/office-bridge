# PowerPoint Shapes

Create, manipulate, and style shapes.

## Add Text Box

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const textBox = slide.shapes.addTextBox("Hello from Claude!", {
  left: 100,
  top: 100,
  width: 300,
  height: 50
});
await context.sync();
return "Text box added";
```

## Add Geometric Shape

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const shape = slide.shapes.addGeometricShape(
  PowerPoint.GeometricShapeType.rectangle,
  {
    left: 100,
    top: 200,
    width: 200,
    height: 100
  }
);
await context.sync();
return "Shape added";
```

## Common Shape Types

```javascript
PowerPoint.GeometricShapeType.rectangle
PowerPoint.GeometricShapeType.ellipse
PowerPoint.GeometricShapeType.triangle
PowerPoint.GeometricShapeType.rightTriangle
PowerPoint.GeometricShapeType.diamond
PowerPoint.GeometricShapeType.pentagon
PowerPoint.GeometricShapeType.hexagon
PowerPoint.GeometricShapeType.star5
PowerPoint.GeometricShapeType.arrow
PowerPoint.GeometricShapeType.line
PowerPoint.GeometricShapeType.roundedRectangle
```

## Get All Shapes

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const shapes = slide.shapes;
shapes.load("items/name,items/type,items/left,items/top,items/width,items/height");
await context.sync();

return shapes.items.map(s => ({
  name: s.name,
  type: s.type,
  position: { left: s.left, top: s.top },
  size: { width: s.width, height: s.height }
}));
```

## Modify Shape Properties

### Position and Size

```javascript
const shapes = slide.shapes;
shapes.load("items/name");
await context.sync();

const shape = shapes.items.find(s => s.name === "MyShape");
shape.left = 200;
shape.top = 150;
shape.width = 300;
shape.height = 200;
await context.sync();
```

### Rotation

```javascript
shape.rotation = 45;  // Degrees
await context.sync();
```

## Shape Text

### Set Text

```javascript
const shapes = slide.shapes;
shapes.load("items/name,items/textFrame");
await context.sync();

const shape = shapes.items.find(s => s.name === "Rectangle 1");
shape.textFrame.textRange.text = "New text content";
await context.sync();
```

### Get Text

```javascript
const shapes = slide.shapes;
shapes.load("items/name,items/textFrame/textRange/text");
await context.sync();

return shapes.items
  .filter(s => s.textFrame)
  .map(s => ({
    name: s.name,
    text: s.textFrame.textRange.text
  }));
```

## Shape Fill

```javascript
const shape = shapes.items[0];
shape.fill.setSolidColor("#FF5733");
await context.sync();
```

## Shape Line (Border)

```javascript
const shape = shapes.items[0];
shape.lineFormat.color = "#000000";
shape.lineFormat.weight = 2;  // Points
await context.sync();
```

## Add Line

```javascript
const slide = slides.items[0];
const line = slide.shapes.addLine(
  PowerPoint.ConnectorType.straight,
  {
    left: 100,
    top: 100,
    width: 200,
    height: 0  // Horizontal line
  }
);
await context.sync();
```

## Delete Shape

```javascript
const shapes = slide.shapes;
shapes.load("items/name");
await context.sync();

const shape = shapes.items.find(s => s.name === "Shape to Delete");
shape.delete();
await context.sync();
return "Shape deleted";
```

## Group Shapes

```javascript
const slide = slides.items[0];
const shapes = slide.shapes;
shapes.load("items/name,items/id");
await context.sync();

// Select shapes to group (by their IDs)
const shapeIds = shapes.items
  .filter(s => s.name.startsWith("Group"))
  .map(s => s.id);

// Note: Grouping requires specific API version support
// Check PowerPoint.GroupShapeCollection if available
```

## Z-Order (Layering)

```javascript
// Bring to front
shape.setZOrder(PowerPoint.ShapeZOrder.bringToFront);

// Send to back
shape.setZOrder(PowerPoint.ShapeZOrder.sendToBack);

// Bring forward one level
shape.setZOrder(PowerPoint.ShapeZOrder.bringForward);

// Send backward one level
shape.setZOrder(PowerPoint.ShapeZOrder.sendBackward);

await context.sync();
```

## Tips

- Position uses points (72 points = 1 inch)
- Shape names are auto-generated like "Rectangle 1", "TextBox 2"
- Use `shapes.getItem(name)` if you know the exact name
- Text frames may not exist on all shapes (lines, connectors)
- Colors use hex format: `"#RRGGBB"`
