# PowerPoint Layouts

Slide layouts, masters, and placeholders.

## Get Available Layouts

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const layout = slide.layout;
layout.load("name,id");
await context.sync();
return { name: layout.name, id: layout.id };
```

## Add Slide with Layout

```javascript
// Add slide at the end
const newSlide = context.presentation.slides.add();
await context.sync();

// Get slide layouts from master
const slideMasters = context.presentation.slideMasters;
slideMasters.load("items");
await context.sync();

const master = slideMasters.items[0];
const layouts = master.layouts;
layouts.load("items/name,items/id");
await context.sync();

return layouts.items.map(l => ({ name: l.name, id: l.id }));
```

## Common Layout Names

Standard PowerPoint layouts:
- `Title Slide` - Title and subtitle
- `Title and Content` - Title with content area
- `Section Header` - Section divider
- `Two Content` - Title with two content areas
- `Comparison` - Two columns with headers
- `Title Only` - Just a title
- `Blank` - Empty slide
- `Content with Caption` - Content with side caption
- `Picture with Caption` - Image with caption

## Working with Placeholders

### Get Placeholders

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const shapes = slide.shapes;
shapes.load("items/name,items/type,items/placeholderType");
await context.sync();

const placeholders = shapes.items.filter(s =>
  s.type === PowerPoint.ShapeType.placeholder
);
return placeholders.map(p => ({
  name: p.name,
  type: p.placeholderType
}));
```

### Placeholder Types

- `Title` - Slide title
- `Subtitle` - Subtitle on title slide
- `Body` - Main content area
- `SlideNumber` - Slide number
- `Footer` - Footer text
- `Header` - Header text
- `DateAndTime` - Date placeholder

### Set Title Text

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const shapes = slide.shapes;
shapes.load("items/name,items/textFrame");
await context.sync();

const title = shapes.items.find(s => s.name.includes("Title"));
if (title) {
  title.textFrame.textRange.text = "New Title Here";
  await context.sync();
  return "Title updated";
}
return "No title found";
```

## Slide Size

```javascript
// Get slide dimensions
const presentation = context.presentation;
presentation.load("slideWidth,slideHeight");
await context.sync();
return {
  width: presentation.slideWidth,   // Points
  height: presentation.slideHeight  // Points
};
// Standard 16:9 = 914.4 x 514.35 points
// Standard 4:3 = 685.8 x 514.35 points
```

## Duplicate Slide

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

// Get first slide
const sourceSlide = slides.items[0];

// Add new slide (will copy from template)
const newSlide = slides.add();
await context.sync();

return "Slide duplicated";
```

## Reorder Slides

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

// Move last slide to position 1 (second position, 0-indexed)
const lastSlide = slides.items[slides.items.length - 1];
lastSlide.moveTo(1);
await context.sync();
return "Slide moved";
```

## Tips

- Slide dimensions are in points (72 points = 1 inch)
- Layout names may vary by language/template
- Use shape names containing "Title", "Subtitle", "Content" to find placeholders
- PowerPoint API layout support is limited compared to desktop VBA
