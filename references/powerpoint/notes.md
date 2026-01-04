# PowerPoint Speaker Notes

Work with speaker notes and presenter view content.

## Get Notes from Slide

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const notesSlide = slide.notesSlide;
notesSlide.load("shapes");
await context.sync();

const shapes = notesSlide.shapes;
shapes.load("items/name,items/type,items/textFrame/textRange/text");
await context.sync();

// Find the notes placeholder
const notesShape = shapes.items.find(s =>
  s.name.includes("Notes") || s.type === PowerPoint.ShapeType.placeholder
);

if (notesShape && notesShape.textFrame) {
  return notesShape.textFrame.textRange.text;
}
return "No notes found";
```

## Set Notes on Slide

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const notesSlide = slide.notesSlide;
notesSlide.load("shapes");
await context.sync();

const shapes = notesSlide.shapes;
shapes.load("items/name,items/textFrame");
await context.sync();

// Find notes placeholder and set text
const notesShape = shapes.items.find(s => s.name.includes("Notes"));
if (notesShape && notesShape.textFrame) {
  notesShape.textFrame.textRange.text = "These are my speaker notes for this slide.";
  await context.sync();
  return "Notes updated";
}
return "Notes placeholder not found";
```

## Get All Slides with Notes

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const results = [];

for (let i = 0; i < slides.items.length; i++) {
  const slide = slides.items[i];
  const notesSlide = slide.notesSlide;
  notesSlide.load("shapes");
  await context.sync();

  const shapes = notesSlide.shapes;
  shapes.load("items/name,items/textFrame/textRange/text");
  await context.sync();

  const notesShape = shapes.items.find(s =>
    s.textFrame && s.name.includes("Notes")
  );

  results.push({
    slideIndex: i,
    notes: notesShape?.textFrame?.textRange?.text || ""
  });
}

return results;
```

## Append to Notes

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const notesSlide = slide.notesSlide;
notesSlide.load("shapes");
await context.sync();

const shapes = notesSlide.shapes;
shapes.load("items/name,items/textFrame/textRange/text");
await context.sync();

const notesShape = shapes.items.find(s => s.name.includes("Notes"));
if (notesShape && notesShape.textFrame) {
  const currentText = notesShape.textFrame.textRange.text;
  notesShape.textFrame.textRange.text = currentText + "\n\nAdditional notes here.";
  await context.sync();
  return "Notes appended";
}
```

## Clear Notes

```javascript
const slide = slides.items[0];
const notesSlide = slide.notesSlide;
notesSlide.load("shapes");
await context.sync();

const shapes = notesSlide.shapes;
shapes.load("items/name,items/textFrame");
await context.sync();

const notesShape = shapes.items.find(s => s.name.includes("Notes"));
if (notesShape && notesShape.textFrame) {
  notesShape.textFrame.textRange.text = "";
  await context.sync();
  return "Notes cleared";
}
```

## Bulk Update Notes

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const notesData = [
  { slide: 0, notes: "Introduction slide notes" },
  { slide: 1, notes: "Main content notes" },
  { slide: 2, notes: "Conclusion notes" }
];

for (const data of notesData) {
  if (data.slide < slides.items.length) {
    const slide = slides.items[data.slide];
    const notesSlide = slide.notesSlide;
    notesSlide.load("shapes");
    await context.sync();

    const shapes = notesSlide.shapes;
    shapes.load("items/name,items/textFrame");
    await context.sync();

    const notesShape = shapes.items.find(s => s.name.includes("Notes"));
    if (notesShape?.textFrame) {
      notesShape.textFrame.textRange.text = data.notes;
      await context.sync();
    }
  }
}
return "All notes updated";
```

## Check if Slide Has Notes

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];
const notesSlide = slide.notesSlide;
notesSlide.load("shapes");
await context.sync();

const shapes = notesSlide.shapes;
shapes.load("items/name,items/textFrame/textRange/text");
await context.sync();

const notesShape = shapes.items.find(s => s.name.includes("Notes"));
const hasNotes = notesShape?.textFrame?.textRange?.text?.trim().length > 0;
return { hasNotes, text: notesShape?.textFrame?.textRange?.text || "" };
```

## Tips

- Notes slides are accessed via `slide.notesSlide`
- The notes placeholder name varies by template (often contains "Notes")
- Notes support basic text only through Office.js (no rich formatting)
- Each slide has its own notes slide with separate shapes
- Load shapes before accessing textFrame
