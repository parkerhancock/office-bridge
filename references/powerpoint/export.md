# PowerPoint Export

Export slides as images and other formats.

## Export Slide as Image (Bridge Helper)

The bridge provides a helper method for slide images:

```typescript
const bridge = await connect();
const presentations = await bridge.powerpoint();
const ppt = presentations[0];

// Get single slide image (0-indexed)
const image = await ppt.getSlideImage(0);
console.log(image.data);  // "data:image/png;base64,..."

// Get all slides
const allImages = await ppt.getSlideImages();
// Returns: [{ slideIndex: 0, data: "data:image/png;base64,..." }, ...]
```

## Get Slide Count

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();
return slides.items.length;
```

## Export Specific Slides

```typescript
// From calling code
const slidesToExport = [0, 2, 4];  // First, third, fifth slides
const images = [];

for (const idx of slidesToExport) {
  const image = await ppt.getSlideImage(idx);
  images.push({ index: idx, data: image.data });
}
```

## Save Image to File (Node.js)

```typescript
import { writeFileSync } from 'fs';

const image = await ppt.getSlideImage(0);
// Remove data URL prefix
const base64Data = image.data.replace(/^data:image\/png;base64,/, '');
const buffer = Buffer.from(base64Data, 'base64');
writeFileSync('slide-1.png', buffer);
```

## Export All Slides to Files

```typescript
import { writeFileSync, mkdirSync } from 'fs';

const images = await ppt.getSlideImages();
mkdirSync('slides', { recursive: true });

for (const img of images) {
  const base64Data = img.data.replace(/^data:image\/png;base64,/, '');
  const buffer = Buffer.from(base64Data, 'base64');
  writeFileSync(`slides/slide-${img.slideIndex + 1}.png`, buffer);
}
```

## Get Slide Image with Scale

```typescript
// Higher resolution export
const image = await ppt.getSlideImage(0, { scale: 2 });
// Returns larger image (2x dimensions)
```

## Export for Thumbnail

```typescript
// Smaller image for previews
const thumbnail = await ppt.getSlideImage(0, { scale: 0.5 });
```

## Create Presentation Summary

```typescript
const presentations = await bridge.powerpoint();
const ppt = presentations[0];

// Get slide images
const images = await ppt.getSlideImages();

// Get slide info via executeJs
const slideInfo = await ppt.executeJs(`
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();

  const info = [];
  for (const slide of slides.items) {
    const shapes = slide.shapes;
    shapes.load("items/name,items/textFrame/textRange/text");
    await context.sync();

    const titleShape = shapes.items.find(s => s.name.includes("Title"));
    info.push({
      title: titleShape?.textFrame?.textRange?.text || "Untitled"
    });
  }
  return info;
`);

// Combine
const summary = images.map((img, i) => ({
  index: i,
  title: slideInfo[i]?.title || "Untitled",
  image: img.data
}));
```

## Batch Export with Metadata

```typescript
async function exportPresentation(ppt) {
  const slides = await ppt.executeJs(`
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const results = [];
    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      const shapes = slide.shapes;
      shapes.load("items/name,items/textFrame/textRange/text");
      await context.sync();

      const texts = shapes.items
        .filter(s => s.textFrame)
        .map(s => s.textFrame.textRange.text)
        .filter(t => t.trim());

      results.push({
        index: i,
        textContent: texts
      });
    }
    return results;
  `);

  const images = await ppt.getSlideImages();

  return slides.map((slide, i) => ({
    ...slide,
    image: images[i].data
  }));
}
```

## Tips

- Slide images are PNG format by default
- Scale factor affects resolution (1 = native, 2 = 2x resolution)
- Image export uses the slide's current state (including animations at first frame)
- Export is useful for visual verification of changes
- Combine with executeJs to get both images and metadata
