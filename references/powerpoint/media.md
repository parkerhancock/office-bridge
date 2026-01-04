# PowerPoint Media

Images, videos, and audio.

## Add Image from Base64

```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

const slide = slides.items[0];

// Base64 image data (without data:image/png;base64, prefix)
const base64Data = "iVBORw0KGgoAAAANSUhEUg...";

const image = slide.shapes.addImage(base64Data, {
  left: 100,
  top: 100,
  width: 400,
  height: 300
});
await context.sync();
return "Image added";
```

## Image from URL (via fetch)

```javascript
// First fetch the image and convert to base64
// This would be done in your calling code, then passed to executeJs

const slide = slides.items[0];
const image = slide.shapes.addImage(base64ImageData, {
  left: 50,
  top: 50,
  width: 500,
  height: 375
});
await context.sync();
```

## Get Image Dimensions

```javascript
const shapes = slide.shapes;
shapes.load("items/name,items/type,items/width,items/height");
await context.sync();

const images = shapes.items.filter(s =>
  s.type === PowerPoint.ShapeType.image
);
return images.map(img => ({
  name: img.name,
  width: img.width,
  height: img.height
}));
```

## Resize Image Proportionally

```javascript
const shapes = slide.shapes;
shapes.load("items/name,items/type,items/width,items/height");
await context.sync();

const image = shapes.items.find(s => s.name === "Picture 1");
const aspectRatio = image.width / image.height;

// Set new width, calculate height
const newWidth = 500;
image.width = newWidth;
image.height = newWidth / aspectRatio;
await context.sync();
```

## Center Image on Slide

```javascript
const presentation = context.presentation;
presentation.load("slideWidth,slideHeight");
await context.sync();

const shapes = slide.shapes;
shapes.load("items/name,items/width,items/height");
await context.sync();

const image = shapes.items.find(s => s.name === "Picture 1");
image.left = (presentation.slideWidth - image.width) / 2;
image.top = (presentation.slideHeight - image.height) / 2;
await context.sync();
return "Image centered";
```

## Replace Image

```javascript
// Delete old image and add new one
const shapes = slide.shapes;
shapes.load("items/name,items/type,items/left,items/top,items/width,items/height");
await context.sync();

const oldImage = shapes.items.find(s => s.name === "Picture 1");
const position = {
  left: oldImage.left,
  top: oldImage.top,
  width: oldImage.width,
  height: oldImage.height
};

oldImage.delete();
await context.sync();

slide.shapes.addImage(newBase64Data, position);
await context.sync();
return "Image replaced";
```

## Add Video (Limited Support)

```javascript
// Video support is limited in Office.js
// For full video support, consider using VBA or desktop automation

// Basic approach - add as embedded object
// Note: This may not work in all versions
const slide = slides.items[0];

// Video embedding typically requires:
// 1. Converting video to supported format
// 2. Using Office Add-in with custom UI
// 3. Or embedding via OLE object
```

## Add Audio (Limited Support)

```javascript
// Audio embedding has similar limitations to video
// Consider these alternatives:
// 1. Link to external audio file
// 2. Use hyperlink to audio resource
// 3. Desktop automation for full support
```

## Image as Slide Background

```javascript
// Note: Direct background image API may be limited
// Alternative: Add full-slide image and send to back

const slide = slides.items[0];
const presentation = context.presentation;
presentation.load("slideWidth,slideHeight");
await context.sync();

const bgImage = slide.shapes.addImage(base64Data, {
  left: 0,
  top: 0,
  width: presentation.slideWidth,
  height: presentation.slideHeight
});

bgImage.setZOrder(PowerPoint.ShapeZOrder.sendToBack);
await context.sync();
return "Background set";
```

## Export Slide as Image

```javascript
// Use the bridge's built-in slide image helper
const presentations = await bridge.powerpoint();
const ppt = presentations[0];

// Get specific slide image
const slideImage = await ppt.getSlideImage(0);  // 0-indexed
// Returns { data: "data:image/png;base64,..." }

// Get all slide images
const allImages = await ppt.getSlideImages();
// Returns array of { slideIndex, data }
```

## Tips

- Base64 images should not include the `data:image/...;base64,` prefix
- Supported image formats: PNG, JPEG, GIF, BMP
- Video/audio support is limited in Office.js API
- Use slide images export for visual verification
- Position and size in points (72 = 1 inch)
