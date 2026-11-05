# Images — Inline Pictures

## Key Types
- `Word.InlinePicture` — `width`, `height`, `altTextTitle`, `altTextDescription`, `hyperlink`, `paragraph`
- `Word.Body.insertInlinePictureFromBase64()` — insert image from base64 string (without `data:image/...;base64,` prefix)

## Insert Image from Base64 (With Sample Valid PNG Icon)

```javascript
await Word.run(async (context) => {
  const body = context.document.body;

  // Sample valid 64x64 blue square PNG in raw Base64 (no data:image/png;base64, prefix!)
  const sampleBase64Png =
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAdaElEQVR42u3bQREAAAjDMO5fNCCDkC4P2B0J" +
    "ECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIE" +
    "CBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgMDbAIx7AAH108b6AAAAAElFTkSuQmCC";

  const picture = body.insertInlinePictureFromBase64(
    sampleBase64Png,
    Word.InsertLocation.end
  );

  picture.width = 100;
  picture.height = 100;
  picture.altTextTitle = "Logo minh họa";
  picture.altTextDescription = "Hình ảnh minh họa được chèn tự động";
  picture.paragraph.alignment = Word.Alignment.centered;

  await context.sync();
});
```

## Get All Images

```javascript
await Word.run(async (context) => {
  const pictures = context.document.body.inlinePictures;
  pictures.load("items");
  await context.sync();

  for (const pic of pictures.items) {
    pic.load("width,height,altTextTitle");
  }
  await context.sync();

  return pictures.items.map(pic => ({
    title: pic.altTextTitle,
    width: pic.width,
    height: pic.height,
  }));
});
```

## Resize and Center All Images

```javascript
await Word.run(async (context) => {
  const pictures = context.document.body.inlinePictures;
  pictures.load("items");
  await context.sync();

  for (const pic of pictures.items) {
    pic.load("width,height");
  }
  await context.sync();

  for (const pic of pictures.items) {
    if (pic.width > 350) {
      const ratio = 350 / pic.width;
      pic.width = 350;
      pic.height = Math.round(pic.height * ratio);
    }
    pic.paragraph.alignment = Word.Alignment.centered;
  }
  await context.sync();
});
```

## Read Image Data as Base64

```javascript
await Word.run(async (context) => {
  const pictures = context.document.body.inlinePictures;
  pictures.load("items");
  await context.sync();

  if (pictures.items.length === 0) return null;

  const pic = pictures.items[0];
  // getBase64ImageSrc() returns ClientResult<string> — do NOT call .load()!
  const base64 = pic.getBase64ImageSrc();
  await context.sync();

  return base64.value;
});
```

## Delete an Image

```javascript
await Word.run(async (context) => {
  const pictures = context.document.body.inlinePictures;
  pictures.load("items");
  await context.sync();

  if (pictures.items.length > 0) {
    pictures.items[0].delete();
    await context.sync();
  }
});
```

## Common Pitfalls

- Base64 string must NOT include the `data:image/png;base64,` prefix — strip it with `.replace(/^data:image\/\w+;base64,/, "")`
- Width and height are in points (`1 pt = 1/72 inch`)
- To align an inline picture, set `pic.paragraph.alignment = Word.Alignment.centered`
- `getBase64ImageSrc()` returns `ClientResult<string>` — do NOT call `.load()`, just `await context.sync()` and read `base64.value`
