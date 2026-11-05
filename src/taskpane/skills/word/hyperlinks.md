# Hyperlinks — Insert, Read, and Remove Links

## Key Rules
- `range.hyperlink` is a string property on **`Word.Range`** (or `Word.Paragraph`).
- **CRITICAL**: `getHyperlinkRanges()` is on **`Word.Range`** (`context.document.body.getRange().getHyperlinkRanges()`), NOT on `Word.Body`!

## Insert a Hyperlink and List All Hyperlinks

```javascript
await Word.run(async (context) => {
  const body = context.document.body;
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();

  if (selection.text && selection.text.trim()) {
    selection.hyperlink = "https://office.com";
  } else {
    const p = body.insertParagraph("Tham khảo cổng thông tin điện tử tại đây", Word.InsertLocation.end);
    p.font.name = "Times New Roman";
    p.font.size = 13;
    const range = p.getRange();
    range.hyperlink = "https://chinhphu.vn";
  }
  await context.sync();

  const linkRanges = body.getRange().getHyperlinkRanges();
  linkRanges.load("items");
  await context.sync();

  for (const r of linkRanges.items) {
    r.load("text,hyperlink");
  }
  await context.sync();

  return linkRanges.items.map(r => ({ text: r.text, url: r.hyperlink }));
});
```

## Remove a Hyperlink

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.hyperlink = "";
  await context.sync();
});
```
