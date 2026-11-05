# Footnotes & Endnotes — NoteItem API

## Key Types
- `Word.NoteItem` — `body`, `reference`, `type`
- `Word.NoteItemCollection` — `items` array
- `Word.Body.footnotes` / `Word.Body.endnotes` — collections on the document body
- `Word.NoteItemType` — `footnote` | `endnote`

## Insert a Footnote (Safe Pattern — Works With or Without Selection)

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("text");
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  let targetRange = selection;
  if (!selection.text || selection.text.trim().length === 0) {
    for (const p of paragraphs.items) {
      p.load("text");
    }
    await context.sync();
    const firstNonEmpty = paragraphs.items.find(p => p.text && p.text.trim().length > 0);
    if (firstNonEmpty) {
      targetRange = firstNonEmpty.getRange(Word.RangeLocation.end);
    } else {
      const p = context.document.body.insertParagraph("Văn bản tham chiếu", Word.InsertLocation.end);
      targetRange = p.getRange(Word.RangeLocation.end);
    }
  }

  const footnote = targetRange.insertFootnote("Căn cứ theo Bộ luật Lao động năm 2019.");
  await context.sync();
});
```

## Insert an Endnote

```javascript
await Word.run(async (context) => {
  const endRange = context.document.body.getRange(Word.RangeLocation.end);
  const endnote = endRange.insertEndnote("Ghi chú cuối tài liệu.");
  await context.sync();
});
```

## Read All Footnotes

```javascript
await Word.run(async (context) => {
  const footnotes = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  for (const note of footnotes.items) {
    note.body.load("text");
    note.reference.load("text"); // the superscript reference mark
  }
  await context.sync();

  return footnotes.items.map(n => ({
    reference: n.reference.text,
    text: n.body.text,
  }));
});
```

## Read All Endnotes

```javascript
await Word.run(async (context) => {
  const endnotes = context.document.body.endnotes;
  endnotes.load("items");
  await context.sync();

  for (const note of endnotes.items) {
    note.body.load("text");
  }
  await context.sync();

  return endnotes.items.map(n => n.body.text);
});
```

## Modify Footnote Content

```javascript
await Word.run(async (context) => {
  const footnotes = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  if (footnotes.items.length > 0) {
    const first = footnotes.items[0];
    first.body.clear();
    first.body.insertText("Nội dung chú thích đã cập nhật.", Word.InsertLocation.start);
    await context.sync();
  }
});
```

## Delete a Footnote

```javascript
await Word.run(async (context) => {
  const footnotes = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  if (footnotes.items.length > 0) {
    footnotes.items[0].delete();
    await context.sync();
  }
});
```

## Common Pitfalls

- `insertFootnote` / `insertEndnote` are on `Word.Range` (NOT on `Word.Body` directly) — call them on `selection`, `paragraph.getRange(Word.RangeLocation.end)`, or `body.getRange(Word.RangeLocation.end)`
- `NoteItem.body` is a full `Word.Body`; use standard body APIs (`insertText`, `insertParagraph`, etc.) to edit it
- `NoteItem.reference` is the superscript number/symbol in the main text — it is read-only
- Requires Word JS API 1.5 or later
