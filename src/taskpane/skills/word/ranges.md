# Ranges — Selection, Manipulation, InsertLocation

## Key Types
- `Word.Range` — `text`, `font`, `paragraphs`, `hyperlink`, `insertText()`, `insertHtml()`, `insertBreak()`, `delete()`, `expand()`, `compareLocationWith()`
- `Word.InsertLocation` — `before`, `after`, `start`, `end`, `replace`
- `Word.RangeLocation` — `whole`, `start`, `end`, `before`, `after`, `content`
- `Word.LocationRelation` — `equal`, `before`, `after`, `contains`, `inside`, `adjacentBefore`, `adjacentAfter`, `overlappingBefore`, `overlappingAfter`, `disconnected`

## Read Selection or Document Text (Safe Fallback When Nothing Selected)

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  const body = context.document.body;
  selection.load("text");
  body.load("text");
  await context.sync();

  const selectedText = (selection.text || "").trim();
  return {
    hasSelection: selectedText.length > 0,
    text: selectedText.length > 0 ? selectedText : (body.text || "").trim(),
  };
});
```

## Insert Text at Location

```javascript
await Word.run(async (context) => {
  const body = context.document.body;

  // Insert at end of document
  body.insertText("Nội dung mới", Word.InsertLocation.end);

  // Insert line break
  body.insertBreak(Word.BreakType.line, Word.InsertLocation.end);

  await context.sync();
});
```

## Insert HTML

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.insertHtml(
    "<b>Chữ in đậm</b> và <i>chữ in nghiêng</i>",
    Word.InsertLocation.replace
  );
  await context.sync();
});
```

## Expand a Range to a Whole Paragraph

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();

  // Expand to cover the full paragraph containing the selection
  const para = selection.paragraphs.getFirst();
  const paraRange = para.getRange();
  paraRange.load("text");
  await context.sync();

  return paraRange.text;
});
```

## Compare Range Positions

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  const body = context.document.body;
  const bodyRange = body.getRange(Word.RangeLocation.whole);

  // compareLocationWith returns ClientResult<LocationRelation> — do NOT call .load()!
  const relation = selection.compareLocationWith(bodyRange);
  await context.sync();

  return relation.value;
});
```

## Programmatically Select / Highlight Text

```javascript
await Word.run(async (context) => {
  const results = context.document.body.search("ĐƠN XIN NGHỈ PHÉP");
  results.load("items");
  await context.sync();

  if (results.items.length > 0) {
    results.items[0].select(); // highlight/select first match
    await context.sync();
  }
});
```

## Common Pitfalls

- A `Range` from `getSelection()` is empty (`selection.text === ""`) when the user hasn't highlighted anything — always fall back to `context.document.body.text` when reading text for summarization/analysis!
- **NEVER** overwrite `selection` when the user only asks to **summarize**, **analyze**, or **count** text — return the text from `execute_code` and answer in chat!
- `compareLocationWith` returns `ClientResult<LocationRelation>` — do NOT call `.load()`, just `await context.sync()` and read `.value`
