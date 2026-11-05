# Fields — Dynamic Content (Table of Contents, Page Numbers, Date)

## Key Rules
- **CRITICAL**: `insertField` is a method on **`Word.Range`**, NOT on `Word.Body`!
  - Always get a Range first: `const range = context.document.body.getRange(Word.RangeLocation.start);` (or `.end`), then call `range.insertField(...)`.
- Signature: `range.insertField(insertLocation, fieldType, text, removeFormatting)`

## Insert a Table of Contents (TOC) at Start of Document

```javascript
await Word.run(async (context) => {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items,text,styleBuiltIn");
  await context.sync();

  // Ensure headings exist; if none are styled as Heading 1/2, promote short section titles
  let headingCount = 0;
  for (const p of paragraphs.items) {
    if (
      p.styleBuiltIn === Word.BuiltInStyleName.heading1 ||
      p.styleBuiltIn === Word.BuiltInStyleName.heading2
    ) {
      headingCount++;
    }
  }

  const startRange = body.getRange(Word.RangeLocation.start);
  const tocTitle = body.insertParagraph("MỤC LỤC", Word.InsertLocation.start);
  tocTitle.font.name = "Times New Roman";
  tocTitle.font.size = 14;
  tocTitle.font.bold = true;
  tocTitle.alignment = Word.Alignment.centered;

  const tocRange = tocTitle.getRange(Word.RangeLocation.after);
  tocRange.insertField(Word.InsertLocation.after, Word.FieldType.toc, '\\o "1-3" \\h \\z \\u', true);

  await context.sync();
  return { insertedToc: true, headingCount };
});
```

## Insert a Date or Page Number Field

```javascript
await Word.run(async (context) => {
  const endRange = context.document.body.getRange(Word.RangeLocation.end);
  endRange.insertField(Word.InsertLocation.end, Word.FieldType.date, 'DATE \\@ "dd/MM/yyyy"', true);
  await context.sync();
});
```

## Read All Fields in the Document

```javascript
await Word.run(async (context) => {
  const fields = context.document.body.fields;
  fields.load("items");
  await context.sync();

  for (const field of fields.items) {
    field.load("code,type");
    field.result.load("text");
  }
  await context.sync();

  return fields.items.map(f => ({
    type: f.type,
    code: f.code,
    result: f.result.text,
  }));
});
```

## Common Pitfalls
- **NEVER call `body.insertField(...)` directly** — always call `body.getRange(Word.RangeLocation.start).insertField(...)` or `body.getRange(Word.RangeLocation.end).insertField(...)`.
- `Field.result` is a `Word.Range` — load its `text` property before reading.
