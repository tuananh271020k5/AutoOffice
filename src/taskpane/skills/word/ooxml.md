# OOXML & HTML — Rich Content Boxes, Import & Export

## Key Rules
- **CRITICAL**: `getOoxml()` and `getHtml()` return `OfficeExtension.ClientResult<string>`. Do **NOT** call `.load()` on `ooxml` or `html`! Simply call `const ooxml = body.getOoxml(); await context.sync();` and read `ooxml.value`.
- **CRITICAL FOR STYLED BOXES / CALLOUTS**: Never hand-write partial `<w:p>` XML into `insertOoxml` without a complete `<pkg:package>` wrapper, or Word will throw `GeneralException`!
  - To create a styled notification box, callout box, or bordered frame, use **either** `body.insertTable(1, 1, ...)` with border/shading **or** the complete `<pkg:package>` template below!

## Insert a Styled Callout / Notification Box (Reliable Table + Border / OOXML)

```javascript
await Word.run(async (context) => {
  const body = context.document.body;

  // 1x1 styled callout box with double blue border and light blue background
  const box = body.insertTable(1, 1, Word.InsertLocation.end, [
    ["📌 THÔNG BÁO QUAN TRỌNG: Vui lòng kiểm tra kỹ nội dung và thời hạn trước khi trình ký."]
  ]);
  const cell = box.getCell(0, 0);
  cell.shadingColor = "#EBF3FC";
  cell.setCellPadding(Word.CellPaddingLocation.top, 8);
  cell.setCellPadding(Word.CellPaddingLocation.bottom, 8);
  cell.setCellPadding(Word.CellPaddingLocation.left, 12);
  cell.setCellPadding(Word.CellPaddingLocation.right, 12);

  const border = box.getBorder(Word.BorderLocation.all);
  border.type = Word.BorderType.double;
  border.color = "#0078D4";
  border.width = 1.5;

  const row = box.getRow(0);
  row.font.name = "Times New Roman";
  row.font.size = 13;
  row.font.bold = true;
  row.font.color = "#004578";

  await context.sync();
  return "Inserted styled notification box";
});
```

## Export Body as OOXML or HTML

```javascript
await Word.run(async (context) => {
  const body = context.document.body;
  const html = body.getHtml();
  await context.sync(); // NO html.load() needed!

  return html.value.slice(0, 500);
});
```

## Insert HTML Content

```javascript
await Word.run(async (context) => {
  const body = context.document.body;
  body.insertHtml(
    "<p style='font-family:Times New Roman;font-size:13pt;'><b>Ghi chú:</b> Nội dung được chèn từ HTML.</p>",
    Word.InsertLocation.end
  );
  await context.sync();
});
```
