# Tables — Creation, Rows, Columns, Cells

## Key Types
- `Word.Table` — rows, columns, values, style, getCell(), addRows(), addColumns(), getBorder()
- `Word.TableRow` — cells, font, horizontalAlignment, shadingColor
- `Word.TableCell` — body, value, columnWidth, shadingColor
- `Word.TableBorder` — type, color, width

## Create a Table

```javascript
await Word.run(async (context) => {
  const body = context.document.body;

  // Insert a 3x3 table with values (all values must be strings)
  const values = [
    ["STT", "Họ và tên", "Chức vụ"],
    ["1", "Nguyễn Văn A", "Trưởng phòng"],
    ["2", "Trần Thị B", "Chuyên viên"]
  ];

  const table = body.insertTable(values.length, values[0].length, Word.InsertLocation.end, values);

  // Style the header row
  table.getRow(0).font.bold = true;
  table.getRow(0).shadingColor = "#D9E2F3";

  // Set visible borders
  const border = table.getBorder(Word.BorderLocation.all);
  border.type = Word.BorderType.single;
  border.color = "#BFBFBF";
  border.width = 0.75;

  await context.sync();
});
```

## Read Table Data

```javascript
await Word.run(async (context) => {
  const tables = context.document.body.tables;
  tables.load("items");
  await context.sync();

  if (tables.items.length === 0) return [];

  const table = tables.items[0];
  table.load("rowCount,columnCount,values");
  await context.sync();

  return table.values; // 2D string array
});
```

## Modify Cells

```javascript
await Word.run(async (context) => {
  const tables = context.document.body.tables;
  tables.load("items");
  await context.sync();
  if (tables.items.length === 0) return;

  const table = tables.items[0];
  // Get a specific cell (0-indexed)
  const cell = table.getCell(1, 2); // row 1, col 2
  cell.body.clear();
  cell.body.insertText("Chuyên viên chính", Word.InsertLocation.start);
  cell.shadingColor = "#FFF2CC";

  await context.sync();
});
```

## Add Rows and Columns

```javascript
await Word.run(async (context) => {
  const tables = context.document.body.tables;
  tables.load("items");
  await context.sync();
  if (tables.items.length === 0) return;

  const table = tables.items[0];
  // Add a row at the end
  table.addRows(Word.InsertLocation.end, 1, [["3", "Lê Văn C", "Nhân viên"]]);

  await context.sync();
});
```

## Delete a Table

```javascript
await Word.run(async (context) => {
  const tables = context.document.body.tables;
  tables.load("items");
  await context.sync();
  if (tables.items.length > 0) {
    tables.items[0].delete();
    await context.sync();
  }
});
```

## Merge Cells

```javascript
await Word.run(async (context) => {
  const table = context.document.body.tables.getFirst();

  // mergeCells(topRow, firstCell, bottomRow, lastCell) — 0-indexed
  table.mergeCells(0, 0, 0, 2); // merge first 3 cells of row 0 into one
  await context.sync();
});
```

## Cell Padding

```javascript
await Word.run(async (context) => {
  const table = context.document.body.tables.getFirst();
  const cell = table.getCell(0, 0);

  // Padding in points; cellPaddingLocation: top | bottom | left | right
  cell.setCellPadding(Word.CellPaddingLocation.top,    4);
  cell.setCellPadding(Word.CellPaddingLocation.bottom, 4);
  cell.setCellPadding(Word.CellPaddingLocation.left,   8);
  cell.setCellPadding(Word.CellPaddingLocation.right,  8);
  await context.sync();
});
```

```javascript
// Read cell padding (getCellPadding returns ClientResult<number> — do NOT call .load()!)
await Word.run(async (context) => {
  const table = context.document.body.tables.getFirst();
  const cell = table.getCell(0, 0);

  const top = cell.getCellPadding(Word.CellPaddingLocation.top);
  await context.sync();

  return top.value;
});
```

## Auto-fit and Distribute Columns

```javascript
await Word.run(async (context) => {
  const table = context.document.body.tables.getFirst();
  table.autoFitWindow();     // fit table to page width
  table.distributeColumns(); // make all columns equal width
  await context.sync();
});
```

## Apply a Built-In Table Style

```javascript
await Word.run(async (context) => {
  const table = context.document.body.tables.getFirst();
  table.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
  await context.sync();
});
```

## Set Table Borders

```javascript
await Word.run(async (context) => {
  const table = context.document.body.tables.getFirst();

  // borderLocation: top | bottom | left | right | insideHorizontal | insideVertical | all
  const border = table.getBorder(Word.BorderLocation.all);
  border.type  = Word.BorderType.single;
  border.color = "#000000";
  border.width = 1.0; // points

  await context.sync();
});
```

## Common Pitfalls

- `insertTable` values is a 2D string array — all values must be strings
- Cell indices are 0-based
- Always `load("values")` before reading table data
- Prefer `table.styleBuiltIn = Word.BuiltInStyleName.tableGrid` over locale-dependent `table.style`
- `getCellPadding` returns `ClientResult<number>` — do NOT call `.load()`, just `await context.sync()` and read `.value`
