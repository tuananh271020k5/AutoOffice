# Page Setup — Margins, Orientation, Paper Size

## Key Rules & Units
- All margin and dimension values are in **points**:
  - `1 inch = 72 points`
  - `1 cm = 28.35 points` (e.g., `2 cm = 56.7 pt`, `3 cm = 85.05 pt`, `1.5 cm = 42.52 pt`)
- Common paper sizes in points:
  - **A4**: width `595.3` pt (210 mm), height `841.9` pt (297 mm)
  - **US Letter**: width `612` pt, height `792` pt
- **IMPORTANT COMPATIBILITY RULE**: `section.pageSetup` requires `WordApiDesktop 1.3`. In some Word builds `section.pageSetup` may be `undefined`. Always check `if (section.pageSetup)` and fall back to updating `<w:pgSz>` / `<w:pgMar>` in `body.getOoxml()` when `section.pageSetup` is unavailable!

## Set A4 Paper Size & Standard Margins (with Universal Fallback)

```javascript
await Word.run(async (context) => {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  const topPt = 56.7;    // 2 cm
  const bottomPt = 56.7; // 2 cm
  const leftPt = 85.05;  // 3 cm
  const rightPt = 42.52; // 1.5 cm
  const widthPt = 595.3; // A4 width (210mm)
  const heightPt = 841.9;// A4 height (297mm)

  const firstSection = sections.items[0];
  if (firstSection && firstSection.pageSetup) {
    for (const sec of sections.items) {
      const ps = sec.pageSetup;
      ps.orientation = Word.PageOrientation.portrait;
      ps.pageWidth = widthPt;
      ps.pageHeight = heightPt;
      ps.topMargin = topPt;
      ps.bottomMargin = bottomPt;
      ps.leftMargin = leftPt;
      ps.rightMargin = rightPt;
    }
    await context.sync();
    return "Updated page setup via Section.pageSetup";
  }

  // Fallback for Word versions where section.pageSetup is undefined (uses OOXML twips: 1 pt = 20 twips)
  const body = context.document.body;
  const ooxmlResult = body.getOoxml();
  await context.sync();

  const wTwips = Math.round(widthPt * 20);
  const hTwips = Math.round(heightPt * 20);
  const topTwips = Math.round(topPt * 20);
  const bottomTwips = Math.round(bottomPt * 20);
  const leftTwips = Math.round(leftPt * 20);
  const rightTwips = Math.round(rightPt * 20);

  let xml = ooxmlResult.value;
  xml = xml.replace(
    /<w:pgSz\b[^>]*\/>/g,
    `<w:pgSz w:w="${wTwips}" w:h="${hTwips}" w:orient="portrait"/>`
  );
  xml = xml.replace(
    /<w:pgMar\b[^>]*\/>/g,
    `<w:pgMar w:top="${topTwips}" w:right="${rightTwips}" w:bottom="${bottomTwips}" w:left="${leftTwips}" w:header="720" w:footer="720" w:gutter="0"/>`
  );
  body.insertOoxml(xml, Word.InsertLocation.replace);
  await context.sync();
  return "Updated A4 paper size and margins via OOXML fallback";
});
```

## Read Current Page Setup (with Universal Fallback)

```javascript
await Word.run(async (context) => {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  const firstSection = sections.items[0];
  if (firstSection && firstSection.pageSetup) {
    const ps = firstSection.pageSetup;
    ps.load("topMargin,bottomMargin,leftMargin,rightMargin,orientation,pageWidth,pageHeight");
    await context.sync();
    return {
      topMarginCm: +(ps.topMargin / 28.35).toFixed(2),
      bottomMarginCm: +(ps.bottomMargin / 28.35).toFixed(2),
      leftMarginCm: +(ps.leftMargin / 28.35).toFixed(2),
      rightMarginCm: +(ps.rightMargin / 28.35).toFixed(2),
      orientation: ps.orientation,
      pageWidthPt: ps.pageWidth,
      pageHeightPt: ps.pageHeight,
    };
  }

  const ooxml = context.document.body.getOoxml();
  await context.sync();
  const pgSzMatch = ooxml.value.match(/<w:pgSz[^>]*w:w="(\d+)"[^>]*w:h="(\d+)"/);
  const pgMarMatch = ooxml.value.match(/<w:pgMar[^>]*w:top="(\d+)"[^>]*w:right="(\d+)"[^>]*w:bottom="(\d+)"[^>]*w:left="(\d+)"/);
  return {
    widthCm: pgSzMatch ? +(parseInt(pgSzMatch[1], 10) / 567).toFixed(1) : 21.0,
    heightCm: pgSzMatch ? +(parseInt(pgSzMatch[2], 10) / 567).toFixed(1) : 29.7,
    topMarginCm: pgMarMatch ? +(parseInt(pgMarMatch[1], 10) / 567).toFixed(2) : 2.0,
    rightMarginCm: pgMarMatch ? +(parseInt(pgMarMatch[2], 10) / 567).toFixed(2) : 1.5,
    bottomMarginCm: pgMarMatch ? +(parseInt(pgMarMatch[3], 10) / 567).toFixed(2) : 2.0,
    leftMarginCm: pgMarMatch ? +(parseInt(pgMarMatch[4], 10) / 567).toFixed(2) : 3.0,
  };
});
```
