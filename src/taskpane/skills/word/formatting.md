# Formatting — Font, Color, Paragraph

## Key Types
- `Word.Font` — bold, italic, color, size, name, underline, highlightColor, subscript, superscript, strikeThrough, doubleStrikeThrough, hidden
- `Word.Paragraph` — alignment, lineSpacing, spaceAfter, spaceBefore, firstLineIndent, leftIndent, rightIndent, styleBuiltIn

## CRITICAL RULE FOR `lineSpacing`
- In Word `office.js`, **`para.lineSpacing` is measured in POINTS (1 line = 12 points)**, NOT line multipliers!
  - `1.0` single spacing = `12`
  - `1.15` spacing = `13.8`
  - `1.3` spacing = `15.6`
  - `1.5` spacing = `18`
  - `2.0` double spacing = `24`
- **NEVER** set `para.lineSpacing = 1.3` or `1.5` directly (that sets line height to 1.5pt and squishes lines on top of each other!). Always multiply line multiples by `12` (e.g. `1.3 * 12 = 15.6`).

## Font Formatting

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();

  const target = selection.text && selection.text.trim()
    ? selection
    : context.document.body;

  target.font.name = "Times New Roman";
  target.font.size = 13;
  target.font.color = "#000000";
  await context.sync();
});
```

## Paragraph Formatting (Standard Administrative Style)

```javascript
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items,text,alignment");
  await context.sync();

  for (const para of paragraphs.items) {
    para.font.name = "Times New Roman";
    para.font.size = 13;
    para.lineSpacing = 15.6; // 1.3 lines (1.3 * 12 pt)
    para.spaceAfter = 6;     // 6 points after paragraph
    para.spaceBefore = 0;
    // Keep centered headers centered; justify body paragraphs
    if (para.alignment !== Word.Alignment.centered && para.alignment !== "Centered") {
      para.alignment = Word.Alignment.justified;
    }
  }

  await context.sync();
});
```

## Format Specific Paragraphs by Style

```javascript
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items,style,styleBuiltIn");
  await context.sync();

  for (const para of paragraphs.items) {
    if (para.styleBuiltIn === Word.BuiltInStyleName.heading1 || para.style === "Heading 1") {
      para.font.color = "#1F4E79";
      para.font.size = 16;
      para.font.bold = true;
    }
  }
  await context.sync();
});
```

## Paragraph Flow Control & Indentation

```javascript
await Word.run(async (context) => {
  const para = context.document.body.paragraphs.getFirst();
  para.leftIndent = 0;
  para.rightIndent = 0;
  para.firstLineIndent = 28.35; // 1 cm first-line indent (1 cm ≈ 28.35 pt)
  para.keepWithNext = true;
  await context.sync();
});
```

## Format Standard Administrative Document (Quốc hiệu, Tiêu ngữ, Tiêu đề Centered + Body Justified)

```javascript
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  for (const para of paragraphs.items) {
    para.load("text,alignment");
  }
  await context.sync();

  for (const para of paragraphs.items) {
    const text = (para.text || "").trim();
    const upper = text.toUpperCase();

    para.font.name = "Times New Roman";
    para.font.size = 13;
    para.lineSpacing = 15.6; // 1.3 lines in points (12pt = 1 line)
    para.spaceAfter = 6;

    // 1. Explicitly center & bold Quốc hiệu, Tiêu ngữ, and Document Title (NEVER rely on paragraph index!)
    const isNationalHeaderOrTitle =
      upper.includes("CỘNG HÒA XÃ HỘI") ||
      text.includes("Độc lập - Tự do - Hạnh phúc") ||
      /^(ĐƠN |BÁO CÁO|QUYẾT ĐỊNH|BIÊN BẢN|TỜ TRÌNH|HỢP ĐỒNG|THÔNG BÁO|KẾ HOẠCH|GIẤY )/.test(upper);

    const isSignatureOrDate =
      /^(NGÀY\s|\.\.\.,\s*NGÀY|HÀ NỘI|TP\.|NGƯỜI LÀM ĐƠN|NGƯỜI VIẾT ĐƠN|KÝ VÀ GHI RÕ|\(KÝ)/.test(upper);

    if (isNationalHeaderOrTitle) {
      para.alignment = Word.Alignment.centered;
      para.font.bold = true;
      if (upper.startsWith("ĐƠN ") || upper.startsWith("BÁO CÁO") || upper.startsWith("QUYẾT ĐỊNH")) {
        para.font.size = 15;
        para.spaceBefore = 12;
        para.spaceAfter = 12;
      }
    } else if (isSignatureOrDate) {
      para.alignment = Word.Alignment.right;
    } else if (text.length > 0) {
      para.alignment = Word.Alignment.justified;
    }
  }

  await context.sync();
});
```

## Common Pitfalls
- **Line spacing unit**: Always use points (`12` = 1.0 line, `15.6` = 1.3 lines, `18` = 1.5 lines).
- **Alignment enum**: Use `Word.Alignment.left` (`"Left"`), `Word.Alignment.centered` (`"Centered"`), `Word.Alignment.right` (`"Right"`), `Word.Alignment.justified` (`"Justified"`).
- **Font color**: Accepts hex strings (`"#1F4E79"`) or named colors (`"red"`).
- **Highlight color**: Uses `Word.HighlightColor.yellow`, etc.
