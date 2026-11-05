# Content Controls — Rich Text, Plain Text, Dropdowns, Checkboxes

## Key Types
- `Word.ContentControl` — `tag`, `title`, `type` (read-only), `appearance`, `color`, `placeholderText`, `cannotDelete`, `cannotEdit`, `removeWhenEdited`
- Types (passed to `range.insertContentControl(type)`): `Word.ContentControlType.richText`, `plainText`, `dropDownList`, `checkBox`, `comboBox`

## Insert a Content Control (Safe Pattern — Works With or Without Selection)

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();

  // If the user has selected text, wrap the selection; otherwise insert a new paragraph at the end
  let targetRange = selection;
  if (!selection.text || selection.text.trim().length === 0) {
    const para = context.document.body.insertParagraph("Nhập họ và tên tại đây", Word.InsertLocation.end);
    targetRange = para.getRange();
  }

  const cc = targetRange.insertContentControl();
  cc.title = "Họ và tên";
  cc.tag = "full_name";
  cc.appearance = Word.ContentControlAppearance.boundingBox;
  cc.color = "#0078D4";
  cc.placeholderText = "Nhập họ và tên...";

  await context.sync();
});
```

## Find Content Controls by Tag

```javascript
await Word.run(async (context) => {
  const controls = context.document.contentControls;
  controls.load("items");
  await context.sync();

  for (const cc of controls.items) {
    cc.load("tag,title,text");
  }
  await context.sync();

  return controls.items.map(cc => ({
    tag: cc.tag,
    title: cc.title,
    text: cc.text,
  }));
});
```

## Set Content Control Value

```javascript
await Word.run(async (context) => {
  const controls = context.document.contentControls.getByTag("full_name");
  controls.load("items");
  await context.sync();

  if (controls.items.length > 0) {
    controls.items[0].insertText("Nguyễn Văn A", Word.InsertLocation.replace);
    await context.sync();
  }
});
```

## Checkbox Content Control

```javascript
// Insert a checkbox (pass type to insertContentControl — cc.type is read-only!)
await Word.run(async (context) => {
  const para = context.document.body.insertParagraph(" Đồng ý điều khoản", Word.InsertLocation.end);
  const startRange = para.getRange(Word.RangeLocation.start);
  const cc = startRange.insertContentControl(Word.ContentControlType.checkBox);
  cc.tag = "agree_terms";
  cc.title = "Agree to Terms";
  cc.checkboxContentControl.isChecked = true;
  await context.sync();
});
```

```javascript
// Read and toggle checkbox state
await Word.run(async (context) => {
  const controls = context.document.contentControls.getByTag("agree_terms");
  controls.load("items");
  await context.sync();

  if (controls.items.length > 0) {
    const cc = controls.items[0];
    cc.checkboxContentControl.load("isChecked");
    await context.sync();

    cc.checkboxContentControl.isChecked = !cc.checkboxContentControl.isChecked;
    await context.sync();
  }
});
```

## Dropdown List Content Control

```javascript
await Word.run(async (context) => {
  const para = context.document.body.insertParagraph("", Word.InsertLocation.end);
  const cc = para.getRange().insertContentControl(Word.ContentControlType.dropDownList);
  cc.tag = "status";
  cc.title = "Trạng thái";

  // Add choices via dropDownListContentControl.addListItem(displayText, value)
  cc.dropDownListContentControl.addListItem("Bản nháp", "draft");
  cc.dropDownListContentControl.addListItem("Đang duyệt", "review");
  cc.dropDownListContentControl.addListItem("Đã phê duyệt", "approved");

  await context.sync();
});
```

## Protect a Content Control

```javascript
await Word.run(async (context) => {
  const controls = context.document.contentControls.getByTag("full_name");
  controls.load("items");
  await context.sync();

  if (controls.items.length > 0) {
    const cc = controls.items[0];
    cc.cannotDelete = true;  // user cannot delete the CC
    cc.cannotEdit   = true;  // user cannot edit the CC's content
    await context.sync();
  }
});
```

## Temporary Content Control (Remove When Edited)

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  const cc = selection.insertContentControl();
  cc.removeWhenEdited = true; // CC is removed when user edits inside it (mutually exclusive with cannotDelete)
  await context.sync();
});
```

## Clear or Delete a Content Control

```javascript
await Word.run(async (context) => {
  const controls = context.document.contentControls.getByTag("full_name");
  controls.load("items");
  await context.sync();

  if (controls.items.length > 0) {
    // keepContent: true preserves the inner text; false removes it too
    controls.items[0].delete(true);
    await context.sync();
  }
});
```

## Common Pitfalls

- `cc.type` is **read-only**! Pass `Word.ContentControlType.*` as an argument to `range.insertContentControl(type)`
- Use `cc.removeWhenEdited = true` (NOT `cc.temporary = true`)
- Use `cc.checkboxContentControl.isChecked` (NOT `cc.getCheckedState()`)
- Use `cc.dropDownListContentControl.addListItem(text, value)` (NOT `cc.dropdownListEntries.add`)
- `cannotEdit = true` also blocks API writes to that control's content — set it only after populating
