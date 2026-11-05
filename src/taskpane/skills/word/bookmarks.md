# Bookmarks — Named Ranges & Navigation

## Key Rules
- Bookmark names must ONLY contain letters, numbers, and underscores (`_`), must start with a letter, and cannot contain spaces (e.g., `"MocQuanTrong_1"`).
- **CRITICAL**:
  - `insertBookmark(name)` is on **`Word.Range`** (e.g. `range.insertBookmark("MocQuanTrong_1")`).
  - `getBookmarks()` is on **`Word.Range`** (`context.document.body.getRange().getBookmarks(false, false)`), NOT on `Word.Body`!
  - `getBookmarks()` returns `OfficeExtension.ClientResult<string[]>` — do **NOT** call `.load("items")` on it! Just `await context.sync()` and read `bookmarks.value`.

## Insert a Bookmark & List All Bookmarks

```javascript
await Word.run(async (context) => {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items,text");
  await context.sync();

  // Bookmark the first non-empty paragraph (or create one if empty)
  let targetRange;
  const firstNonEmpty = paragraphs.items.find(p => p.text && p.text.trim().length > 0);
  if (firstNonEmpty) {
    targetRange = firstNonEmpty.getRange();
  } else {
    const p = body.insertParagraph("Đoạn văn được đánh dấu Bookmark", Word.InsertLocation.end);
    targetRange = p.getRange();
  }

  targetRange.insertBookmark("ViTriQuanTrong");
  const bookmarksResult = body.getRange().getBookmarks(false, false);
  await context.sync();

  return {
    addedBookmark: "ViTriQuanTrong",
    allBookmarks: bookmarksResult.value,
  };
});
```

## Navigate to or Replace Text at a Bookmark

```javascript
await Word.run(async (context) => {
  const range = context.document.getBookmarkRangeOrNullObject("ViTriQuanTrong");
  range.load("isNullObject,text");
  await context.sync();

  if (range.isNullObject) {
    return "Bookmark not found";
  }
  range.select();
  await context.sync();
  return range.text;
});
```
