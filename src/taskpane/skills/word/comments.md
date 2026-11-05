# Comments — Add, Read, Reply

## Key Types
- `Word.Comment` — `id`, `content`, `authorName`, `createdDate`, `replies`, `resolved`
- `Word.CommentReply` — `content`, `authorName`, `createdDate`

## Add a Comment (Safe Pattern — Works With or Without Selection)

```javascript
await Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("text");
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  // If selection is non-empty, comment on selection; otherwise comment on the first non-empty paragraph
  let targetRange = selection;
  if (!selection.text || selection.text.trim().length === 0) {
    for (const p of paragraphs.items) {
      p.load("text");
    }
    await context.sync();
    const firstNonEmpty = paragraphs.items.find(p => p.text && p.text.trim().length > 0);
    if (firstNonEmpty) {
      targetRange = firstNonEmpty.getRange();
    } else {
      const p = context.document.body.insertParagraph("Nội dung cần rà soát", Word.InsertLocation.end);
      targetRange = p.getRange();
    }
  }

  const comment = targetRange.insertComment("Cần kiểm tra và xác nhận lại thông tin ở mục này.");
  comment.load("id,content");
  await context.sync();
  return { id: comment.id, content: comment.content };
});
```

## Read All Comments

```javascript
await Word.run(async (context) => {
  const comments = context.document.body.getComments();
  comments.load("items");
  await context.sync();

  for (const comment of comments.items) {
    comment.load("id,content,authorName,createdDate");
  }
  await context.sync();

  return comments.items.map(c => ({
    id: c.id,
    author: c.authorName,
    content: c.content,
    date: c.createdDate,
  }));
});
```

## Reply to a Comment

```javascript
await Word.run(async (context) => {
  const comments = context.document.body.getComments();
  comments.load("items");
  await context.sync();

  if (comments.items.length > 0) {
    const reply = comments.items[0].reply("Đã kiểm tra và cập nhật.");
    reply.load("content");
    await context.sync();
  }
});
```

## Delete a Comment

```javascript
await Word.run(async (context) => {
  const comments = context.document.body.getComments();
  comments.load("items");
  await context.sync();

  if (comments.items.length > 0) {
    comments.items[0].delete();
    await context.sync();
  }
});
```

## Common Pitfalls

- Comments API requires WordApi 1.4+
- `insertComment()` is on `Word.Range` (such as `selection` or `paragraph.getRange()`), NOT on `context.document`
- `getComments()` is on `context.document.body` or a `Word.Range`
- If the user has not selected any text, attach the comment to `paragraphs.items[0].getRange()` so the comment has a clear anchor text
