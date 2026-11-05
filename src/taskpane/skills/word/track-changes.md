# Track Changes — TrackedChange API

## Key Types
- `Word.TrackedChange` — `author`, `date`, `type`, `text`, `getRange()`, `accept()`, `reject()`
- `Word.TrackedChangeCollection` — `items` array; `acceptAll()` / `rejectAll()`
- `Word.Document.changeTrackingMode` — enable/disable tracking
- `Word.ChangeTrackingMode` — `off` | `trackAll` | `trackMineOnly`
- `Word.TrackedChangeType` — `added` (`"Added"`) | `deleted` (`"Deleted"`) | `formatted` (`"Formatted"`) | `none` (`"None"`)

## Enable / Disable Track Changes

```javascript
await Word.run(async (context) => {
  // Turn tracking on for all authors
  context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  await context.sync();
});
```

```javascript
await Word.run(async (context) => {
  context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
  await context.sync();
});
```

## Read All Tracked Changes

```javascript
await Word.run(async (context) => {
  const changes = context.document.body.getTrackedChanges();
  changes.load("items");
  await context.sync();

  for (const change of changes.items) {
    change.load("author,date,type,text");
  }
  await context.sync();

  return changes.items.map(c => ({
    author: c.author,
    date: c.date,
    type: c.type,
    text: c.text,
  }));
});
```

## Accept All Tracked Changes

```javascript
await Word.run(async (context) => {
  const changes = context.document.body.getTrackedChanges();
  changes.load("items");
  await context.sync();

  changes.acceptAll();
  await context.sync();
});
```

## Reject All Tracked Changes

```javascript
await Word.run(async (context) => {
  const changes = context.document.body.getTrackedChanges();
  changes.load("items");
  await context.sync();

  changes.rejectAll();
  await context.sync();
});
```

## Accept or Reject Individual Changes

```javascript
await Word.run(async (context) => {
  const changes = context.document.body.getTrackedChanges();
  changes.load("items");
  await context.sync();

  for (const change of changes.items) {
    change.load("author,type,text");
  }
  await context.sync();

  for (const change of changes.items) {
    if (change.author === "Alice") {
      change.accept();
    } else {
      change.reject();
    }
  }
  await context.sync();
});
```

## Check Current Tracking Mode

```javascript
await Word.run(async (context) => {
  context.document.load("changeTrackingMode");
  await context.sync();

  return context.document.changeTrackingMode;
  // "Off" | "TrackAll" | "TrackMineOnly"
});
```

## Paragraph-Scoped Tracked Changes

```javascript
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();

  if (paragraphs.items.length === 0) return [];

  const changes = paragraphs.items[0].getTrackedChanges();
  changes.load("items");
  await context.sync();

  for (const change of changes.items) {
    change.load("author,type,text");
  }
  await context.sync();

  return changes.items.map(c => ({ author: c.author, type: c.type, text: c.text }));
});
```

## Common Pitfalls

- `TrackedChange.text` is a direct string property on `TrackedChange` — load `"author,date,type,text"` directly on `change`
- To get the `Word.Range` of a tracked change, call `change.getRange()` (method), NOT `change.range`
- Always load `changeTrackingMode` before reading it
- `changes.acceptAll()` / `changes.rejectAll()` operate on the collection
- Requires Word JS API 1.6 or later
