# Context & Sync — The office.js Batching Model

## Core Concept

office.js uses a **proxy-based batching model**. Objects you get from the API are local proxies. Property reads and method calls are queued as commands. Nothing actually executes until you call `context.sync()`.

## Critical Rules

1. **load() before read**: You must call `.load("propertyName")` on a proxy object before you can read its value.
2. **sync() before access**: After `.load()`, you must `await context.sync()` before the property has its real value.
3. **Batch writes, then sync**: Multiple writes can be batched before a single `context.sync()`.
4. **Return data from `Word.run`**: Anything returned inside `await Word.run(async (context) => { ... return data; })` is returned as the result of `execute_code`.

## Correct Pattern

```javascript
return await Word.run(async (context) => {
  const body = context.document.body;
  const paragraphs = body.paragraphs;

  // Queue loading
  paragraphs.load("items");

  // Execute the queued commands
  await context.sync();

  for (const para of paragraphs.items) {
    para.load("text,style");
  }
  await context.sync();

  // NOW you can read text/style and return a summary
  return paragraphs.items.map(p => ({ text: p.text, style: p.style }));
});
```

## Common Mistakes

### Reading before sync
```javascript
// WRONG
const body = context.document.body;
body.load("text");
console.log(body.text); // undefined! sync() not called yet

// RIGHT
const body = context.document.body;
body.load("text");
await context.sync();
console.log(body.text); // now has the value
```

### Forgetting to load
```javascript
// WRONG
const para = context.document.body.paragraphs.getFirst();
await context.sync();
console.log(para.text); // PropertyNotLoaded error!

// RIGHT
const para = context.document.body.paragraphs.getFirst();
para.load("text");
await context.sync();
console.log(para.text);
```

## InsertLocation Enum

When inserting content, use `Word.InsertLocation`:
- `Word.InsertLocation.start` — At the start of the target's content (valid on `body`, `paragraph`, `range`)
- `Word.InsertLocation.end` — At the end of the target's content (valid on `body`, `paragraph`, `range`)
- `Word.InsertLocation.replace` — Replaces the target's content (valid on `body`, `paragraph`, `range`)
- `Word.InsertLocation.before` — Before the target (**ONLY** valid on `Paragraph` or `Range`, **NEVER** on `context.document.body`!)
- `Word.InsertLocation.after` — After the target (**ONLY** valid on `Paragraph` or `Range`, **NEVER** on `context.document.body`!)

## ClientResult<T> Objects (Do NOT call `.load()`)

Methods that return primitive values or arrays wrapped in `OfficeExtension.ClientResult<T>` — such as `getOoxml()`, `getHtml()`, `getBookmarks()`, `getCellPadding()`, `compareLocationWith()`, `getBase64ImageSrc()` — do **NOT** have a `.load()` method!
Simply call the method, then `await context.sync()`, and read `result.value`.
