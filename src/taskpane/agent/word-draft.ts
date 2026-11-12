export interface WordDraft {
  seal(): Promise<void>;
  accept(): void;
  reject(): Promise<void>;
}

export function supportsInlineDraft(tool: string, op: string): boolean {
  return tool === 'word_edit' || tool === 'word_format' ||
    (tool === 'word_structure' && ['create_table', 'format_table', 'insert_list', 'insert_callout_box', 'insert_break'].includes(op));
}

async function readBodyXml(): Promise<string> {
  return Word.run(async context => {
    const xml = context.document.body.getOoxml();
    await context.sync();
    return xml.value;
  });
}

/** Ignore Word's generated edit IDs and package timestamps, not content or formatting. */
export function canonicalBodyXml(xml: string): string {
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  if (document.getElementsByTagName('parsererror').length) throw new Error('Word returned invalid document XML.');
  const goBackIds = new Set<string>();
  const wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  for (const mark of Array.from(document.getElementsByTagNameNS(wordNs, 'bookmarkStart'))) {
    if (mark.getAttributeNS(wordNs, 'name') === '_GoBack') {
      goBackIds.add(mark.getAttributeNS(wordNs, 'id') || '');
      mark.remove();
    }
  }
  for (const mark of Array.from(document.getElementsByTagNameNS(wordNs, 'bookmarkEnd'))) {
    if (goBackIds.has(mark.getAttributeNS(wordNs, 'id') || '')) mark.remove();
  }
  for (const element of Array.from(document.getElementsByTagName('*'))) {
    for (const attribute of Array.from(element.attributes)) {
      if (/^rsid/.test(attribute.localName) || ['paraId', 'textId'].includes(attribute.localName)) element.removeAttributeNode(attribute);
    }
    if (element.namespaceURI === wordNs && ['proofErr', 'lastRenderedPageBreak', 'rsids'].includes(element.localName)) element.remove();
    if (element.localName === 'part') {
      const name = element.getAttributeNS('http://schemas.microsoft.com/office/2006/xmlPackage', 'name') || '';
      // Word refreshes counters, pagination and settings during idle/AutoSave. These
      // are not body edits and must not make a pending draft impossible to reject.
      if (name.startsWith('/docProps/') || ['/word/settings.xml', '/word/webSettings.xml'].includes(name)) element.remove();
    }
  }
  return new XMLSerializer().serializeToString(document);
}

async function fingerprint(xml: string): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(canonicalBodyXml(xml)));
  return Array.from(new Uint8Array(digest), n => n.toString(16).padStart(2, '0')).join('');
}

/** Keep the edit visible in the current document until Accept or Reject. */
export async function beginWordDraft(): Promise<WordDraft> {
  const before = await readBodyXml();
  if (before.length > 8 * 1024 * 1024) throw new Error('This document is too large for inline review (8 MB XML limit).');
  const originalXml = new DOMParser().parseFromString(before, 'application/xml');
  const wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  if (['ins', 'del', 'rPrChange', 'pPrChange', 'tblPrChange'].some(name => originalXml.getElementsByTagNameNS(wordNs, name).length > 0)) {
    throw new Error('Review existing Word revisions before using inline Accept / Reject. They will not be accepted or rejected automatically.');
  }
  await Word.run(async context => {
    let body = context.document.getSelection().parentBody;
    for (let depth = 0; depth < 8; depth++) {
      body.load('type');
      await context.sync();
      if (body.type === 'MainDoc' || body.type === 'Section') break;
      if (body.type !== 'TableCell' || depth === 7) throw new Error('Inline review supports the main document body. Select text in the body before editing.');
      body = body.parentBody;
    }
    context.document.load('changeTrackingMode');
    await context.sync();
    if (context.document.changeTrackingMode !== Word.ChangeTrackingMode.off) {
      throw new Error('Turn off Word Track Changes before using inline Accept / Reject. Existing revisions must be reviewed separately.');
    }
  });
  let after = '';
  let done = false;
  return {
    async seal() { after = await fingerprint(await readBodyXml()); },
    accept() { done = true; },
    async reject() {
      if (done) throw new Error('This edit was already reviewed.');
      if (!after || await fingerprint(await readBodyXml()) !== after) {
        throw new Error('The document changed after the AI edit. Reject was stopped to preserve your newer edits. Use Word Undo or review the document manually.');
      }
      await Word.run(async context => {
        context.document.load('changeTrackingMode');
        await context.sync();
        if (context.document.changeTrackingMode !== Word.ChangeTrackingMode.off) throw new Error('Track Changes was enabled while reviewing. Reject was stopped.');
        context.document.body.insertOoxml(before, Word.InsertLocation.replace);
        await context.sync();
      });
      done = true;
    },
  };
}
