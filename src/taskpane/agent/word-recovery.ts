export const MAX_RECOVERY_BYTES = 20 * 1024 * 1024;
const TIMEOUT_MS = 30_000;

export interface WordRecoveryCopy {
  base64: string;
  bytes: number;
  createdAt: number;
  documentLabel: string;
}

function slice(file: Office.File, index: number): Promise<number[]> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Recovery slice timed out')), TIMEOUT_MS);
    try { file.getSliceAsync(index, result => {
      clearTimeout(timer);
      if (result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value.data)) resolve(result.value.data);
      else reject(new Error('Could not read the compressed Word document'));
    }); } catch (error) { clearTimeout(timer); reject(error); }
  });
}

/** The compressed DOCX includes headers, tables, images and formatting, not just body text. */
export async function captureWordRecovery(documentLabel: string): Promise<WordRecoveryCopy> {
  if (typeof Office === 'undefined' || !Office.context?.document?.getFileAsync) throw new Error('Document backup API unavailable');
  const file = await new Promise<Office.File>((resolve, reject) => {
    let expired = false;
    const timer = setTimeout(() => { expired = true; reject(new Error('Document backup timed out')); }, TIMEOUT_MS);
    try { Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 64 * 1024 }, result => {
      clearTimeout(timer);
      if (result.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error('Could not create a document backup')); return; }
      if (expired) { result.value.closeAsync(); return; }
      resolve(result.value);
    }); } catch (error) { clearTimeout(timer); reject(error); }
  });
  try {
    if (file.size <= 0 || file.size > MAX_RECOVERY_BYTES) throw new Error('Recovery document exceeds the 20 MB limit or is empty');
    const parts: string[] = [];
    let bytes = 0;
    for (let index = 0; index < file.sliceCount; index++) {
      const data = await slice(file, index);
      bytes += data.length;
      if (bytes > MAX_RECOVERY_BYTES || data.some(n => !Number.isInteger(n) || n < 0 || n > 255)) throw new Error('Invalid recovery document data');
      for (let start = 0; start < data.length; start += 8192) parts.push(String.fromCharCode(...data.slice(start, start + 8192)));
    }
    if (bytes !== file.size) throw new Error('Incomplete recovery document');
    return { base64: btoa(parts.join('')), bytes, createdAt: Date.now(), documentLabel };
  } finally { file.closeAsync(); }
}

/** Opens a separate copy, preserving all subsequent edits in the current document. */
export async function openWordRecovery(copy: WordRecoveryCopy): Promise<void> {
  if (typeof Word === 'undefined') throw new Error('Recovery must be opened in Word');
  await Word.run(async context => {
    context.application.createDocument(copy.base64).open();
    await context.sync();
  });
}
