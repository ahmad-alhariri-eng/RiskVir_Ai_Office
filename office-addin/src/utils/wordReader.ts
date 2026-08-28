export interface WordContext {
  bodyText: string;
  selectedText: string;
  wordCount: number;
  paragraphCount: number;
  title: string;
  author: string;
}

/** Returns rich context about the current Word document */
export async function getSelectedWordText(): Promise<WordContext> {
  return Word.run(async (context) => {
    const doc = context.document;
    const body = doc.body;
    const selection = doc.getSelection();

    // Load everything in parallel
    body.load(['text', 'paragraphs']);
    selection.load('text');
    doc.properties.load(['title', 'author', 'wordCount']);
    await context.sync();

    const bodyText = body.text;
    const selectedText = selection.text;
    const wordCount = (doc.properties as any).wordCount ?? bodyText.split(/\s+/).filter(Boolean).length;
    const paragraphCount = bodyText.split('\n').filter(t => t.trim().length > 0).length;

    return {
      bodyText: bodyText.length > 3000 ? bodyText.slice(0, 3000) + '…[truncated]' : bodyText,
      selectedText,
      wordCount,
      paragraphCount,
      title: doc.properties.title || '',
      author: doc.properties.author || '',
    };
  });
}
