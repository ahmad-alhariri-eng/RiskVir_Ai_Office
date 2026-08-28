import type { WordAction } from '../types/actions';

export async function executeWordAction(action: WordAction) {
  try {
    await Word.run(async (context) => {
      const doc = context.document;

      switch (action.command) {

        // ── Insert Text ──────────────────────────────────────────
        case 'insertText': {
          if (!action.text) break;
          const selection = doc.getSelection();
          selection.insertText(action.text, Word.InsertLocation.after);
          break;
        }

        // ── Insert HTML ──────────────────────────────────────────
        case 'insertHtml': {
          if (!action.html) break;
          const selection = doc.getSelection();
          selection.insertHtml(action.html, Word.InsertLocation.after);
          break;
        }

        // ── Replace Selection ────────────────────────────────────
        case 'replaceSelection': {
          if (!action.text) break;
          const selection = doc.getSelection();
          selection.insertText(action.text, Word.InsertLocation.replace);
          break;
        }

        // ── Apply Style ──────────────────────────────────────────
        case 'applyStyle': {
          if (!action.styleName) break;
          if (action.targetText) {
            // Search and apply style to matching text
            const results = doc.body.search(action.targetText, { matchCase: false });
            results.load('items');
            await context.sync();
            results.items.forEach((item) => {
              item.style = action.styleName;
            });
          } else {
            // Apply to current selection
            const selection = doc.getSelection();
            selection.style = action.styleName;
          }
          break;
        }

        // ── Insert Page Break ────────────────────────────────────
        case 'insertPageBreak': {
          const selection = doc.getSelection();
          selection.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
          break;
        }

        // ── Insert Table of Contents ─────────────────────────────
        case 'insertTOC': {
          // Word JS API doesn't expose a direct TOC insert yet;
          // insert a field code as the closest equivalent.
          const levels = action.levels ?? 3;
          const tocHtml = `<p>[Table of Contents — Headings 1–${levels}]<br/>
            <em>To generate: References → Table of Contents in Word.</em></p>`;
          const selection = doc.getSelection();
          selection.insertHtml(tocHtml, Word.InsertLocation.before);
          break;
        }

        // ── Find and Replace ─────────────────────────────────────
        case 'findAndReplace': {
          if (!action.find || action.replace === undefined) break;
          const results = doc.body.search(action.find, {
            matchCase: action.matchCase ?? false,
          });
          results.load('items');
          await context.sync();
          results.items.forEach((item) => {
            item.insertText(action.replace, Word.InsertLocation.replace);
          });
          break;
        }

        // ── Set Document Property ────────────────────────────────
        case 'setDocumentProperty': {
          if (!action.property || action.value === undefined) break;
          const props = doc.properties;
          props.load('customProperties');
          await context.sync();
          const prop = action.property.toLowerCase();
          if (prop === 'title')   { props.title = action.value; }
          else if (prop === 'author')  { props.author = action.value; }
          else if (prop === 'subject') { props.subject = action.value; }
          else {
            const customProps = props.customProperties;
            customProps.add(action.property, action.value);
          }
          break;
        }

        default:
          console.warn(`Unknown Word command: ${(action as any).command}`);
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Failed to execute Word action:", error);
    throw error;
  }
}
