import type { WordAction } from '../types/actions';

export async function executeWordAction(action: WordAction) {
  try {
    await Word.run(async (context) => {
      const doc = context.document;
      const body = doc.body;

      switch (action.command) {

        // ── Insert Text ──────────────────────────────────────────────────────
        case 'insertText': {
          if (!action.text) break;
          const selection = doc.getSelection();
          selection.insertText(action.text, Word.InsertLocation.after);
          break;
        }

        // ── Insert HTML ──────────────────────────────────────────────────────
        case 'insertHtml': {
          if (!action.html) break;
          const selection = doc.getSelection();
          selection.insertHtml(action.html, Word.InsertLocation.after);
          break;
        }

        // ── Replace Selection ────────────────────────────────────────────────
        case 'replaceSelection': {
          if (!action.text) break;
          const selection = doc.getSelection();
          selection.insertText(action.text, Word.InsertLocation.replace);
          break;
        }

        // ── Apply Style ──────────────────────────────────────────────────────
        case 'applyStyle': {
          if (!action.styleName) break;
          if (action.targetText) {
            const results = body.search(action.targetText, { matchCase: false });
            results.load('items');
            await context.sync();
            results.items.forEach((item) => { item.style = action.styleName; });
          } else {
            const selection = doc.getSelection();
            selection.style = action.styleName;
          }
          break;
        }

        // ── Insert Page Break ────────────────────────────────────────────────
        case 'insertPageBreak': {
          const selection = doc.getSelection();
          selection.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
          break;
        }

        // ── Insert Section Break ─────────────────────────────────────────────
        case 'insertSectionBreak': {
          const a = action as any;
          const selection = doc.getSelection();
          const breakType = a.breakType === 'continuous'
            ? Word.BreakType.sectionContinuous
            : a.breakType === 'evenPage'
              ? Word.BreakType.sectionEven
              : Word.BreakType.sectionNext;
          selection.insertBreak(breakType, Word.InsertLocation.after);
          break;
        }

        // ── Insert TOC ──────────────────────────────────────────────────────
        case 'insertTOC': {
          const levels = action.levels ?? 3;
          const tocHtml = `
            <p><strong>Table of Contents</strong></p>
            <p><em>Auto-generated TOC for Headings 1–${levels}.</em></p>
            <p style="color:#888;font-size:11pt;">➡ In Word: References → Table of Contents → Automatic Table to generate the live TOC.</p>
          `;
          const selection = doc.getSelection();
          selection.insertHtml(tocHtml, Word.InsertLocation.before);
          break;
        }

        // ── Find and Replace ─────────────────────────────────────────────────
        case 'findAndReplace': {
          if (!action.find || action.replace === undefined) break;
          const results = body.search(action.find, { matchCase: action.matchCase ?? false });
          results.load('items');
          await context.sync();
          results.items.forEach((item) => {
            item.insertText(action.replace, Word.InsertLocation.replace);
          });
          break;
        }

        // ── Set Document Property ────────────────────────────────────────────
        case 'setDocumentProperty': {
          if (!action.property || action.value === undefined) break;
          const props = doc.properties;
          await context.sync();
          const prop = action.property.toLowerCase();
          if (prop === 'title')        props.title   = action.value;
          else if (prop === 'author')  props.author  = action.value;
          else if (prop === 'subject') props.subject = action.value;
          else {
            props.customProperties.add(action.property, action.value);
          }
          break;
        }

        // ── Insert Paragraph ────────────────────────────────────────────────
        case 'insertParagraph': {
          const a = action as any;
          const para = body.insertParagraph(a.text ?? '', Word.InsertLocation.end);
          if (a.styleName) para.style = a.styleName;
          if (a.bold !== undefined) para.font.bold = a.bold;
          if (a.fontSize)  para.font.size  = a.fontSize;
          if (a.fontColor) para.font.color = a.fontColor;
          if (a.alignment) para.alignment = a.alignment;
          break;
        }

        // ── Insert Heading ───────────────────────────────────────────────────
        case 'insertHeading': {
          const a = action as any;
          const level = Math.min(Math.max(a.level ?? 1, 1), 6);
          const para = body.insertParagraph(a.text ?? '', Word.InsertLocation.end);
          para.style = `Heading ${level}`;
          break;
        }

        // ── Insert Bullet List ───────────────────────────────────────────────
        case 'insertBulletList': {
          const a = action as any;
          const items: string[] = a.items ?? [];
          const listPara = body.insertParagraph('', Word.InsertLocation.end);
          listPara.style = 'List Paragraph';

          for (let i = 0; i < items.length; i++) {
            const p = i === 0 ? listPara : body.insertParagraph(items[i], Word.InsertLocation.end);
            if (i === 0) p.insertText(items[0], Word.InsertLocation.replace);
            p.style = 'List Paragraph';
            (p as any).listItem?.level ?? p;
          }
          break;
        }

        // ── Insert Table ─────────────────────────────────────────────────────
        case 'insertTable': {
          const a = action as any;
          const tableData: string[][] = a.tableData ?? [];
          if (!tableData.length) break;

          const rows = tableData.length;
          const cols = tableData[0].length;
          const selection = doc.getSelection();
          const table = selection.insertTable(rows, cols, Word.InsertLocation.after, tableData);
          table.load('rows');
          await context.sync();

          // Style header row
          if (a.headerRow !== false) {
            const firstRow = table.rows.getFirst();
            firstRow.load('cells');
            await context.sync();
            firstRow.cells.load('items');
            await context.sync();
            firstRow.cells.items.forEach((cell) => {
              (cell as any).shading.color = a.headerColor ?? '#2E75B6';
              cell.body.paragraphs.getFirst().font.color = '#FFFFFF';
              cell.body.paragraphs.getFirst().font.bold = true;
            });
          }
          break;
        }

        // ── Insert Comment ───────────────────────────────────────────────────
        case 'insertComment': {
          const a = action as any;
          if (!a.commentText) break;

          let range: Word.Range;
          if (a.targetText) {
            const results = body.search(a.targetText, { matchCase: false });
            results.load('items');
            await context.sync();
            if (!results.items.length) break;
            range = results.items[0];
          } else {
            range = doc.getSelection();
          }
          (range as any).insertComment(a.commentText);
          break;
        }

        // ── Format Selection ─────────────────────────────────────────────────
        case 'formatSelection': {
          const a = action as any;
          const sel = doc.getSelection();
          if (a.bold !== undefined)      sel.font.bold      = a.bold;
          if (a.italic !== undefined)    sel.font.italic    = a.italic;
          if (a.underline !== undefined) sel.font.underline = a.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
          if (a.fontSize)                sel.font.size      = a.fontSize;
          if (a.fontColor)               sel.font.color     = a.fontColor;
          if (a.highlightColor)          sel.font.highlightColor = a.highlightColor;
          if (a.styleName)               sel.style          = a.styleName;
          break;
        }

        // ── Insert Horizontal Rule ───────────────────────────────────────────
        case 'insertHorizontalRule': {
          const sel = doc.getSelection();
          sel.insertHtml('<hr style="border:1px solid #CCCCCC;"/>', Word.InsertLocation.after);
          break;
        }

        // ── Clear Document ───────────────────────────────────────────────────
        case 'clearDocument': {
          // body.clear() only clears formatting; insertText replace truly empties it
          body.insertText('', Word.InsertLocation.replace);
          break;
        }

        // ── Delete Content (alias for clearDocument) ─────────────────────────
        case 'deleteContent': {
          body.insertText('', Word.InsertLocation.replace);
          break;
        }

        // ── Delete Selected Text ─────────────────────────────────────────────
        case 'deleteSelection': {
          const sel = doc.getSelection();
          sel.insertText('', Word.InsertLocation.replace);
          break;
        }

        // ── Delete Found Text ────────────────────────────────────────────────
        case 'deleteText': {
          const a = action as any;
          if (!a.text) break;
          const results = body.search(a.text, { matchCase: false, matchWholeWord: false });
          results.load('items');
          await context.sync();
          results.items.forEach((item) => {
            item.insertText('', Word.InsertLocation.replace);
          });
          break;
        }

        // ── Insert at End ────────────────────────────────────────────────────
        case 'insertAtEnd': {
          const a = action as any;
          if (a.html) {
            body.insertHtml(a.html, Word.InsertLocation.end);
          } else if (a.text) {
            body.insertText(a.text, Word.InsertLocation.end);
          }
          break;
        }

        // ── Append Section ───────────────────────────────────────────────────
        case 'appendSection': {
          const a = action as any;
          if (a.heading) {
            const level = a.level ?? 1;
            const headingPara = body.insertParagraph(a.heading, Word.InsertLocation.end);
            headingPara.style = `Heading ${level}`;
          }
          if (a.content) {
            if (a.contentType === 'html') {
              body.insertHtml(a.content, Word.InsertLocation.end);
            } else {
              body.insertParagraph(a.content, Word.InsertLocation.end);
            }
          }
          break;
        }

        // ── Track Changes ────────────────────────────────────────────────────
        case 'trackChanges': {
          const a = action as any;
          // Word JS API v1.4+
          try {
            (doc as any).changeTrackingMode = a.enable
              ? Word.ChangeTrackingMode.trackAll
              : Word.ChangeTrackingMode.off;
          } catch {
            console.warn('[RiskVir AI] Track changes requires Word API 1.4+');
          }
          break;
        }

        // ── Accept All Changes ───────────────────────────────────────────────
        case 'acceptAllChanges': {
          try {
            doc.body.paragraphs.load('items');
            await context.sync();
            // Accept tracked changes via body revision marks
            (doc as any).revisions?.acceptAll();
          } catch {
            console.warn('[RiskVir AI] acceptAllChanges requires higher API version');
          }
          break;
        }

        default:
          console.warn(`[RiskVir AI] Unknown Word command: ${(action as any).command}`);
      }

      await context.sync();
    });
  } catch (error) {
    console.error('[RiskVir AI] Word action failed:', error);
    throw error;
  }
}
