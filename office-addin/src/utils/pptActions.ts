import type { PowerPointAction } from '../types/actions';

export async function executePptAction(action: PowerPointAction) {

  try {
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;

      switch (action.command) {

        // ── Add Slide ────────────────────────────────────────────
        case 'addSlide': {
          presentation.slides.add();
          await context.sync();

          // Get the last (newly added) slide
          presentation.slides.load('items');
          await context.sync();
          const slides = presentation.slides.items;
          const newSlide = slides[slides.length - 1];

          newSlide.shapes.load('items');
          await context.sync();

          // Use existing placeholder shapes if present, otherwise add text boxes
          const shapes = newSlide.shapes.items;
          const titleShape = shapes.find((s: any) => s.placeholderType === 'Title') as any;
          const bodyShape  = shapes.find((s: any) => s.placeholderType === 'Body' || s.placeholderType === 'Subtitle') as any;

          if (action.title) {
            if (titleShape) {
              titleShape.textFrame.textRange.text = action.title;
            } else {
              const tb = newSlide.shapes.addTextBox(action.title) as any;
              tb.left = 30; tb.top = 20; tb.width = 600; tb.height = 60;
              if (tb.textFrame) {
                tb.textFrame.textRange.font.size = 28;
                tb.textFrame.textRange.font.bold = true;
              }
            }
          }

          if (action.bulletPoints && Array.isArray(action.bulletPoints)) {
            const bulletText = action.bulletPoints.join('\n');
            if (bodyShape) {
              bodyShape.textFrame.textRange.text = bulletText;
            } else {
              const tb = newSlide.shapes.addTextBox(bulletText) as any;
              tb.left = 30; tb.top = 100; tb.width = 600; tb.height = 300;
              if (tb.textFrame) tb.textFrame.textRange.font.size = 18;
            }
          }
          break;
        }

        // ── Add Slide With Table ─────────────────────────────────
        case 'addSlideWithTable': {
          presentation.slides.add();
          await context.sync();

          presentation.slides.load('items');
          await context.sync();
          const slides = presentation.slides.items;
          const newSlide = slides[slides.length - 1];

          if (action.title) {
            const tb = newSlide.shapes.addTextBox(action.title) as any;
            tb.left = 30; tb.top = 20; tb.width = 600; tb.height = 60;
            if (tb.textFrame) {
              tb.textFrame.textRange.font.size = 24;
              tb.textFrame.textRange.font.bold = true;
            }
          }

          if (action.tableData) {
            const { headers, rows } = action.tableData;
            const colCount = headers.length;
            const rowCount = rows.length + 1; // +1 for header row

            const table = newSlide.shapes.addTable(rowCount, colCount) as any;
            table.left = 30; table.top = 100; table.width = 620;

            // Header row
            headers.forEach((h, ci) => {
              const cell = table.rows.getItemAt(0).cells.getItemAt(ci) as any;
              cell.text = h;
              if (cell.fill) cell.fill.setSolidColor(action.headerBackgroundColor || '#2E75B6');
              if (cell.textFrame) {
                cell.textFrame.textRange.font.color = action.headerFontColor || '#FFFFFF';
                cell.textFrame.textRange.font.bold = true;
              }
            });

            // Data rows
            rows.forEach((row, ri) => {
              row.forEach((val, ci) => {
                const cell = table.rows.getItemAt(ri + 1).cells.getItemAt(ci) as any;
                cell.text = val;
              });
            });
          }
          break;
        }

        // ── Add Slide With Chart (descriptive placeholder) ───────
        case 'addSlideWithChart': {
          presentation.slides.add();
          await context.sync();

          presentation.slides.load('items');
          await context.sync();
          const slides = presentation.slides.items;
          const newSlide = slides[slides.length - 1];

          const title = action.title || 'Chart';
          const tb = newSlide.shapes.addTextBox(title) as any;
          tb.left = 30; tb.top = 20; tb.width = 600; tb.height = 60;
          if (tb.textFrame) {
            tb.textFrame.textRange.font.size = 24;
            tb.textFrame.textRange.font.bold = true;
          }

          // Build a descriptive placeholder (PowerPoint JS API has limited chart support)
          if (action.chartData) {
            const { labels, series } = action.chartData;
            const lines: string[] = [`Chart Type: ${action.chartType || 'column'}`];
            lines.push(`Categories: ${labels.join(', ')}`);
            series.forEach((s) => {
              lines.push(`${s.name}: ${s.values.join(', ')}`);
            });
            const desc = newSlide.shapes.addTextBox(lines.join('\n')) as any;
            desc.left = 30; desc.top = 100; desc.width = 620; desc.height = 300;
            if (desc.textFrame) desc.textFrame.textRange.font.size = 14;
          }
          break;
        }

        // ── Add Speaker Notes ────────────────────────────────────
        case 'addSpeakerNotes': {
          if (!action.notes) break;
          presentation.slides.load('items');
          await context.sync();

          const slides = presentation.slides.items;
          const idx = (action.slideIndex === undefined || action.slideIndex === -1)
            ? slides.length - 1
            : action.slideIndex;

          if (idx < 0 || idx >= slides.length) break;

          const slide = slides[idx] as any;
          if (slide.notesPage) {
            slide.notesPage.body.clear();
            slide.notesPage.body.insertText(action.notes, Word.InsertLocation.start);
          } else if (slide.notes !== undefined) {
            slide.notes = action.notes;
          }
          break;
        }

        // ── Format Slide ─────────────────────────────────────────
        case 'formatSlide': {
          presentation.slides.load('items');
          await context.sync();
          const slides = presentation.slides.items;
          const idx = action.slideIndex ?? 0;
          if (idx < 0 || idx >= slides.length) break;

          const slide = slides[idx];
          slide.shapes.load('items');
          await context.sync();

          slide.shapes.items.forEach((shape: any) => {
            if (!shape.textFrame) return;
            const isTitle =
              shape.placeholderType === 'Title' ||
              shape.placeholderType === 'CenteredTitle';

            const font = shape.textFrame.textRange?.font;
            if (!font) return;

            if (isTitle) {
              if (action.titleFontSize)  font.size  = action.titleFontSize;
              if (action.titleFontColor) font.color = action.titleFontColor;
            } else {
              if (action.bodyFontSize)   font.size  = action.bodyFontSize;
            }
          });

          if (action.backgroundColor && (slide as any).background) {
            (slide as any).background.fill.setSolidColor(action.backgroundColor);
          }
          break;
        }

        default:
          console.warn(`Unknown PowerPoint command: ${(action as any).command}`);
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Failed to execute PowerPoint action:", error);
    throw error;
  }
}
