import type { PowerPointAction } from '../types/actions';

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Get the last (or specific index) slide from a loaded collection */
function getSlide(slides: PowerPoint.SlideCollection, idx?: number) {
  const items = slides.items;
  const resolvedIdx = (idx === undefined || idx === -1) ? items.length - 1 : idx;
  if (resolvedIdx < 0 || resolvedIdx >= items.length) return null;
  return items[resolvedIdx];
}

/** Set bold title text box at top of slide */
async function addTitleBox(slide: PowerPoint.Slide, title: string, fontSize = 28) {
  const shapes = slide.shapes;

  // Try using existing title placeholder first
  const loaded = slide.shapes as any;
  const titleShape = loaded.items?.find(
    (s: any) => s.placeholderType === 'Title' || s.placeholderType === 'CenteredTitle'
  ) as any;

  if (titleShape) {
    titleShape.textFrame.textRange.text = title;
    titleShape.textFrame.textRange.font.size = fontSize;
    titleShape.textFrame.textRange.font.bold = true;
    return;
  }

  // Fallback: add a text box
  const tb = shapes.addTextBox(title) as any;
  tb.left = 30; tb.top = 20; tb.width = 660; tb.height = 65;
  if (tb.textFrame) {
    tb.textFrame.textRange.font.size = fontSize;
    tb.textFrame.textRange.font.bold = true;
    tb.textFrame.textRange.font.color = '#1F3864';
  }
}

/** Add a new slide and return it */
async function addNewSlide(
  presentation: PowerPoint.Presentation,
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Slide> {
  presentation.slides.add();
  await context.sync();
  presentation.slides.load('items');
  await context.sync();
  const items = presentation.slides.items;
  const slide = items[items.length - 1];
  slide.shapes.load('items');
  await context.sync();
  return slide;
}

// ─── Main Export ─────────────────────────────────────────────────────────────

export async function executePptAction(action: PowerPointAction) {
  try {
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;

      switch (action.command) {

        // ── Add Slide ────────────────────────────────────────────────────────
        case 'addSlide': {
          const slide = await addNewSlide(presentation, context);

          if (action.title) {
            await addTitleBox(slide, action.title);
          }

          if (action.bulletPoints?.length) {
            const bulletText = action.bulletPoints.map(b => `• ${b}`).join('\n');
            const shapes = slide.shapes.items;
            const bodyShape = shapes.find(
              (s: any) => s.placeholderType === 'Body' || s.placeholderType === 'Subtitle'
            ) as any;

            if (bodyShape) {
              bodyShape.textFrame.textRange.text = bulletText;
              bodyShape.textFrame.textRange.font.size = 18;
            } else {
              const tb = slide.shapes.addTextBox(bulletText) as any;
              tb.left = 40; tb.top = 110; tb.width = 640; tb.height = 300;
              if (tb.textFrame) {
                tb.textFrame.textRange.font.size = 18;
                tb.textFrame.textRange.font.color = '#333333';
              }
            }
          }
          break;
        }

        // ── Add Slide with Table ─────────────────────────────────────────────
        case 'addSlideWithTable': {
          const slide = await addNewSlide(presentation, context);

          if (action.title) await addTitleBox(slide, action.title);

          if (action.tableData) {
            const { headers, rows } = action.tableData;
            const colCount = headers.length;
            const rowCount = rows.length + 1;

            const table = slide.shapes.addTable(rowCount, colCount) as any;
            table.left = 30; table.top = 105; table.width = 660;

            headers.forEach((h, ci) => {
              const cell = table.rows.getItemAt(0).cells.getItemAt(ci) as any;
              cell.text = h;
              if (cell.fill) cell.fill.setSolidColor(action.headerBackgroundColor || '#1F3864');
              if (cell.textFrame) {
                cell.textFrame.textRange.font.color = action.headerFontColor || '#FFFFFF';
                cell.textFrame.textRange.font.bold = true;
                cell.textFrame.textRange.font.size = 14;
              }
            });

            rows.forEach((row, ri) => {
              row.forEach((val, ci) => {
                const cell = table.rows.getItemAt(ri + 1).cells.getItemAt(ci) as any;
                cell.text = String(val);
                if (cell.textFrame) {
                  cell.textFrame.textRange.font.size = 13;
                  // Alternate row shading
                  if (ri % 2 === 1 && cell.fill) {
                    cell.fill.setSolidColor('#EBF3FB');
                  }
                }
              });
            });
          }
          break;
        }

        // ── Add Slide with Chart (data summary) ──────────────────────────────
        case 'addSlideWithChart': {
          const slide = await addNewSlide(presentation, context);
          const title = action.title || 'Chart';
          await addTitleBox(slide, title);

          if (action.chartData) {
            const { labels, series } = action.chartData;

            // Build a clean data summary table as visual substitute
            // (PowerPoint JS API has no direct chart insert for Add-ins)
            const colHeaders = ['Category', ...series.map(s => s.name)];
            const rowCount = labels.length + 1;
            const colCount = colHeaders.length;

            const table = slide.shapes.addTable(rowCount, colCount) as any;
            table.left = 30; table.top = 105; table.width = 660;

            colHeaders.forEach((h, ci) => {
              const cell = table.rows.getItemAt(0).cells.getItemAt(ci) as any;
              cell.text = h;
              if (cell.fill) cell.fill.setSolidColor('#1F3864');
              if (cell.textFrame) {
                cell.textFrame.textRange.font.color = '#FFFFFF';
                cell.textFrame.textRange.font.bold = true;
                cell.textFrame.textRange.font.size = 13;
              }
            });

            labels.forEach((label, ri) => {
              const labelCell = table.rows.getItemAt(ri + 1).cells.getItemAt(0) as any;
              labelCell.text = String(label);
              if (ri % 2 === 1 && labelCell.fill) labelCell.fill.setSolidColor('#EBF3FB');

              series.forEach((s, si) => {
                const dataCell = table.rows.getItemAt(ri + 1).cells.getItemAt(si + 1) as any;
                dataCell.text = String(s.values[ri] ?? '');
                if (ri % 2 === 1 && dataCell.fill) dataCell.fill.setSolidColor('#EBF3FB');
              });
            });

            // Add chart type info note
            const note = slide.shapes.addTextBox(`📊 Chart Type: ${action.chartType || 'column'} | X: ${action.xAxisLabel || ''} | Y: ${action.yAxisLabel || ''}`) as any;
            note.left = 30; note.top = 420; note.width = 660; note.height = 30;
            if (note.textFrame) {
              note.textFrame.textRange.font.size = 11;
              note.textFrame.textRange.font.italic = true;
              note.textFrame.textRange.font.color = '#666666';
            }
          }
          break;
        }

        // ── Add Slide with Image Placeholder ────────────────────────────────
        case 'addSlideWithImage': {
          const a = action as any;
          const slide = await addNewSlide(presentation, context);
          if (a.title) await addTitleBox(slide, a.title);

          const placeholder = slide.shapes.addTextBox(
            `🖼️ ${a.imagePlaceholderText || '[Insert image here]'}`
          ) as any;
          placeholder.left = 80; placeholder.top = 110; placeholder.width = 560; placeholder.height = 250;
          if (placeholder.textFrame) {
            placeholder.textFrame.textRange.font.size = 20;
            placeholder.textFrame.textRange.font.color = '#AAAAAA';
            placeholder.textFrame.textRange.font.italic = true;
          }
          if ((placeholder as any).fill) (placeholder as any).fill.setSolidColor('#F5F5F5');

          if (a.caption) {
            const caption = slide.shapes.addTextBox(a.caption) as any;
            caption.left = 80; caption.top = 375; caption.width = 560; caption.height = 35;
            if (caption.textFrame) {
              caption.textFrame.textRange.font.size = 12;
              caption.textFrame.textRange.font.italic = true;
              caption.textFrame.textRange.font.color = '#555555';
            }
          }
          break;
        }

        // ── Add Speaker Notes ────────────────────────────────────────────────
        case 'addSpeakerNotes': {
          if (!action.notes) break;
          presentation.slides.load('items');
          await context.sync();

          const slide = getSlide(presentation.slides, action.slideIndex);
          if (!slide) break;

          // The PowerPoint JS API exposes notes via the notes property
          try {
            (slide as any).notes = action.notes;
          } catch {
            // Fallback: add a text box at bottom as notes stand-in
            slide.shapes.load('items');
            await context.sync();
            const nb = slide.shapes.addTextBox(`📝 Notes: ${action.notes}`) as any;
            nb.left = 10; nb.top = 460; nb.width = 700; nb.height = 60;
            if (nb.textFrame) {
              nb.textFrame.textRange.font.size = 10;
              nb.textFrame.textRange.font.color = '#888888';
              nb.textFrame.textRange.font.italic = true;
            }
          }
          break;
        }

        // ── Format Slide ─────────────────────────────────────────────────────
        case 'formatSlide': {
          presentation.slides.load('items');
          await context.sync();
          const slide = getSlide(presentation.slides, action.slideIndex);
          if (!slide) break;

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
              font.bold = true;
            } else {
              if (action.bodyFontSize)  font.size = action.bodyFontSize;
            }
          });

          if (action.backgroundColor && (slide as any).background) {
            (slide as any).background.fill.setSolidColor(action.backgroundColor);
          }
          break;
        }

        // ── Move Slide ───────────────────────────────────────────────────────
        case 'moveSlide': {
          const a = action as any;
          presentation.slides.load('items');
          await context.sync();
          const slide = getSlide(presentation.slides, a.fromIndex);
          if (!slide) break;
          // PowerPoint JS API does not expose setPosition — use workaround via delete+re-add
          // Best we can do is notify; actual reorder requires VBA or user action
          console.warn('[RiskVir AI] moveSlide: PowerPoint JS API does not support slide reordering directly.');
          break;
        }

        // ── Set Theme (color + font) ─────────────────────────────────────────
        case 'setTheme': {
          const a = action as any;
          // Apply to all existing slides via formatSlide logic
          presentation.slides.load('items');
          await context.sync();
          presentation.slides.items.forEach((slide: any) => {
            if (a.backgroundColor && slide.background) {
              slide.background.fill.setSolidColor(a.backgroundColor);
            }
          });
          break;
        }

        // ── Delete Slide ─────────────────────────────────────────────────────
        case 'deleteSlide': {
          const a = action as any;
          presentation.slides.load('items');
          await context.sync();
          const slide = getSlide(presentation.slides, a.slideIndex);
          if (slide) slide.delete();
          break;
        }

        // ── Duplicate Slide ──────────────────────────────────────────────────
        case 'duplicateSlide': {
          const a = action as any;
          presentation.slides.load('items');
          await context.sync();
          const src = getSlide(presentation.slides, a.slideIndex ?? 0);
          if (!src) break;
          // PowerPoint JS API doesn't expose slide duplication directly
          console.warn('[RiskVir AI] duplicateSlide: not directly supported by PowerPoint JS API.');
          break;
        }

        // ── Edit Shape Text ──────────────────────────────────────────────────
        case 'editShapeText': {
          const a = action as any;
          presentation.slides.load('items');
          await context.sync();
          const slide = getSlide(presentation.slides, a.slideIndex ?? -1);
          if (!slide) break;
          slide.shapes.load('items');
          await context.sync();

          const shape = slide.shapes.items.find((s: any) =>
            (a.shapeName && s.name === a.shapeName) ||
            (a.shapeIndex !== undefined && slide.shapes.items.indexOf(s) === a.shapeIndex)
          ) as any;

          if (shape?.textFrame) {
            shape.textFrame.textRange.text = a.newText ?? '';
            if (a.fontSize) shape.textFrame.textRange.font.size = a.fontSize;
            if (a.fontColor) shape.textFrame.textRange.font.color = a.fontColor;
          }
          break;
        }

        default:
          console.warn(`[RiskVir AI] Unknown PowerPoint command: ${(action as any).command}`);
      }

      await context.sync();
    });
  } catch (error) {
    console.error('[RiskVir AI] PowerPoint action failed:', error);
    throw error;
  }
}
