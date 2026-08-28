export async function getCurrentSlideText(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;

    // Load the selected slides collection
    const selectedSlides = presentation.getSelectedSlides();
    selectedSlides.load('items');
    await context.sync();

    if (selectedSlides.items.length === 0) return '';

    const slide = selectedSlides.items[0];

    // Load shapes and their text properties
    slide.shapes.load('items');
    await context.sync();

    // Load textFrame for all shapes in one sync
    slide.shapes.items.forEach((shape: PowerPoint.Shape) => {
      if (shape.textFrame) {
        shape.textFrame.load('textRange');
      }
    });
    await context.sync();

    const parts: string[] = [];

    slide.shapes.items.forEach((shape: PowerPoint.Shape) => {
      try {
        const frame = shape.textFrame;
        if (!frame) return;

        const txt = (frame as any).textRange?.text?.trim();
        if (txt) parts.push(txt);
      } catch {
        // Shape may not have a textFrame — ignore
      }
    });

    return parts.join('\n');
  });
}
