export interface PptContext {
  totalSlides: number;
  currentSlideIndex: number;
  currentSlideText: string;
  allSlidesTitles: string[];
}

/** Returns rich context about the current presentation */
export async function getCurrentSlideText(): Promise<PptContext> {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;

    // Load all slides for titles
    presentation.slides.load('items');
    await context.sync();

    const allSlides = presentation.slides.items;
    const totalSlides = allSlides.length;

    // Load shapes for each slide to get titles
    allSlides.forEach((slide) => slide.shapes.load('items'));
    await context.sync();

    allSlides.forEach((slide) => {
      slide.shapes.items.forEach((shape: any) => {
        if (shape.textFrame) shape.textFrame.load('textRange');
      });
    });
    await context.sync();

    // Extract titles from all slides
    const allSlidesTitles = allSlides.map((slide, idx) => {
      const titleShape = slide.shapes.items.find(
        (s: any) => s.placeholderType === 'Title' || s.placeholderType === 'CenteredTitle'
      ) as any;
      const title = titleShape?.textFrame?.textRange?.text?.trim() || `Slide ${idx + 1}`;
      return title;
    });

    // Get selected slide text
    let currentSlideIndex = allSlides.length - 1;
    let currentSlideText = '';

    try {
      const selectedSlides = presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();

      if (selectedSlides.items.length > 0) {
        const selectedSlide = selectedSlides.items[0];
        // Find index
        const selectedId = (selectedSlide as any).id;
        const foundIdx = allSlides.findIndex((s: any) => s.id === selectedId);
        if (foundIdx >= 0) currentSlideIndex = foundIdx;

        selectedSlide.shapes.load('items');
        await context.sync();

        selectedSlide.shapes.items.forEach((shape: PowerPoint.Shape) => {
          if ((shape as any).textFrame) {
            (shape as any).textFrame.load('textRange');
          }
        });
        await context.sync();

        const parts: string[] = [];
        selectedSlide.shapes.items.forEach((shape: PowerPoint.Shape) => {
          try {
            const txt = (shape as any).textFrame?.textRange?.text?.trim();
            if (txt) parts.push(txt);
          } catch { /* skip */ }
        });
        currentSlideText = parts.join('\n');
      }
    } catch {
      // Fallback if getSelectedSlides not available
      const lastSlide = allSlides[allSlides.length - 1];
      if (lastSlide) {
        const parts: string[] = [];
        lastSlide.shapes.items.forEach((shape: any) => {
          try {
            const txt = shape.textFrame?.textRange?.text?.trim();
            if (txt) parts.push(txt);
          } catch { /* skip */ }
        });
        currentSlideText = parts.join('\n');
      }
    }

    return {
      totalSlides,
      currentSlideIndex,
      currentSlideText,
      allSlidesTitles,
    };
  });
}
