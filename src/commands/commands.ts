/* global Office, PowerPoint */

Office.onReady(() => {
  // 리본 버튼 함수 등록
});

/**
 * 리본 "제목 일괄 적용" 버튼 핸들러
 * 저장된 첫 번째 프리셋의 폰트를 모든 슬라이드 제목에 적용
 */
async function applyToAllTitles(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const presetsRaw = Office.context.document.settings.get('ppt-style-addin-presets');
      const presets = presetsRaw ? JSON.parse(presetsRaw) : [];
      const preset = presets[0];
      if (!preset) return;

      for (const slide of slides.items) {
        const shapes = slide.shapes;
        shapes.load('items/name,items/textFrame');
        await context.sync();

        for (const shape of shapes.items) {
          // placeholderType은 PresentationAPI 1.3+ — any 캐스팅 후 이름 fallback 병용
          const shapeAny = shape as unknown as { placeholderType?: string };
          const isTitle =
            shapeAny.placeholderType === PowerPoint.PlaceholderType.title ||
            isTitleShapeName(shape.name);
          if (!isTitle || !shape.textFrame) continue;

          applyFont(shape.textFrame.textRange, preset.font);
        }
      }
      await context.sync();
    });
  } catch (e) {
    console.error('applyToAllTitles error:', e);
  } finally {
    event.completed();
  }
}

/**
 * 리본 "본문 일괄 적용" 버튼 핸들러
 */
async function applyToAllBodies(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const presetsRaw = Office.context.document.settings.get('ppt-style-addin-presets');
      const presets = presetsRaw ? JSON.parse(presetsRaw) : [];
      const preset = presets[0];
      if (!preset) return;

      for (const slide of slides.items) {
        const shapes = slide.shapes;
        shapes.load('items/name,items/textFrame');
        await context.sync();

        for (const shape of shapes.items) {
          const shapeAny = shape as unknown as { placeholderType?: string };
          const isBody =
            shapeAny.placeholderType === PowerPoint.PlaceholderType.body ||
            isBodyShapeName(shape.name);
          if (!isBody || !shape.textFrame) continue;

          applyFont(shape.textFrame.textRange, preset.font);
        }
      }
      await context.sync();
    });
  } catch (e) {
    console.error('applyToAllBodies error:', e);
  } finally {
    event.completed();
  }
}

function applyFont(range: PowerPoint.TextRange, font: Record<string, unknown>) {
  if (font.name) range.font.name = font.name as string;
  if (font.size) range.font.size = font.size as number;
  if (font.bold !== undefined) range.font.bold = font.bold as boolean;
  if (font.italic !== undefined) range.font.italic = font.italic as boolean;
  if (font.underline !== undefined)
    (range.font as unknown as { underline: string }).underline =
      font.underline ? 'Single' : 'None';
  if (font.color) range.font.color = font.color as string;
}

function isTitleShapeName(name: string): boolean {
  return ['title', 'Title', '제목'].some((k) => name.includes(k));
}

function isBodyShapeName(name: string): boolean {
  return ['content', 'Content', 'body', 'Body', '내용', '본문', 'Text'].some((k) =>
    name.includes(k)
  );
}

Office.actions.associate('applyToAllTitles', applyToAllTitles);
Office.actions.associate('applyToAllBodies', applyToAllBodies);
