/* global PowerPoint, Office */
import type { FontStyle, ParagraphStyle, ApplyTarget, ShapeStyleRecord } from '../store/useStore';
import { hasPlaceholderTypeSupport, hasLineSpacingSupport, isTitleShape, isBodyShape, isPowerPointReady } from './officeService';

export interface ApplyOptions {
  font: FontStyle;
  paragraph?: ParagraphStyle;
}

type ShapeFilter = 'all' | 'title' | 'body';

function applyFontToRange(range: PowerPoint.TextRange, font: FontStyle): void {
  if (font.name !== undefined) range.font.name = font.name;
  if (font.size !== undefined) range.font.size = font.size;
  if (font.bold !== undefined) range.font.bold = font.bold;
  if (font.italic !== undefined) range.font.italic = font.italic;
  // underline은 Office.js에서 ShapeFontUnderlineStyle('Single'|'None') 사용
  if (font.underline !== undefined)
    (range.font as unknown as { underline: string }).underline = font.underline ? 'Single' : 'None';
  if (font.color !== undefined) range.font.color = font.color;
}

function applyParagraphToFrame(
  range: PowerPoint.TextRange,
  paragraph: ParagraphStyle,
  showLineSpacingWarning: () => void
): void {
  if (paragraph.alignment !== undefined) {
    const alignMap: Record<string, PowerPoint.ParagraphHorizontalAlignment> = {
      left: PowerPoint.ParagraphHorizontalAlignment.left,
      center: PowerPoint.ParagraphHorizontalAlignment.center,
      right: PowerPoint.ParagraphHorizontalAlignment.right,
      justify: PowerPoint.ParagraphHorizontalAlignment.justify,
    };
    const align = alignMap[paragraph.alignment];
    if (align) range.paragraphFormat.horizontalAlignment = align;
  }

  if (paragraph.lineSpacing !== undefined) {
    if (hasLineSpacingSupport()) {
      // spaceWithin: 줄간격 배수 (100 = 1.0배, PresentationAPI 1.5+)
      (range.paragraphFormat as unknown as { spaceWithin: number }).spaceWithin =
        paragraph.lineSpacing * 100;
    } else {
      showLineSpacingWarning();
    }
  }
}

/** 슬라이드 하나의 shapes를 순회하며 스타일 적용 */
async function applyToSlide(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  options: ApplyOptions,
  filter: ShapeFilter,
  showLineSpacingWarning: () => void
): Promise<void> {
  const shapes = slide.shapes;
  shapes.load('items/name,items/textFrame,items/placeholderType');
  await context.sync();

  for (const shape of shapes.items) {
    if (!shape.textFrame) continue;

    // placeholderType은 PresentationAPI 1.3+ 전용 — any 캐스팅으로 런타임 접근
    const shapeAny = shape as unknown as { placeholderType?: string };
    if (filter === 'title') {
      const isTitle = hasPlaceholderTypeSupport()
        ? shapeAny.placeholderType === PowerPoint.PlaceholderType.title
        : isTitleShape(shape.name);
      if (!isTitle) continue;
    } else if (filter === 'body') {
      const isBody = hasPlaceholderTypeSupport()
        ? shapeAny.placeholderType === PowerPoint.PlaceholderType.body
        : isBodyShape(shape.name);
      if (!isBody) continue;
    }

    applyFontToRange(shape.textFrame.textRange, options.font);
    if (options.paragraph) {
      applyParagraphToFrame(shape.textFrame.textRange, options.paragraph, showLineSpacingWarning);
    }
  }
}

/** 메인 스타일 적용 함수 — 모든 ApplyTarget을 통합 처리 */
export async function applyStyle(
  target: ApplyTarget,
  options: ApplyOptions,
  onLineSpacingUnsupported: () => void
): Promise<void> {
  if (!isPowerPointReady()) {
    throw new Error('PowerPoint 내에서 실행해주세요. 일반 브라우저에서는 사용할 수 없습니다.');
  }

  if (target === 'selection-text' || target === 'selection-shape') {
    await PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load('items/textFrame,items/name');
      await context.sync();

      for (const shape of selectedShapes.items) {
        if (!shape.textFrame) continue;
        // getSelectedTextRange는 PresentationAPI 1.5+ — any 캐스팅
        const range =
          target === 'selection-text'
            ? (shape.textFrame as unknown as { getSelectedTextRange: () => PowerPoint.TextRange }).getSelectedTextRange()
            : shape.textFrame.textRange;
        applyFontToRange(range, options.font);
        if (options.paragraph) {
          applyParagraphToFrame(range, options.paragraph, onLineSpacingUnsupported);
        }
      }
      await context.sync();
    });
    return;
  }

  const filter: ShapeFilter =
    target === 'current-title' || target === 'all-titles'
      ? 'title'
      : target === 'current-body' || target === 'all-bodies'
      ? 'body'
      : 'all';

  if (target.startsWith('current-')) {
    await PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      for (const slide of selectedSlides.items) {
        await applyToSlide(context, slide, options, filter, onLineSpacingUnsupported);
      }
      await context.sync();
    });
  } else {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      for (const slide of slides.items) {
        await applyToSlide(context, slide, options, filter, onLineSpacingUnsupported);
      }
      await context.sync();
    });
  }
}

/**
 * 적용 전 스타일 스냅샷 캡처 (실행 취소용)
 * selection-text/shape는 range 단위 캡처가 복잡해 슬라이드 단위만 지원
 */
export async function captureSnapshot(target: ApplyTarget): Promise<ShapeStyleRecord[]> {
  const records: ShapeStyleRecord[] = [];

  if (target === 'selection-text' || target === 'selection-shape') {
    // 선택 범위 단위 스냅샷은 지원하지 않음 (빈 배열 반환)
    return records;
  }

  const filter: ShapeFilter =
    target === 'current-title' || target === 'all-titles'
      ? 'title'
      : target === 'current-body' || target === 'all-bodies'
      ? 'body'
      : 'all';

  await PowerPoint.run(async (context) => {
    let slideItems: PowerPoint.Slide[];

    if (target.startsWith('current-')) {
      const sel = context.presentation.getSelectedSlides();
      sel.load('items');
      await context.sync();
      slideItems = sel.items;
    } else {
      const all = context.presentation.slides;
      all.load('items');
      await context.sync();
      slideItems = all.items;
    }

    for (let si = 0; si < slideItems.length; si++) {
      const shapes = slideItems[si].shapes;
      shapes.load('items/name,items/textFrame,items/placeholderType');
      await context.sync();

      for (let shi = 0; shi < shapes.items.length; shi++) {
        const shape = shapes.items[shi];
        if (!shape.textFrame) continue;

        const shapeAny = shape as unknown as { placeholderType?: string };
        if (filter === 'title') {
          const isTitle = hasPlaceholderTypeSupport()
            ? shapeAny.placeholderType === PowerPoint.PlaceholderType.title
            : isTitleShape(shape.name);
          if (!isTitle) continue;
        } else if (filter === 'body') {
          const isBody = hasPlaceholderTypeSupport()
            ? shapeAny.placeholderType === PowerPoint.PlaceholderType.body
            : isBodyShape(shape.name);
          if (!isBody) continue;
        }

        // 폰트 속성 로드
        const range = shape.textFrame.textRange;
        range.load('font/name,font/size,font/bold,font/italic,font/underline,font/color');
        range.paragraphFormat.load('horizontalAlignment');
        await context.sync();

        const underlineVal = (range.font as unknown as { underline?: string }).underline;
        records.push({
          slideIndex: si,
          shapeIndex: shi,
          font: {
            name: range.font.name ?? undefined,
            size: range.font.size ?? undefined,
            bold: range.font.bold ?? undefined,
            italic: range.font.italic ?? undefined,
            // 'Single' → true, 'None'/'null' → false
            underline: underlineVal != null ? underlineVal !== 'None' : undefined,
            color: range.font.color ?? undefined,
          },
          alignment: range.paragraphFormat.horizontalAlignment ?? undefined,
        });
      }
    }
  });

  return records;
}

/**
 * 스냅샷 복원 (실행 취소)
 */
export async function restoreSnapshot(records: ShapeStyleRecord[]): Promise<void> {
  if (records.length === 0) return;

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    for (const record of records) {
      const slide = slides.items[record.slideIndex];
      if (!slide) continue;

      const shapes = slide.shapes;
      shapes.load('items/textFrame');
      await context.sync();

      const shape = shapes.items[record.shapeIndex];
      if (!shape?.textFrame) continue;

      const range = shape.textFrame.textRange;
      applyFontToRange(range, record.font);

      if (record.alignment) {
        const alignMap: Record<string, PowerPoint.ParagraphHorizontalAlignment> = {
          left: PowerPoint.ParagraphHorizontalAlignment.left,
          center: PowerPoint.ParagraphHorizontalAlignment.center,
          right: PowerPoint.ParagraphHorizontalAlignment.right,
          justify: PowerPoint.ParagraphHorizontalAlignment.justify,
          Left: PowerPoint.ParagraphHorizontalAlignment.left,
          Center: PowerPoint.ParagraphHorizontalAlignment.center,
          Right: PowerPoint.ParagraphHorizontalAlignment.right,
          Justify: PowerPoint.ParagraphHorizontalAlignment.justify,
        };
        const align = alignMap[record.alignment];
        if (align) range.paragraphFormat.horizontalAlignment = align;
      }
    }
    await context.sync();
  });
}
