/* global PowerPoint, Office */
import type { FontStyle, ParagraphStyle, ApplyTarget, ShapeStyleRecord } from '../store/useStore';
import { hasLineSpacingSupport, isTitleShape, isBodyShape, isPowerPointReady } from './officeService';

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
  if (font.underline !== undefined) {
    try {
      (range.font as unknown as { underline: string }).underline = font.underline ? 'Single' : 'None';
    } catch {
      // underline 미지원 환경 무시
    }
  }
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
      (range.paragraphFormat as unknown as { spaceWithin: number }).spaceWithin =
        paragraph.lineSpacing * 100;
    } else {
      showLineSpacingWarning();
    }
  }
}

/** 개별 도형에 안전하게 스타일 적용 (textFrame 없는 도형은 건너뜀) */
async function safeApplyToShape(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  options: ApplyOptions,
  showLineSpacingWarning: () => void
): Promise<boolean> {
  try {
    const textFrame = shape.textFrame;
    const textRange = textFrame.textRange;
    textRange.load('text');
    await context.sync();

    applyFontToRange(textRange, options.font);
    if (options.paragraph) {
      applyParagraphToFrame(textRange, options.paragraph, showLineSpacingWarning);
    }
    await context.sync();
    return true;
  } catch {
    return false;
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
  shapes.load('items/name');
  await context.sync();

  for (const shape of shapes.items) {
    if (filter === 'title' && !isTitleShape(shape.name)) continue;
    if (filter === 'body' && !isBodyShape(shape.name)) continue;

    await safeApplyToShape(context, shape, options, showLineSpacingWarning);
  }
}

/** 위치 허용 오차 (포인트 단위) */
const POSITION_TOLERANCE = 50;

/** 두 도형이 같은 위치에 있는지 판정 (좌상단 좌표 기준) */
function isSamePosition(
  a: { left: number; top: number },
  b: { left: number; top: number }
): boolean {
  return (
    Math.abs(a.left - b.left) <= POSITION_TOLERANCE &&
    Math.abs(a.top - b.top) <= POSITION_TOLERANCE
  );
}

/** 선택한 도형의 위치를 기준으로 모든 슬라이드의 같은 위치 텍스트에 스타일 적용 */
async function applyToSamePosition(
  options: ApplyOptions,
  showLineSpacingWarning: () => void
): Promise<void> {
  await PowerPoint.run(async (context) => {
    // 1. 선택한 도형의 위치 수집
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load('items/left,items/top');
    await context.sync();

    if (selectedShapes.items.length === 0) {
      throw new Error('먼저 슬라이드에서 텍스트 상자를 선택해주세요.');
    }

    const targetPositions = selectedShapes.items.map((s) => ({
      left: s.left,
      top: s.top,
    }));

    // 2. 모든 슬라이드에서 같은 위치의 도형에 스타일 적용
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load('items/left,items/top');
      await context.sync();

      for (const shape of shapes.items) {
        const matched = targetPositions.some((tp) => isSamePosition(tp, {
          left: shape.left, top: shape.top,
        }));
        if (!matched) continue;

        await safeApplyToShape(context, shape, options, showLineSpacingWarning);
      }
    }
  });
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
      selectedShapes.load('items/name');
      await context.sync();

      for (const shape of selectedShapes.items) {
        try {
          const textFrame = shape.textFrame;
          const range =
            target === 'selection-text'
              ? (textFrame as unknown as { getSelectedTextRange: () => PowerPoint.TextRange }).getSelectedTextRange()
              : textFrame.textRange;
          range.load('text');
          await context.sync();

          applyFontToRange(range, options.font);
          if (options.paragraph) {
            applyParagraphToFrame(range, options.paragraph, onLineSpacingUnsupported);
          }
          await context.sync();
        } catch {
          continue;
        }
      }
    });
    return;
  }

  if (target === 'all-same-position') {
    await applyToSamePosition(options, onLineSpacingUnsupported);
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
    });
  } else {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      for (const slide of slides.items) {
        await applyToSlide(context, slide, options, filter, onLineSpacingUnsupported);
      }
    });
  }
}

/**
 * 적용 전 스타일 스냅샷 캡처 (실행 취소용)
 */
export async function captureSnapshot(target: ApplyTarget): Promise<ShapeStyleRecord[]> {
  const records: ShapeStyleRecord[] = [];

  if (target === 'selection-text' || target === 'selection-shape') {
    return records;
  }

  if (target === 'all-same-position') {
    // 위치 기반 스냅샷: 선택한 도형 위치와 같은 도형들의 스타일 캡처
    await PowerPoint.run(async (context) => {
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load('items/left,items/top');
      await context.sync();

      if (selectedShapes.items.length === 0) return;

      const targetPositions = selectedShapes.items.map((s) => ({
        left: s.left,
        top: s.top,
      }));

      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      for (let si = 0; si < slides.items.length; si++) {
        const shapes = slides.items[si].shapes;
        shapes.load('items/left,items/top');
        await context.sync();

        for (let shi = 0; shi < shapes.items.length; shi++) {
          const shape = shapes.items[shi];
          const matched = targetPositions.some((tp) => isSamePosition(tp, {
            left: shape.left, top: shape.top,
          }));
          if (!matched) continue;

          try {
            const range = shape.textFrame.textRange;
            range.load('font/name,font/size,font/bold,font/italic,font/color');
            range.paragraphFormat.load('horizontalAlignment');
            await context.sync();

            records.push({
              slideIndex: si,
              shapeIndex: shi,
              font: {
                name: range.font.name ?? undefined,
                size: range.font.size ?? undefined,
                bold: range.font.bold ?? undefined,
                italic: range.font.italic ?? undefined,
                underline: undefined,
                color: range.font.color ?? undefined,
              },
              alignment: range.paragraphFormat.horizontalAlignment ?? undefined,
            });
          } catch {
            continue;
          }
        }
      }
    });

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
      shapes.load('items/name');
      await context.sync();

      for (let shi = 0; shi < shapes.items.length; shi++) {
        const shape = shapes.items[shi];

        if (filter === 'title' && !isTitleShape(shape.name)) continue;
        if (filter === 'body' && !isBodyShape(shape.name)) continue;

        try {
          const range = shape.textFrame.textRange;
          range.load('font/name,font/size,font/bold,font/italic,font/color');
          range.paragraphFormat.load('horizontalAlignment');
          await context.sync();

          records.push({
            slideIndex: si,
            shapeIndex: shi,
            font: {
              name: range.font.name ?? undefined,
              size: range.font.size ?? undefined,
              bold: range.font.bold ?? undefined,
              italic: range.font.italic ?? undefined,
              underline: undefined,
              color: range.font.color ?? undefined,
            },
            alignment: range.paragraphFormat.horizontalAlignment ?? undefined,
          });
        } catch {
          continue;
        }
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
      shapes.load('items/name');
      await context.sync();

      const shape = shapes.items[record.shapeIndex];
      if (!shape) continue;

      try {
        const range = shape.textFrame.textRange;
        range.load('text');
        await context.sync();

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
        await context.sync();
      } catch {
        continue;
      }
    }
  });
}
