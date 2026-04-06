/* global Office, PowerPoint */

const PRESETS_KEY = 'ppt-style-addin-presets';
const SLOT_KEY_PREFIX = 'ppt-style-addin-slot-';

Office.onReady(() => {
  // 리본 버튼 함수 등록
});

/** settings에서 프리셋 목록과 특정 슬롯에 지정된 프리셋을 가져옴 */
function getPresetForSlot(slotIndex: number): Record<string, unknown> | null {
  const presetsRaw = Office.context.document.settings.get(PRESETS_KEY);
  const presets = presetsRaw ? JSON.parse(presetsRaw) : [];
  const slotPresetId = Office.context.document.settings.get(`${SLOT_KEY_PREFIX}${slotIndex}`);
  if (slotPresetId) {
    return presets.find((p: { id: string }) => p.id === slotPresetId) ?? null;
  }
  return null;
}

/** 선택한 도형에 프리셋 스타일 적용 */
async function applyPresetToSelection(preset: Record<string, unknown>): Promise<void> {
  const font = preset.font as Record<string, unknown>;
  if (!font) return;

  await PowerPoint.run(async (context) => {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load('items');
    await context.sync();

    for (const shape of selectedShapes.items) {
      try {
        const textRange = shape.textFrame.textRange;
        textRange.load('text');
        await context.sync();
        applyFont(textRange, font);
        await context.sync();
      } catch {
        continue;
      }
    }
  });
}

/** 프리셋 슬롯 핸들러 생성 */
function createSlotHandler(slotIndex: number) {
  return async (event: Office.AddinCommands.Event) => {
    try {
      const preset = getPresetForSlot(slotIndex);
      if (!preset) {
        console.warn(`슬롯 ${slotIndex}에 지정된 프리셋이 없습니다.`);
        return;
      }
      await applyPresetToSelection(preset);
    } catch (e) {
      console.error(`applyPresetSlot${slotIndex} error:`, e);
    } finally {
      event.completed();
    }
  };
}

/**
 * 리본 "제목 일괄 적용" 버튼 핸들러
 */
async function applyToAllTitles(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const presetsRaw = Office.context.document.settings.get(PRESETS_KEY);
      const presets = presetsRaw ? JSON.parse(presetsRaw) : [];
      const titlePresetId = Office.context.document.settings.get('ppt-style-addin-title-preset-id');
      const preset = titlePresetId
        ? presets.find((p: { id: string }) => p.id === titlePresetId)
        : presets[0];
      if (!preset) return;

      for (const slide of slides.items) {
        const shapes = slide.shapes;
        shapes.load('items/name');
        await context.sync();

        for (const shape of shapes.items) {
          if (!isTitleShapeName(shape.name)) continue;
          try {
            const textRange = shape.textFrame.textRange;
            textRange.load('text');
            await context.sync();
            applyFont(textRange, preset.font);
            await context.sync();
          } catch {
            continue;
          }
        }
      }
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

      const presetsRaw = Office.context.document.settings.get(PRESETS_KEY);
      const presets = presetsRaw ? JSON.parse(presetsRaw) : [];
      const bodyPresetId = Office.context.document.settings.get('ppt-style-addin-body-preset-id');
      const preset = bodyPresetId
        ? presets.find((p: { id: string }) => p.id === bodyPresetId)
        : presets[0];
      if (!preset) return;

      for (const slide of slides.items) {
        const shapes = slide.shapes;
        shapes.load('items/name');
        await context.sync();

        for (const shape of shapes.items) {
          if (!isBodyShapeName(shape.name)) continue;
          try {
            const textRange = shape.textFrame.textRange;
            textRange.load('text');
            await context.sync();
            applyFont(textRange, preset.font);
            await context.sync();
          } catch {
            continue;
          }
        }
      }
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
Office.actions.associate('applyPresetSlot1', createSlotHandler(1));
Office.actions.associate('applyPresetSlot2', createSlotHandler(2));
Office.actions.associate('applyPresetSlot3', createSlotHandler(3));
Office.actions.associate('applyPresetSlot4', createSlotHandler(4));
Office.actions.associate('applyPresetSlot5', createSlotHandler(5));
