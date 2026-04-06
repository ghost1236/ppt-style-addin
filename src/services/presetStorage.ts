import type { StylePreset } from '../store/useStore';

const STORAGE_KEY = 'ppt-style-addin-presets';
const TITLE_PRESET_KEY = 'ppt-style-addin-title-preset-id';
const BODY_PRESET_KEY = 'ppt-style-addin-body-preset-id';
const SLOT_KEY_PREFIX = 'ppt-style-addin-slot-';

/** Office.context.document.settings 에 프리셋 저장 */
export async function savePresetsToSettings(presets: StylePreset[]): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.settings.set(STORAGE_KEY, JSON.stringify(presets));
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message ?? '저장 실패'));
        }
      });
    } catch (e) {
      reject(e);
    }
  });
}

/** Office.context.document.settings 에서 프리셋 불러오기 */
export function loadPresetsFromSettings(): StylePreset[] {
  try {
    const raw = Office.context.document.settings.get(STORAGE_KEY);
    if (!raw) return [];
    return JSON.parse(raw) as StylePreset[];
  } catch {
    return [];
  }
}

/** JSON 파일로 내보내기 */
export function exportPresetsAsJson(presets: StylePreset[]): void {
  const json = JSON.stringify(presets, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'ppt-style-presets.json';
  a.click();
  URL.revokeObjectURL(url);
}

/** JSON 파일에서 불러오기 */
export function importPresetsFromJson(): Promise<StylePreset[]> {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return reject(new Error('파일 없음'));
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const data = JSON.parse(ev.target?.result as string) as StylePreset[];
          resolve(data);
        } catch {
          reject(new Error('유효하지 않은 JSON 파일입니다'));
        }
      };
      reader.readAsText(file);
    };
    input.click();
  });
}

/** 리본 버튼용 프리셋 ID 저장 */
export async function saveRibbonPresetIds(
  titlePresetId: string | null,
  bodyPresetId: string | null
): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.settings.set(TITLE_PRESET_KEY, titlePresetId);
      Office.context.document.settings.set(BODY_PRESET_KEY, bodyPresetId);
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message ?? '저장 실패'));
        }
      });
    } catch (e) {
      reject(e);
    }
  });
}

/** 리본 버튼용 프리셋 ID 불러오기 */
export function loadRibbonPresetIds(): { titlePresetId: string | null; bodyPresetId: string | null } {
  try {
    const titlePresetId = Office.context.document.settings.get(TITLE_PRESET_KEY) ?? null;
    const bodyPresetId = Office.context.document.settings.get(BODY_PRESET_KEY) ?? null;
    return { titlePresetId, bodyPresetId };
  } catch {
    return { titlePresetId: null, bodyPresetId: null };
  }
}

/** 슬롯에 프리셋 ID 저장 */
export async function saveSlotPresetId(slotIndex: number, presetId: string | null): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.settings.set(`${SLOT_KEY_PREFIX}${slotIndex}`, presetId);
      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message ?? '저장 실패'));
        }
      });
    } catch (e) {
      reject(e);
    }
  });
}

/** 슬롯 프리셋 ID 불러오기 */
export function loadSlotPresetIds(): Record<number, string | null> {
  const slots: Record<number, string | null> = {};
  try {
    for (let i = 1; i <= 5; i++) {
      slots[i] = Office.context.document.settings.get(`${SLOT_KEY_PREFIX}${i}`) ?? null;
    }
  } catch {
    // ignore
  }
  return slots;
}

/** 고유 ID 생성 */
export function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}
