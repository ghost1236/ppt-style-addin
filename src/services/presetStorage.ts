import type { StylePreset } from '../store/useStore';

const STORAGE_KEY = 'ppt-style-addin-presets';
const TITLE_PRESET_KEY = 'ppt-style-addin-title-preset-id';
const BODY_PRESET_KEY = 'ppt-style-addin-body-preset-id';
const SLOT_KEY_PREFIX = 'ppt-style-addin-slot-';
const MODE_KEY = 'ppt-style-addin-storage-mode';

/** 현재 저장 모드 */
function getMode(): 'local' | 'document' {
  return (localStorage.getItem(MODE_KEY) as 'local' | 'document') || 'local';
}

// ── localStorage 방식 ──

function localSet(key: string, value: string) {
  localStorage.setItem(key, value);
}

function localGet(key: string): string | null {
  return localStorage.getItem(key);
}

// ── Office document.settings 방식 ──

function docSet(key: string, value: string) {
  try {
    Office.context.document.settings.set(key, value);
    Office.context.document.settings.saveAsync(() => {});
  } catch { /* ignore */ }
}

function docGet(key: string): string | null {
  try {
    return Office.context.document.settings.get(key) ?? null;
  } catch {
    return null;
  }
}

// ── 통합 인터페이스 ──

function storageSet(key: string, value: string) {
  if (getMode() === 'document') {
    docSet(key, value);
  } else {
    localSet(key, value);
  }
}

function storageGet(key: string): string | null {
  if (getMode() === 'document') {
    return docGet(key);
  } else {
    return localGet(key);
  }
}

/** 프리셋 저장 */
export async function savePresetsToSettings(presets: StylePreset[]): Promise<void> {
  storageSet(STORAGE_KEY, JSON.stringify(presets));
}

/** 프리셋 불러오기 */
export function loadPresetsFromSettings(): StylePreset[] {
  try {
    const raw = storageGet(STORAGE_KEY);
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
  storageSet(TITLE_PRESET_KEY, titlePresetId ?? '');
  storageSet(BODY_PRESET_KEY, bodyPresetId ?? '');
}

/** 리본 버튼용 프리셋 ID 불러오기 */
export function loadRibbonPresetIds(): { titlePresetId: string | null; bodyPresetId: string | null } {
  return {
    titlePresetId: storageGet(TITLE_PRESET_KEY) || null,
    bodyPresetId: storageGet(BODY_PRESET_KEY) || null,
  };
}

/** 슬롯에 프리셋 ID 저장 */
export async function saveSlotPresetId(slotIndex: number, presetId: string | null): Promise<void> {
  storageSet(`${SLOT_KEY_PREFIX}${slotIndex}`, presetId ?? '');
}

/** 슬롯 프리셋 ID 불러오기 */
export function loadSlotPresetIds(): Record<number, string | null> {
  const slots: Record<number, string | null> = {};
  for (let i = 1; i <= 5; i++) {
    slots[i] = storageGet(`${SLOT_KEY_PREFIX}${i}`) || null;
  }
  return slots;
}

/** 고유 ID 생성 */
export function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}
