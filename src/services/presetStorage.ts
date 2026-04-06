import type { StylePreset } from '../store/useStore';

const LOCAL_PRESETS_KEY = 'ppt-style-addin-presets';
const DOC_PRESETS_KEY = 'ppt-style-addin-doc-presets';
const TITLE_PRESET_KEY = 'ppt-style-addin-title-preset-id';
const BODY_PRESET_KEY = 'ppt-style-addin-body-preset-id';
const SLOT_KEY_PREFIX = 'ppt-style-addin-slot-';

// ── localStorage (전역) ──

function localGet(key: string): string | null {
  try { return localStorage.getItem(key); } catch { return null; }
}
function localSet(key: string, value: string) {
  try { localStorage.setItem(key, value); } catch { /* ignore */ }
}

// ── document.settings (파일별) ──

function docGet(key: string): string | null {
  try { return Office.context.document.settings.get(key) ?? null; } catch { return null; }
}
function docSet(key: string, value: string) {
  try {
    Office.context.document.settings.set(key, value);
    Office.context.document.settings.saveAsync(() => {});
  } catch { /* ignore */ }
}

// ── 프리셋 저장/불러오기 (전역 + 파일별 합산) ──

/** 전역 프리셋만 저장 */
function saveLocalPresets(presets: StylePreset[]) {
  const localPresets = presets.filter((p) => p.storage !== 'document');
  localSet(LOCAL_PRESETS_KEY, JSON.stringify(localPresets));
}

/** 파일별 프리셋만 저장 */
function saveDocPresets(presets: StylePreset[]) {
  const docPresets = presets.filter((p) => p.storage === 'document');
  docSet(DOC_PRESETS_KEY, JSON.stringify(docPresets));
}

/** 프리셋 저장 (전역/파일별 분리 저장) */
export async function savePresetsToSettings(presets: StylePreset[]): Promise<void> {
  saveLocalPresets(presets);
  saveDocPresets(presets);
}

/** 프리셋 불러오기 (전역 + 파일별 합산) */
export function loadPresetsFromSettings(): StylePreset[] {
  const localPresets: StylePreset[] = [];
  const docPresets: StylePreset[] = [];

  try {
    const localRaw = localGet(LOCAL_PRESETS_KEY);
    if (localRaw) {
      const parsed = JSON.parse(localRaw) as StylePreset[];
      parsed.forEach((p) => { p.storage = 'local'; });
      localPresets.push(...parsed);
    }
  } catch { /* ignore */ }

  try {
    const docRaw = docGet(DOC_PRESETS_KEY);
    if (docRaw) {
      const parsed = JSON.parse(docRaw) as StylePreset[];
      parsed.forEach((p) => { p.storage = 'document'; });
      docPresets.push(...parsed);
    }
  } catch { /* ignore */ }

  // 중복 ID 제거 (파일별이 우선)
  const ids = new Set(docPresets.map((p) => p.id));
  const merged = [...docPresets, ...localPresets.filter((p) => !ids.has(p.id))];
  return merged;
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
  localSet(TITLE_PRESET_KEY, titlePresetId ?? '');
  localSet(BODY_PRESET_KEY, bodyPresetId ?? '');
}

/** 리본 버튼용 프리셋 ID 불러오기 */
export function loadRibbonPresetIds(): { titlePresetId: string | null; bodyPresetId: string | null } {
  return {
    titlePresetId: localGet(TITLE_PRESET_KEY) || null,
    bodyPresetId: localGet(BODY_PRESET_KEY) || null,
  };
}

/** 슬롯에 프리셋 ID 저장 */
export async function saveSlotPresetId(slotIndex: number, presetId: string | null): Promise<void> {
  localSet(`${SLOT_KEY_PREFIX}${slotIndex}`, presetId ?? '');
}

/** 슬롯 프리셋 ID 불러오기 */
export function loadSlotPresetIds(): Record<number, string | null> {
  const slots: Record<number, string | null> = {};
  for (let i = 1; i <= 5; i++) {
    slots[i] = localGet(`${SLOT_KEY_PREFIX}${i}`) || null;
  }
  return slots;
}

/** 고유 ID 생성 */
export function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}
