import { create } from 'zustand';

export type ApplyTarget =
  | 'selection-text'
  | 'selection-shape'
  | 'current-title'
  | 'current-body'
  | 'current-all'
  | 'all-titles'
  | 'all-bodies'
  | 'all-all';

export interface FontStyle {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
}

export interface ParagraphStyle {
  alignment?: 'left' | 'center' | 'right' | 'justify';
  lineSpacing?: number;
}

export interface StylePreset {
  id: string;
  name: string;
  font: FontStyle;
  paragraph?: ParagraphStyle;
}

/** 실행 취소를 위해 적용 전 스냅샷으로 저장하는 개별 shape의 스타일 */
export interface ShapeStyleRecord {
  slideIndex: number;
  shapeIndex: number;
  font: FontStyle;
  alignment?: string;
}

/** undoStack 한 항목: 적용 전 상태 스냅샷 */
export interface UndoEntry {
  timestamp: number;
  description: string;
  shapes: ShapeStyleRecord[];
}

interface StoreState {
  currentFont: FontStyle;
  currentParagraph: ParagraphStyle;
  applyTarget: ApplyTarget;
  presets: StylePreset[];
  /** 최대 10단계 실행 취소 스택 */
  undoStack: UndoEntry[];
  activeTab: 'editor' | 'presets';
  officeVersion: string;
  isLegacyApi: boolean;

  setCurrentFont: (font: Partial<FontStyle>) => void;
  setCurrentParagraph: (para: Partial<ParagraphStyle>) => void;
  setApplyTarget: (target: ApplyTarget) => void;
  setPresets: (presets: StylePreset[]) => void;
  addPreset: (preset: StylePreset) => void;
  updatePreset: (preset: StylePreset) => void;
  deletePreset: (id: string) => void;
  pushUndo: (entry: UndoEntry) => void;
  popUndo: () => UndoEntry | undefined;
  setActiveTab: (tab: 'editor' | 'presets') => void;
  setOfficeInfo: (version: string, isLegacy: boolean) => void;
  loadPresetToEditor: (preset: StylePreset) => void;
}

export const useStore = create<StoreState>((set, get) => ({
  currentFont: {
    name: 'Malgun Gothic',
    size: 18,
    bold: false,
    italic: false,
    underline: false,
    color: '#333333',
  },
  currentParagraph: {
    alignment: 'left',
    lineSpacing: 1.5,
  },
  applyTarget: 'current-all',
  presets: [],
  undoStack: [],
  activeTab: 'editor',
  officeVersion: '',
  isLegacyApi: false,

  setCurrentFont: (font) =>
    set((state) => ({ currentFont: { ...state.currentFont, ...font } })),

  setCurrentParagraph: (para) =>
    set((state) => ({ currentParagraph: { ...state.currentParagraph, ...para } })),

  setApplyTarget: (target) => set({ applyTarget: target }),

  setPresets: (presets) => set({ presets }),

  addPreset: (preset) =>
    set((state) => ({ presets: [...state.presets, preset] })),

  updatePreset: (preset) =>
    set((state) => ({
      presets: state.presets.map((p) => (p.id === preset.id ? preset : p)),
    })),

  deletePreset: (id) =>
    set((state) => ({ presets: state.presets.filter((p) => p.id !== id) })),

  pushUndo: (entry) =>
    set((state) => ({
      undoStack: [entry, ...state.undoStack].slice(0, 10),
    })),

  popUndo: () => {
    const stack = get().undoStack;
    if (stack.length === 0) return undefined;
    const [top, ...rest] = stack;
    set({ undoStack: rest });
    return top;
  },

  setActiveTab: (tab) => set({ activeTab: tab }),

  setOfficeInfo: (version, isLegacy) =>
    set({ officeVersion: version, isLegacyApi: isLegacy }),

  loadPresetToEditor: (preset) =>
    set({
      currentFont: { ...preset.font },
      currentParagraph: preset.paragraph
        ? { ...preset.paragraph }
        : { alignment: 'left', lineSpacing: 1.5 },
      activeTab: 'editor',
    }),
}));
