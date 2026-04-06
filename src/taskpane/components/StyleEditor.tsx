import { useState, useEffect } from 'react';
import {
  Button,
  Input,
  Label,
  Select,
  ToggleButton,
  Tooltip,
  makeStyles,
  tokens,
  Toast,
  ToastTitle,
  useToastController,
  Toaster,
  useId,
} from '@fluentui/react-components';
import {
  TextBoldRegular,
  TextItalicRegular,
  TextUnderlineRegular,
  TextAlignLeftRegular,
  TextAlignCenterRegular,
  TextAlignRightRegular,
  TextAlignJustifyRegular,
  ArrowUndoRegular,
  SaveRegular,
} from '@fluentui/react-icons';
import { useStore } from '../../store/useStore';
import { ColorPicker } from './ColorPicker';
import { PresetModal } from './PresetModal';
import { applyStyle, captureSnapshot, restoreSnapshot } from '../../services/styleService';
import { hasLineSpacingSupport } from '../../services/officeService';

const useStyles = makeStyles({
  container: {
    padding: '12px',
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    paddingBottom: '4px',
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  label: {
    width: '52px',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    flexShrink: 0,
  },
  sizeInput: {
    width: '70px',
  },
  toggleGroup: {
    display: 'flex',
    gap: '4px',
  },
  alignGroup: {
    display: 'flex',
    gap: '4px',
  },
  lineSpacingInput: {
    width: '70px',
  },
  actions: {
    display: 'flex',
    gap: '8px',
    paddingTop: '4px',
  },
  applyBtn: {
    flex: 1,
  },
  undoRow: {
    display: 'flex',
    justifyContent: 'flex-end',
  },
});

/** 감지할 폰트 후보 목록 (한글 + 영문 + 시스템) */
const FONT_CANDIDATES = [
  // 한글 폰트
  'Malgun Gothic', '맑은 고딕',
  'Pretendard',
  '나눔고딕', 'NanumGothic', '나눔고딕코딩',
  '나눔명조', 'NanumMyeongjo',
  '나눔바른고딕', 'NanumBarunGothic',
  '나눔스퀘어', 'NanumSquare', 'NanumSquareRound',
  '돋움', 'Dotum', '굴림', 'Gulim',
  '바탕', 'Batang', '궁서', 'Gungsuh',
  'Apple SD Gothic Neo',
  'Noto Sans KR', 'Noto Serif KR',
  '본고딕', 'Source Han Sans K',
  '본명조', 'Source Han Serif K',
  'Spoqa Han Sans Neo',
  'IBM Plex Sans KR',
  'KoPubWorldDotum', 'KoPubWorldBatang',
  'Gmarket Sans',
  'Noto Sans CJK KR',
  '함초롬돋움', '함초롬바탕',
  'D2Coding',
  // 영문 폰트
  'Arial', 'Arial Black', 'Arial Narrow',
  'Calibri', 'Calibri Light',
  'Cambria', 'Cambria Math',
  'Times New Roman',
  'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold',
  'Verdana', 'Tahoma', 'Trebuchet MS',
  'Georgia',
  'Helvetica', 'Helvetica Neue',
  'Garamond',
  'Palatino', 'Palatino Linotype', 'Book Antiqua',
  'Century Gothic',
  'Franklin Gothic Medium',
  'Lucida Sans', 'Lucida Console',
  'Consolas', 'Courier New',
  'Impact',
  'Comic Sans MS',
  'Candara', 'Constantia', 'Corbel',
  'Rockwell',
  'Futura',
  'Avenir', 'Avenir Next',
  'Gill Sans', 'Gill Sans MT',
  'Optima',
  'San Francisco', 'SF Pro Display', 'SF Pro Text',
  'Roboto', 'Open Sans', 'Lato', 'Montserrat', 'Poppins',
  'Inter',
  'Aptos', 'Aptos Display',
];

/** document.fonts.check()로 설치된 폰트 감지 */
function detectAvailableFonts(): string[] {
  const available: string[] = [];
  const seen = new Set<string>();

  for (const font of FONT_CANDIDATES) {
    if (seen.has(font.toLowerCase())) continue;
    try {
      if (document.fonts.check(`12px "${font}"`)) {
        available.push(font);
        seen.add(font.toLowerCase());
      }
    } catch {
      // ignore
    }
  }

  return available.length > 0 ? available : ['Arial', 'Calibri', 'Malgun Gothic'];
}

export function StyleEditor() {
  const styles = useStyles();
  const toasterId = useId('toaster');
  const { dispatchToast } = useToastController(toasterId);
  const [showPresetModal, setShowPresetModal] = useState(false);
  const [isApplying, setIsApplying] = useState(false);
  const [isUndoing, setIsUndoing] = useState(false);
  const [availableFonts, setAvailableFonts] = useState<string[]>([]);

  useEffect(() => {
    // 폰트 감지는 document.fonts가 준비된 후 실행
    if (document.fonts?.ready) {
      document.fonts.ready.then(() => {
        setAvailableFonts(detectAvailableFonts());
      });
    } else {
      setAvailableFonts(detectAvailableFonts());
    }
  }, []);

  const {
    currentFont,
    currentParagraph,
    applyTarget,
    setCurrentFont,
    setCurrentParagraph,
    pushUndo,
    popUndo,
    undoStack,
  } = useStore();

  const lineSpacingSupported = hasLineSpacingSupport();

  function showToast(message: string, intent: 'success' | 'warning' | 'error' = 'success') {
    dispatchToast(
      <Toast>
        <ToastTitle>{message}</ToastTitle>
      </Toast>,
      { intent, position: 'bottom' }
    );
  }

  const undoSupported = applyTarget !== 'selection-text' && applyTarget !== 'selection-shape';

  async function handleApply() {
    setIsApplying(true);
    try {
      const shapes = await captureSnapshot(applyTarget);
      pushUndo({
        timestamp: Date.now(),
        description: `스타일 적용: ${applyTarget}`,
        shapes,
      });

      await applyStyle(
        applyTarget,
        { font: currentFont, paragraph: currentParagraph },
        () => showToast('줄간격은 Microsoft 365에서만 지원됩니다', 'warning')
      );
      showToast('스타일이 적용되었습니다');
    } catch (e) {
      showToast(`오류: ${(e as Error).message}`, 'error');
    } finally {
      setIsApplying(false);
    }
  }

  async function handleUndo() {
    const entry = popUndo();
    if (!entry) return;

    if (entry.shapes.length === 0) {
      showToast('선택 범위 적용은 실행 취소가 지원되지 않습니다', 'warning');
      return;
    }

    setIsUndoing(true);
    try {
      await restoreSnapshot(entry.shapes);
      showToast('실행 취소 완료');
    } catch (e) {
      showToast(`실행 취소 실패: ${(e as Error).message}`, 'error');
    } finally {
      setIsUndoing(false);
    }
  }

  return (
    <div className={styles.container}>
      <Toaster toasterId={toasterId} />

      {/* 폰트 섹션 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>폰트</div>

        <div className={styles.row}>
          <Label className={styles.label}>폰트명</Label>
          <Select
            value={currentFont.name ?? 'Malgun Gothic'}
            onChange={(_, d) => setCurrentFont({ name: d.value })}
            size="small"
          >
            {availableFonts.map((f) => (
              <option key={f} value={f}>{f}</option>
            ))}
          </Select>
        </div>

        <div className={styles.row}>
          <Label className={styles.label}>크기</Label>
          <Input
            className={styles.sizeInput}
            type="number"
            value={String(currentFont.size ?? 18)}
            onChange={(_, d) => setCurrentFont({ size: Number(d.value) })}
            size="small"
            min={1}
            max={400}
            contentAfter={<span>pt</span>}
          />
        </div>

        <div className={styles.row}>
          <Label className={styles.label}>스타일</Label>
          <div className={styles.toggleGroup}>
            <ToggleButton
              size="small"
              checked={currentFont.bold}
              onClick={() => setCurrentFont({ bold: !currentFont.bold })}
              icon={<TextBoldRegular />}
              title="굵게"
            />
            <ToggleButton
              size="small"
              checked={currentFont.italic}
              onClick={() => setCurrentFont({ italic: !currentFont.italic })}
              icon={<TextItalicRegular />}
              title="기울임"
            />
            <ToggleButton
              size="small"
              checked={currentFont.underline}
              onClick={() => setCurrentFont({ underline: !currentFont.underline })}
              icon={<TextUnderlineRegular />}
              title="밑줄"
            />
          </div>
        </div>

        <div className={styles.row}>
          <Label className={styles.label}>색상</Label>
          <ColorPicker
            color={currentFont.color ?? '#333333'}
            onChange={(c) => setCurrentFont({ color: c })}
          />
        </div>
      </div>

      {/* 단락 섹션 */}
      <div className={styles.section}>
        <div className={styles.sectionTitle}>단락</div>

        <div className={styles.row}>
          <Label className={styles.label}>정렬</Label>
          <div className={styles.alignGroup}>
            {(['left', 'center', 'right', 'justify'] as const).map((align) => (
              <ToggleButton
                key={align}
                size="small"
                checked={currentParagraph.alignment === align}
                onClick={() => setCurrentParagraph({ alignment: align })}
                icon={
                  align === 'left' ? <TextAlignLeftRegular /> :
                  align === 'center' ? <TextAlignCenterRegular /> :
                  align === 'right' ? <TextAlignRightRegular /> :
                  <TextAlignJustifyRegular />
                }
                title={
                  align === 'left' ? '왼쪽' :
                  align === 'center' ? '가운데' :
                  align === 'right' ? '오른쪽' : '양쪽'
                }
              />
            ))}
          </div>
        </div>

        <div className={styles.row}>
          <Label className={styles.label}>줄간격</Label>
          <Tooltip
            content={lineSpacingSupported ? '' : '이 기능은 Microsoft 365에서만 지원됩니다'}
            relationship="label"
          >
            <Input
              className={styles.lineSpacingInput}
              type="number"
              value={String(currentParagraph.lineSpacing ?? 1.5)}
              onChange={(_, d) => setCurrentParagraph({ lineSpacing: Number(d.value) })}
              size="small"
              min={0.5}
              max={5}
              step={0.1}
              disabled={!lineSpacingSupported}
              contentAfter={<span>배</span>}
            />
          </Tooltip>
        </div>
      </div>

      {/* 실행 취소 */}
      <div className={styles.undoRow}>
        <Button
          size="small"
          icon={<ArrowUndoRegular />}
          disabled={undoStack.length === 0 || isUndoing}
          onClick={handleUndo}
        >
          {isUndoing ? '복원 중...' : `실행취소 (${undoStack.length})`}
        </Button>
      </div>

      {/* 적용 버튼 */}
      <div className={styles.actions}>
        <Button
          className={styles.applyBtn}
          appearance="primary"
          onClick={handleApply}
          disabled={isApplying}
        >
          {isApplying ? '적용 중...' : '현재 선택에 적용'}
        </Button>
        <Button
          icon={<SaveRegular />}
          onClick={() => setShowPresetModal(true)}
          title="프리셋으로 저장"
        >
          프리셋 저장
        </Button>
      </div>

      {showPresetModal && (
        <PresetModal
          onClose={() => setShowPresetModal(false)}
          initialFont={currentFont}
          initialParagraph={currentParagraph}
        />
      )}
    </div>
  );
}
