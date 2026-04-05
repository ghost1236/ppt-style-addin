import {
  Button,
  Text,
  makeStyles,
  tokens,
  Toast,
  ToastTitle,
  useToastController,
  Toaster,
  useId,
} from '@fluentui/react-components';
import {
  AddRegular,
  ArrowExportRegular,
  ArrowImportRegular,
} from '@fluentui/react-icons';
import { useStore } from '../../store/useStore';
import { PresetCard } from './PresetCard';
import { PresetModal } from './PresetModal';
import {
  savePresetsToSettings,
  exportPresetsAsJson,
  importPresetsFromJson,
} from '../../services/presetStorage';
import { applyStyle } from '../../services/styleService';
import { useState } from 'react';

const useStyles = makeStyles({
  container: {
    padding: '12px',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  emptyText: {
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
    padding: '24px 0',
  },
  bottomActions: {
    display: 'flex',
    gap: '4px',
    paddingTop: '8px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    flexWrap: 'wrap',
  },
});

export function PresetList() {
  const styles = useStyles();
  const toasterId = useId('preset-toaster');
  const { dispatchToast } = useToastController(toasterId);
  const [showNewModal, setShowNewModal] = useState(false);
  const { presets, deletePreset, setPresets, applyTarget, currentFont, currentParagraph, loadPresetToEditor } = useStore();

  function showToast(msg: string, intent: 'success' | 'warning' | 'error' = 'success') {
    dispatchToast(
      <Toast><ToastTitle>{msg}</ToastTitle></Toast>,
      { intent, position: 'bottom' }
    );
  }

  async function handleApply(preset: typeof presets[number]) {
    try {
      await applyStyle(
        applyTarget,
        { font: preset.font, paragraph: preset.paragraph },
        () => showToast('줄간격은 Microsoft 365에서만 지원됩니다', 'warning')
      );
      showToast(`"${preset.name}" 적용 완료`);
      loadPresetToEditor(preset);
    } catch (e) {
      showToast(`오류: ${(e as Error).message}`, 'error');
    }
  }

  async function handleDelete(id: string) {
    deletePreset(id);
    const updated = presets.filter((p) => p.id !== id);
    await savePresetsToSettings(updated).catch(console.error);
    showToast('프리셋 삭제됨');
  }

  async function handleImport() {
    try {
      const imported = await importPresetsFromJson();
      setPresets(imported);
      await savePresetsToSettings(imported);
      showToast(`${imported.length}개 프리셋 불러옴`);
    } catch (e) {
      showToast(`불러오기 실패: ${(e as Error).message}`, 'error');
    }
  }

  return (
    <div className={styles.container}>
      <Toaster toasterId={toasterId} />

      {presets.length === 0 ? (
        <Text className={styles.emptyText}>
          저장된 프리셋이 없습니다.{'\n'}
          스타일 편집 탭에서 프리셋을 저장해보세요.
        </Text>
      ) : (
        presets.map((preset) => (
          <PresetCard
            key={preset.id}
            preset={preset}
            onApply={handleApply}
            onDelete={handleDelete}
          />
        ))
      )}

      <div className={styles.bottomActions}>
        <Button
          icon={<AddRegular />}
          size="small"
          onClick={() => setShowNewModal(true)}
        >
          새 프리셋
        </Button>
        <Button
          icon={<ArrowExportRegular />}
          size="small"
          onClick={() => exportPresetsAsJson(presets)}
          disabled={presets.length === 0}
        >
          JSON 내보내기
        </Button>
        <Button
          icon={<ArrowImportRegular />}
          size="small"
          onClick={handleImport}
        >
          불러오기
        </Button>
      </div>

      {showNewModal && (
        <PresetModal
          onClose={() => setShowNewModal(false)}
          initialFont={currentFont}
          initialParagraph={currentParagraph}
        />
      )}
    </div>
  );
}
