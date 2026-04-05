import { useState } from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
  Label,
  makeStyles,
} from '@fluentui/react-components';
import { useStore, type FontStyle, type ParagraphStyle, type StylePreset } from '../../store/useStore';
import { savePresetsToSettings, generateId } from '../../services/presetStorage';

const useStyles = makeStyles({
  field: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    marginBottom: '12px',
  },
});

interface PresetModalProps {
  onClose: () => void;
  initialFont: FontStyle;
  initialParagraph: ParagraphStyle;
  existingPreset?: StylePreset;
}

export function PresetModal({ onClose, initialFont, initialParagraph, existingPreset }: PresetModalProps) {
  const styles = useStyles();
  const [name, setName] = useState(existingPreset?.name ?? '');
  const { addPreset, updatePreset, presets } = useStore();

  async function handleSave() {
    if (!name.trim()) return;

    const preset: StylePreset = {
      id: existingPreset?.id ?? generateId(),
      name: name.trim(),
      font: { ...initialFont },
      paragraph: { ...initialParagraph },
    };

    if (existingPreset) {
      updatePreset(preset);
    } else {
      addPreset(preset);
    }

    // settings에 저장
    const updated = existingPreset
      ? presets.map((p) => (p.id === preset.id ? preset : p))
      : [...presets, preset];
    await savePresetsToSettings(updated).catch(console.error);

    onClose();
  }

  return (
    <Dialog open onOpenChange={(_, d) => { if (!d.open) onClose(); }}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{existingPreset ? '프리셋 수정' : '새 프리셋 저장'}</DialogTitle>
          <DialogContent>
            <div className={styles.field}>
              <Label>프리셋 이름</Label>
              <Input
                value={name}
                onChange={(_, d) => setName(d.value)}
                placeholder="예: 제목 스타일"
                autoFocus
              />
            </div>
            <div style={{ fontSize: '12px', color: '#666' }}>
              <div>폰트: {initialFont.name} {initialFont.size}pt</div>
              <div>색상: {initialFont.color}</div>
              {initialFont.bold && <span>굵게 </span>}
              {initialFont.italic && <span>기울임 </span>}
              {initialFont.underline && <span>밑줄</span>}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>취소</Button>
            <Button appearance="primary" onClick={handleSave} disabled={!name.trim()}>
              저장
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
