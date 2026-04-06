import { useState } from 'react';
import {
  Button,
  Card,
  makeStyles,
  tokens,
  Text,
} from '@fluentui/react-components';
import {
  EditRegular,
  DeleteRegular,
  PlayRegular,
  TextHeader1Regular,
  TextAlignLeftRegular,
} from '@fluentui/react-icons';
import type { StylePreset } from '../../store/useStore';
import { useStore } from '../../store/useStore';
import { PresetModal } from './PresetModal';

const useStyles = makeStyles({
  card: {
    marginBottom: '6px',
    padding: '8px 10px',
  },
  topRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    marginBottom: '4px',
  },
  name: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    flex: 1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  detail: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    marginBottom: '2px',
  },
  colorDot: {
    display: 'inline-block',
    width: '8px',
    height: '8px',
    borderRadius: '50%',
    marginRight: '3px',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    verticalAlign: 'middle',
  },
  badges: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: '4px',
  },
  actions: {
    display: 'flex',
    gap: '3px',
    alignItems: 'center',
    flexWrap: 'wrap',
  },
  spacer: {
    flex: 1,
  },
  storageBadge: {
    fontSize: '10px',
    padding: '2px 8px',
    borderRadius: '10px',
    cursor: 'pointer',
    lineHeight: '16px',
    fontWeight: tokens.fontWeightSemibold,
    display: 'inline-flex',
    alignItems: 'center',
    gap: '3px',
    userSelect: 'none',
  },
  storageLocal: {
    backgroundColor: '#dff6dd',
    color: '#107c10',
    border: '1px solid #a7e3a5',
  },
  storageDoc: {
    backgroundColor: '#f0e6f6',
    color: '#8764b8',
    border: '1px solid #c8a8e0',
  },
  storageIcon: {
    fontSize: '11px',
  },
});

interface PresetCardProps {
  preset: StylePreset;
  onApply: (preset: StylePreset) => void;
  onDelete: (id: string) => void;
  onAssignTitle: (id: string) => void;
  onAssignBody: (id: string) => void;
  onAssignSlot: (presetId: string, slot: number) => void;
  onToggleStorage: (presetId: string) => void;
}

export function PresetCard({ preset, onApply, onDelete, onAssignTitle, onAssignBody, onAssignSlot, onToggleStorage }: PresetCardProps) {
  const styles = useStyles();
  const [showEditModal, setShowEditModal] = useState(false);
  const { titlePresetId, bodyPresetId, slotPresetIds } = useStore();

  const assignedSlot = Object.entries(slotPresetIds).find(
    ([, id]) => id === preset.id
  )?.[0] ?? '';

  const isLocal = preset.storage !== 'document';

  const badges: string[] = [];
  if (titlePresetId === preset.id) badges.push('제목');
  if (bodyPresetId === preset.id) badges.push('본문');
  const slotNum = assignedSlot ? Number(assignedSlot) : 0;
  if (slotNum > 0) badges.push(`P${slotNum}`);

  const fontDetails = [
    preset.font.name,
    preset.font.size ? `${preset.font.size}pt` : null,
    preset.font.bold ? 'B' : null,
    preset.font.italic ? 'I' : null,
  ].filter(Boolean).join(' · ');

  return (
    <>
      <Card className={styles.card} size="small">
        <div className={styles.topRow}>
          <span
            className={styles.colorDot}
            style={{ backgroundColor: preset.font.color || '#333', width: 12, height: 12 }}
          />
          <Text className={styles.name}>{preset.name}</Text>
          <span
            className={`${styles.storageBadge} ${isLocal ? styles.storageLocal : styles.storageDoc}`}
            onClick={() => onToggleStorage(preset.id)}
            title={isLocal ? '전역 저장 중 (클릭: 파일별로 변경)' : '파일별 저장 중 (클릭: 전역으로 변경)'}
          >
            <span className={styles.storageIcon}>{isLocal ? '🌐' : '📄'}</span>
            {isLocal ? '전역' : '파일별'}
          </span>
          <Button size="small" icon={<EditRegular />} onClick={() => setShowEditModal(true)} title="수정" />
          <Button size="small" icon={<DeleteRegular />} onClick={() => onDelete(preset.id)} title="삭제" />
        </div>

        <Text className={styles.detail}>{fontDetails}</Text>

        {badges.length > 0 && (
          <Text className={styles.badges}>
            {badges.map((b) => `[${b}]`).join(' ')}
          </Text>
        )}

        <div className={styles.actions}>
          <Button size="small" appearance="primary" icon={<PlayRegular />} onClick={() => onApply(preset)}>
            적용
          </Button>
          <Button
            size="small"
            appearance={titlePresetId === preset.id ? 'primary' : 'subtle'}
            icon={<TextHeader1Regular />}
            onClick={() => onAssignTitle(preset.id)}
          >
            제목
          </Button>
          <Button
            size="small"
            appearance={bodyPresetId === preset.id ? 'primary' : 'subtle'}
            icon={<TextAlignLeftRegular />}
            onClick={() => onAssignBody(preset.id)}
          >
            본문
          </Button>
          <span className={styles.spacer} />
          {[1, 2, 3].map((s) => (
            <Button
              key={s}
              size="small"
              appearance={Number(assignedSlot) === s ? 'primary' : 'outline'}
              onClick={() => onAssignSlot(preset.id, Number(assignedSlot) === s ? 0 : s)}
              title={`프리셋 ${s}`}
            >
              P{s}
            </Button>
          ))}
        </div>
      </Card>

      {showEditModal && (
        <PresetModal
          onClose={() => setShowEditModal(false)}
          initialFont={preset.font}
          initialParagraph={preset.paragraph ?? { alignment: 'left', lineSpacing: 1.5 }}
          existingPreset={preset}
        />
      )}
    </>
  );
}
