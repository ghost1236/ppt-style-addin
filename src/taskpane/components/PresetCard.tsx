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
    marginBottom: '8px',
    padding: '8px',
  },
  topRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginBottom: '6px',
  },
  name: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    flex: 1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  detail: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    marginBottom: '4px',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  colorDot: {
    display: 'inline-block',
    width: '10px',
    height: '10px',
    borderRadius: '50%',
    marginRight: '4px',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    verticalAlign: 'middle',
  },
  badges: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: '6px',
  },
  actions: {
    display: 'flex',
    gap: '4px',
    alignItems: 'center',
    flexWrap: 'wrap',
  },
  spacer: {
    flex: 1,
  },
});

interface PresetCardProps {
  preset: StylePreset;
  onApply: (preset: StylePreset) => void;
  onDelete: (id: string) => void;
  onAssignTitle: (id: string) => void;
  onAssignBody: (id: string) => void;
  onAssignSlot: (presetId: string, slot: number) => void;
}

export function PresetCard({ preset, onApply, onDelete, onAssignTitle, onAssignBody, onAssignSlot }: PresetCardProps) {
  const styles = useStyles();
  const [showEditModal, setShowEditModal] = useState(false);
  const { titlePresetId, bodyPresetId, slotPresetIds } = useStore();

  const assignedSlot = Object.entries(slotPresetIds).find(
    ([, id]) => id === preset.id
  )?.[0] ?? '';

  const badges: string[] = [];
  if (titlePresetId === preset.id) badges.push('제목용');
  if (bodyPresetId === preset.id) badges.push('본문용');
  const slotNum = assignedSlot ? Number(assignedSlot) : 0;
  if (slotNum > 0) badges.push(`프리셋 ${slotNum}`);

  const fontDetails = [
    preset.font.name,
    preset.font.size ? `${preset.font.size}pt` : null,
    preset.font.bold ? 'B' : null,
    preset.font.italic ? 'I' : null,
  ]
    .filter(Boolean)
    .join(' · ');

  return (
    <>
      <Card className={styles.card} size="small">
        <div className={styles.topRow}>
          <Text className={styles.name}>🎨 {preset.name}</Text>
          <Button
            size="small"
            icon={<EditRegular />}
            onClick={() => setShowEditModal(true)}
            title="수정"
          />
          <Button
            size="small"
            icon={<DeleteRegular />}
            onClick={() => onDelete(preset.id)}
            title="삭제"
          />
        </div>

        <Text className={styles.detail}>
          {fontDetails}
          {preset.font.color && (
            <>
              {' · '}
              <span
                className={styles.colorDot}
                style={{ backgroundColor: preset.font.color }}
              />
              {preset.font.color}
            </>
          )}
        </Text>

        {badges.length > 0 && (
          <Text className={styles.badges}>
            {badges.map((b) => `[ ${b} ]`).join(' ')}
          </Text>
        )}

        <div className={styles.actions}>
          <Button
            size="small"
            appearance="primary"
            icon={<PlayRegular />}
            onClick={() => onApply(preset)}
          >
            적용
          </Button>
          <Button
            size="small"
            appearance={titlePresetId === preset.id ? 'primary' : 'secondary'}
            icon={<TextHeader1Regular />}
            onClick={() => onAssignTitle(preset.id)}
            title="제목용"
          >
            제목
          </Button>
          <Button
            size="small"
            appearance={bodyPresetId === preset.id ? 'primary' : 'secondary'}
            icon={<TextAlignLeftRegular />}
            onClick={() => onAssignBody(preset.id)}
            title="본문용"
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
              title={`프리셋 ${s} 지정`}
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
