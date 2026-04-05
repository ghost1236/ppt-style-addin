import { useState } from 'react';
import {
  Button,
  Card,
  CardHeader,
  makeStyles,
  tokens,
  Text,
} from '@fluentui/react-components';
import {
  EditRegular,
  DeleteRegular,
  PlayRegular,
} from '@fluentui/react-icons';
import type { StylePreset } from '../../store/useStore';
import { PresetModal } from './PresetModal';

const useStyles = makeStyles({
  card: {
    marginBottom: '8px',
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    width: '100%',
  },
  info: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    flex: 1,
  },
  name: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
  },
  detail: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
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
  actions: {
    display: 'flex',
    gap: '4px',
    alignItems: 'center',
    flexShrink: 0,
  },
});

interface PresetCardProps {
  preset: StylePreset;
  onApply: (preset: StylePreset) => void;
  onDelete: (id: string) => void;
}

export function PresetCard({ preset, onApply, onDelete }: PresetCardProps) {
  const styles = useStyles();
  const [showEditModal, setShowEditModal] = useState(false);

  const fontDetails = [
    preset.font.name,
    preset.font.size ? `${preset.font.size}pt` : null,
    preset.font.bold ? 'Bold' : null,
    preset.font.italic ? 'Italic' : null,
  ]
    .filter(Boolean)
    .join(' · ');

  return (
    <>
      <Card className={styles.card} size="small">
        <CardHeader
          header={
            <div className={styles.header}>
              <div className={styles.info}>
                <Text className={styles.name}>🎨 {preset.name}</Text>
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
              </div>
              <div className={styles.actions}>
                <Button
                  size="small"
                  appearance="primary"
                  icon={<PlayRegular />}
                  onClick={() => onApply(preset)}
                  title="적용"
                >
                  적용
                </Button>
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
            </div>
          }
        />
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
