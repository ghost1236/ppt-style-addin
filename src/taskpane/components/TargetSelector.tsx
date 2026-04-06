import {
  Select,
  makeStyles,
  tokens,
  Text,
} from '@fluentui/react-components';
import { useStore, type ApplyTarget } from '../../store/useStore';

const useStyles = makeStyles({
  section: {
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightBold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase' as const,
    letterSpacing: '0.8px',
  },
});

const TARGETS: { value: ApplyTarget; label: string }[] = [
  { value: 'all-all', label: '전체 슬라이드 - 모든 텍스트' },
  { value: 'all-titles', label: '전체 슬라이드 - 제목만' },
  { value: 'all-bodies', label: '전체 슬라이드 - 본문만' },
  { value: 'all-same-position', label: '같은 위치 → 전체 슬라이드' },
  { value: 'selection-shape', label: '선택한 텍스트 상자' },
  { value: 'selection-text', label: '선택한 텍스트만' },
];

export function TargetSelector() {
  const styles = useStyles();
  const { applyTarget, setApplyTarget } = useStore();

  return (
    <div className={styles.section}>
      <Text className={styles.label}>적용 대상</Text>
      <Select
        value={applyTarget}
        onChange={(_, d) => setApplyTarget(d.value as ApplyTarget)}
        size="medium"
      >
        {TARGETS.map((t) => (
          <option key={t.value} value={t.value}>{t.label}</option>
        ))}
      </Select>
    </div>
  );
}
