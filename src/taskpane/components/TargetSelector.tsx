import {
  Select,
  Label,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useStore, type ApplyTarget } from '../../store/useStore';

const useStyles = makeStyles({
  section: {
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
});

const TARGETS: { value: ApplyTarget; label: string }[] = [
  { value: 'selection-text', label: '선택한 텍스트만' },
  { value: 'selection-shape', label: '선택한 텍스트 상자' },
  { value: 'all-same-position', label: '같은 위치 → 전체 슬라이드' },
  { value: 'all-titles', label: '전체 슬라이드 - 제목' },
  { value: 'all-bodies', label: '전체 슬라이드 - 본문' },
  { value: 'all-all', label: '전체 슬라이드 - 모든 텍스트' },
];

export function TargetSelector() {
  const styles = useStyles();
  const { applyTarget, setApplyTarget } = useStore();

  return (
    <div className={styles.section}>
      <Label className={styles.label}>적용 대상</Label>
      <Select
        value={applyTarget}
        onChange={(_, d) => setApplyTarget(d.value as ApplyTarget)}
        size="small"
      >
        {TARGETS.map((t) => (
          <option key={t.value} value={t.value}>{t.label}</option>
        ))}
      </Select>
    </div>
  );
}
