import {
  Radio,
  RadioGroup,
  Label,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useStore, type ApplyTarget } from '../../store/useStore';

const useStyles = makeStyles({
  section: {
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    marginBottom: '6px',
    display: 'block',
  },
  radioGroup: {
    gap: '2px',
  },
  groupLabel: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    marginTop: '6px',
    marginBottom: '2px',
    display: 'block',
    paddingLeft: '4px',
  },
});

const TARGETS: { value: ApplyTarget; label: string }[] = [
  { value: 'selection-text', label: '선택한 텍스트만' },
  { value: 'selection-shape', label: '선택한 텍스트 상자' },
  { value: 'current-title', label: '현재 슬라이드 - 제목' },
  { value: 'current-body', label: '현재 슬라이드 - 본문' },
  { value: 'current-all', label: '현재 슬라이드 - 전체' },
  { value: 'all-titles', label: '모든 슬라이드 - 제목' },
  { value: 'all-bodies', label: '모든 슬라이드 - 본문' },
  { value: 'all-all', label: '모든 슬라이드 - 전체' },
];

export function TargetSelector() {
  const styles = useStyles();
  const { applyTarget, setApplyTarget } = useStore();

  return (
    <div className={styles.section}>
      <Label className={styles.label}>적용 대상</Label>
      <RadioGroup
        value={applyTarget}
        onChange={(_, d) => setApplyTarget(d.value as ApplyTarget)}
        className={styles.radioGroup}
      >
        <span className={styles.groupLabel}>선택 영역</span>
        {TARGETS.slice(0, 2).map((t) => (
          <Radio key={t.value} value={t.value} label={t.label} />
        ))}
        <span className={styles.groupLabel}>현재 슬라이드</span>
        {TARGETS.slice(2, 5).map((t) => (
          <Radio key={t.value} value={t.value} label={t.label} />
        ))}
        <span className={styles.groupLabel}>전체 슬라이드</span>
        {TARGETS.slice(5).map((t) => (
          <Radio key={t.value} value={t.value} label={t.label} />
        ))}
      </RadioGroup>
    </div>
  );
}
