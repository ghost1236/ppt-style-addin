import { useEffect } from 'react';
import {
  Tab,
  TabList,
  makeStyles,
  tokens,
  Badge,
  Text,
  Tooltip,
} from '@fluentui/react-components';
import { InfoRegular } from '@fluentui/react-icons';
import { useStore } from '../../store/useStore';
import { TargetSelector } from './TargetSelector';
import { StyleEditor } from './StyleEditor';
import { PresetList } from './PresetList';
import { loadPresetsFromSettings, loadRibbonPresetIds, loadSlotPresetIds } from '../../services/presetStorage';
import { getOfficeVersion, isApiSupported } from '../../services/officeService';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    overflow: 'hidden',
    fontFamily: tokens.fontFamilyBase,
  },
  header: {
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    backgroundColor: tokens.colorBrandBackground2,
  },
  headerTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorBrandForeground2,
  },
  tabs: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  content: {
    flex: 1,
    overflowY: 'auto',
  },
  footer: {
    padding: '6px 12px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
  },
  badge: {
    flexShrink: 0,
  },
});

export default function App() {
  const styles = useStyles();
  const { activeTab, setActiveTab, setPresets, setOfficeInfo, officeVersion, isLegacyApi, setTitlePresetId, setBodyPresetId, setSlotPresetIds } =
    useStore();

  useEffect(() => {
    try {
      // Office 환경 감지
      const version = getOfficeVersion();
      const legacy = !isApiSupported('PowerPointApi', '1.3');
      setOfficeInfo(version, legacy);

      // 저장된 프리셋 불러오기
      const saved = loadPresetsFromSettings();
      if (saved.length > 0) setPresets(saved);

      // 리본 버튼용 지정 프리셋 불러오기
      const { titlePresetId, bodyPresetId } = loadRibbonPresetIds();
      setTitlePresetId(titlePresetId);
      setBodyPresetId(bodyPresetId);

      // 슬롯 프리셋 불러오기
      const slots = loadSlotPresetIds();
      setSlotPresetIds(slots);
    } catch (e) {
      console.warn('Office 초기화 오류:', e);
    }
  }, []);

  return (
    <div className={styles.root}>
      {/* 헤더 */}
      <div className={styles.header}>
        <Text className={styles.headerTitle}>디자인 도구</Text>
        {isLegacyApi && (
          <Tooltip content="일부 기능이 현재 환경에서 제한됩니다" relationship="label">
            <Badge
              className={styles.badge}
              appearance="outline"
              color="warning"
              icon={<InfoRegular />}
            >
              제한 모드
            </Badge>
          </Tooltip>
        )}
      </div>

      {/* 적용 대상 선택 */}
      <TargetSelector />

      {/* 탭 */}
      <div className={styles.tabs}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, d) => setActiveTab(d.value as 'editor' | 'presets')}
          size="small"
        >
          <Tab value="editor">스타일 편집</Tab>
          <Tab value="presets">저장된 프리셋</Tab>
        </TabList>
      </div>

      {/* 탭 콘텐츠 */}
      <div className={styles.content}>
        {activeTab === 'editor' ? <StyleEditor /> : <PresetList />}
      </div>

      {/* 하단 환경 정보 */}
      <div className={styles.footer}>
        <InfoRegular fontSize={12} />
        <span>
          {isLegacyApi
            ? `감지된 환경: PowerPoint 영구 라이선스 (일부 기능 제한)`
            : `감지된 환경: Microsoft 365 (전체 기능 지원)`}
        </span>
      </div>
    </div>
  );
}
