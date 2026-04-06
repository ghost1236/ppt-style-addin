import { useEffect } from 'react';
import {
  Tab,
  TabList,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
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
    backgroundColor: tokens.colorNeutralBackground1,
  },
  tabs: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  content: {
    flex: 1,
    overflowY: 'auto',
  },
});

export default function App() {
  const styles = useStyles();
  const { activeTab, setActiveTab, setPresets, setOfficeInfo, setTitlePresetId, setBodyPresetId, setSlotPresetIds } =
    useStore();

  useEffect(() => {
    try {
      const version = getOfficeVersion();
      const legacy = !isApiSupported('PowerPointApi', '1.3');
      setOfficeInfo(version, legacy);

      const saved = loadPresetsFromSettings();
      if (saved.length > 0) setPresets(saved);

      const { titlePresetId, bodyPresetId } = loadRibbonPresetIds();
      setTitlePresetId(titlePresetId);
      setBodyPresetId(bodyPresetId);

      const slots = loadSlotPresetIds();
      setSlotPresetIds(slots);
    } catch (e) {
      console.warn('Office 초기화 오류:', e);
    }
  }, []);

  return (
    <div className={styles.root}>
      {/* 탭 */}
      <div className={styles.tabs}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, d) => setActiveTab(d.value as 'editor' | 'presets')}
          size="small"
          appearance="subtle"
        >
          <Tab value="editor">스타일</Tab>
          <Tab value="presets">프리셋</Tab>
        </TabList>
      </div>

      {/* 적용 대상 */}
      <TargetSelector />

      {/* 콘텐츠 */}
      <div className={styles.content}>
        {activeTab === 'editor' ? <StyleEditor /> : <PresetList />}
      </div>
    </div>
  );
}
