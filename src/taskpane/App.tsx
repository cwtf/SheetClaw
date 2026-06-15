import { useState } from 'react';
import {
  Body1Strong,
  Caption1,
  FluentProvider,
  Tab,
  TabList,
  webLightTheme,
  type SelectTabData,
} from '@fluentui/react-components';
import ChatPanel from './components/ChatPanel';
import HistoryPanel from './components/HistoryPanel';
import UsageDashboard from './components/UsageDashboard';
import SettingsPanel from './components/SettingsPanel';
import type { SettingsTabKey } from './components/SettingsPanel';
import AboutPanel from './components/AboutPanel';
import Footer from './components/Footer';

type TabId = 'chat' | 'history' | 'usage' | 'settings' | 'about';

export default function App() {
  const [tab, setTab] = useState<TabId>('chat');
  const [settingsTab, setSettingsTab] = useState<SettingsTabKey | undefined>(undefined);

  return (
    <FluentProvider theme={webLightTheme} style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      <div style={{
        flexShrink: 0,
        display: 'flex',
        alignItems: 'center',
        gap: 8,
        padding: '8px 12px 6px',
        borderBottom: `1px solid ${webLightTheme.colorNeutralStroke2}`,
      }}>
        <span aria-hidden="true" style={{ fontSize: 20, lineHeight: 1 }}>🦞</span>
        <div style={{ minWidth: 0 }}>
          <Body1Strong>SheetClaw</Body1Strong>
          <Caption1
            style={{
              display: 'block',
              color: webLightTheme.colorNeutralForeground3,
              lineHeight: 1.2,
            }}
          >
            Workbook agent · by Icon Learning &amp; Development
          </Caption1>
        </div>
      </div>

      {/* Tab nav */}
      <TabList
        selectedValue={tab}
        onTabSelect={(_, d: SelectTabData) => setTab(d.value as TabId)}
        style={{ flexShrink: 0, borderBottom: `1px solid ${webLightTheme.colorNeutralStroke1}` }}
        size="small"
      >
        <Tab value="chat">Chat</Tab>
        <Tab value="history">History</Tab>
        <Tab value="usage">Usage</Tab>
        <Tab value="settings">Settings</Tab>
        <Tab value="about">About</Tab>
      </TabList>


      {/* Active surface */}
      <div style={{ flex: 1, minHeight: 0, height: '100%' }}>
        {tab === 'chat'     && <ChatPanel onOpenSettings={(target) => { setSettingsTab(target); setTab('settings'); }} />}
        {tab === 'history'  && <HistoryPanel onOpenChat={() => setTab('chat')} />}
        {tab === 'usage'    && <UsageDashboard />}
        {tab === 'settings' && <SettingsPanel initialTab={settingsTab} />}
        {tab === 'about'    && <AboutPanel />}
      </div>

      {/* Persistent footer */}
      <Footer />
    </FluentProvider>
  );
}
