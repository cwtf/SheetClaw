import { useState } from 'react';
import {
  Body1Strong,
  Button,
  Caption1,
  FluentProvider,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  webDarkTheme,
  webLightTheme,
} from '@fluentui/react-components';
import { useStore } from '../store/index';
import ChatPanel from './components/ChatPanel';
import HistoryPanel from './components/HistoryPanel';
import UsageDashboard from './components/UsageDashboard';
import SettingsPanel from './components/SettingsPanel';
import type { SettingsTabKey } from './components/SettingsPanel';
import AboutPanel from './components/AboutPanel';
import Footer from './components/Footer';

type TabId = 'chat' | 'history' | 'usage' | 'settings' | 'about';

const MENU_TABS: Array<{ id: Exclude<TabId, 'chat'>; label: string }> = [
  { id: 'history', label: 'History' },
  { id: 'usage', label: 'Usage' },
  { id: 'settings', label: 'Settings' },
  { id: 'about', label: 'About' },
];

function OverflowIcon() {
  return (
    <svg aria-hidden="true" width="18" height="18" viewBox="0 0 18 18" fill="currentColor">
      <circle cx="3" cy="9" r="1.5" />
      <circle cx="9" cy="9" r="1.5" />
      <circle cx="15" cy="9" r="1.5" />
    </svg>
  );
}

function BackIcon() {
  return (
    <svg aria-hidden="true" width="18" height="18" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="m8 4-6 6 6 6" />
      <path d="M2 10h16" />
    </svg>
  );
}

export default function App() {
  const [tab, setTab] = useState<TabId>('chat');
  const [settingsTab, setSettingsTab] = useState<SettingsTabKey | undefined>(undefined);
  const themePreference = useStore(s => s.appConfig.theme);
  const theme = themePreference === 'dark' ? webDarkTheme : webLightTheme;

  function openMenuTab(nextTab: Exclude<TabId, 'chat'>) {
    if (nextTab === 'settings') setSettingsTab(undefined);
    setTab(nextTab);
  }

  return (
    <FluentProvider
      theme={theme}
      style={{
        height: '100%',
        display: 'flex',
        flexDirection: 'column',
        colorScheme: themePreference,
        color: theme.colorNeutralForeground1,
        background: theme.colorNeutralBackground1,
      }}
    >
      <div style={{
        flexShrink: 0,
        display: 'flex',
        alignItems: 'center',
        gap: 8,
        padding: '8px 12px 6px',
        borderBottom: `1px solid ${theme.colorNeutralStroke2}`,
      }}>
        {tab !== 'chat' && (
          <Button
            appearance="subtle"
            size="small"
            icon={<BackIcon />}
            aria-label="Back to Chat"
            title="Back to Chat"
            onClick={() => setTab('chat')}
            style={{ width: 32, minWidth: 32, height: 32, padding: 0 }}
          />
        )}
        <span aria-hidden="true" style={{ fontSize: 20, lineHeight: 1 }}>🦞</span>
        <div style={{ minWidth: 0, flex: 1 }}>
          <Body1Strong>SheetClaw</Body1Strong>
          <Caption1
            style={{
              display: 'block',
              color: theme.colorNeutralForeground3,
              lineHeight: 1.2,
            }}
          >
            Workbook agent · by Icon Learning &amp; Development
          </Caption1>
        </div>
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button
              appearance="subtle"
              size="small"
              icon={<OverflowIcon />}
              aria-label="Open navigation menu"
              title="Menu"
              style={{ width: 32, minWidth: 32, height: 32, padding: 0 }}
            />
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              {MENU_TABS.map(item => (
                <MenuItem
                  key={item.id}
                  secondaryContent={tab === item.id ? <span style={{ color: webLightTheme.colorBrandForeground1 }}>✓</span> : undefined}
                  onClick={() => openMenuTab(item.id)}
                >
                  {item.label}
                </MenuItem>
              ))}
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>

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
