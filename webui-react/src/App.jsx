import React from 'react'
import { AppShell } from './components/layout/AppShell'
import { TAB_TITLES } from './config/uiContent'
import { useAppController } from './hooks/useAppController'
import { HomePage } from './pages/HomePage'
import { MonitoringPage } from './pages/MonitoringPage'
import { SettingsTab } from './pages/SettingsTab'
import { UsersPage } from './pages/UsersPage'

export default function App() {
  const controller = useAppController()

  return (
    <AppShell
      activeTab={controller.activeTab}
      apiStatus={controller.apiStatus}
      appStyle={controller.appStyle}
      onTabChange={controller.setActiveTab}
      onThemeToggle={() => controller.setTheme((prev) => (prev === 'dark' ? 'light' : 'dark'))}
      theme={controller.theme}
      title={TAB_TITLES[controller.activeTab] || TAB_TITLES.creator}
    >
      {controller.activeTab === 'users' && <UsersPage {...controller.sharedPageProps} />}
      {controller.activeTab === 'monitoring' && (
        <MonitoringPage
          apiStatus={controller.apiStatus}
          domainSuffix={controller.domainSuffix}
          latestLog={controller.latestLog}
          logs={controller.logs}
          selectedOuName={controller.selectedOuName}
          setLogs={controller.setLogs}
          summaryCards={controller.summaryCards}
        />
      )}
      {controller.activeTab === 'settings' && (
        <SettingsTab
          accent={controller.accent}
          brandAccent={controller.brandAccent}
          brandPreviewStyle={{ background: controller.brandPalette.bg }}
          pdfLogs={controller.pdfLogs}
          setAccent={controller.setAccent}
          setBrandAccent={controller.setBrandAccent}
          setPdfLogs={controller.setPdfLogs}
          setSidebarMode={controller.setSidebarMode}
          setTheme={controller.setTheme}
          sidebarMode={controller.sidebarMode}
          theme={controller.theme}
        />
      )}
      {controller.activeTab === 'creator' && <HomePage {...controller.sharedPageProps} />}
    </AppShell>
  )
}
