import React from 'react'
import { BrandMarkIcon } from '../icons'
import { NavItem } from '../ui'

export function AppShell({ activeTab, apiStatus, appStyle, children, onTabChange, theme, onThemeToggle, title }) {
  return (
    <div className="app-shell" style={appStyle}>
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-icon">
            <BrandMarkIcon className="brand-icon-svg" />
          </div>
          <div>
            <div className="brand-kicker">ADUSERCREATOR</div>
            <div className="brand-text">Масове створення користувачів</div>
          </div>
        </div>

        <nav className="menu">
          <NavItem label="Головна" icon="⌂" active={activeTab === 'creator'} onClick={() => onTabChange('creator')} />
          <NavItem label="Користувачі" icon="👤" active={activeTab === 'users'} onClick={() => onTabChange('users')} />
          <NavItem label="Моніторинг" icon="▣" active={activeTab === 'monitoring'} onClick={() => onTabChange('monitoring')} />
          <NavItem label="Налаштування" icon="⚙" active={activeTab === 'settings'} onClick={() => onTabChange('settings')} />
        </nav>
      </aside>

      <main className="main-area">
        <header className="topbar">
          <div className="topbar-left">
            <h1>{title}</h1>
            <p>Один робочий простір для імпорту Excel, перевірки preview, dry-run та фінального створення облікових записів.</p>
          </div>

          <div className="topbar-right">
            <div className={`status-badge ${apiStatus}`}>
              {apiStatus === 'online' ? 'API online' : apiStatus === 'offline' ? 'API offline' : 'API checking'}
            </div>

            <button className="theme-toggle" type="button" onClick={onThemeToggle}>
              <span className="theme-toggle-icon">{theme === 'dark' ? '☀' : '☾'}</span>
              <span>{theme === 'dark' ? 'Світла тема' : 'Темна тема'}</span>
            </button>

            <button className="profile" type="button">
              <span className="avatar">A</span>
              <span className="profile-name">admin</span>
            </button>
          </div>
        </header>

        <section className="content">{children}</section>
      </main>
    </div>
  )
}
