import React from 'react'
import { BrandMarkIcon, ThemeSparkIcon } from '../components/icons'
import { COLORS, COLOR_LABELS } from '../config/appConfig'

export function SettingsTab({
  accent,
  brandAccent,
  brandPreviewStyle,
  pdfLogs,
  setAccent,
  setBrandAccent,
  setPdfLogs,
  setSidebarMode,
  setTheme,
  sidebarMode,
  theme,
}) {
  return (
    <section className="tab-grid">
      <div className="card card-elevated">
        <div className="card-lead">
          <div>
            <div className="section-kicker">Налаштування пакета</div>
            <h2>Параметри інтерфейсу та бренд-блоку</h2>
          </div>
        </div>
      </div>

      <div className="card settings-showcase">
        <div className="settings-showcase-head">
          <div className="settings-showcase-icon">
            <ThemeSparkIcon className="settings-showcase-icon-svg" />
          </div>
          <div>
            <div className="section-kicker">Параметри інтерфейсу</div>
            <h2>Дизайн та Теми</h2>
          </div>
        </div>

        <div className="settings-showcase-grid">
          <section className="settings-panel">
            <div className="settings-panel-head">
              <div className="settings-panel-label">Колір бренд-блоку</div>
              <div className="brand-block-preview" style={brandPreviewStyle}>
                <div className="brand-block-preview-icon">
                  <BrandMarkIcon className="brand-block-preview-svg" />
                </div>
                <div className="brand-block-preview-text">
                  <strong>ADUserCreator</strong>
                  <span>Лого та бренд-блок</span>
                </div>
              </div>
            </div>
            <div className="brand-swatch-row">
              <button
                type="button"
                className={`brand-swatch brand-swatch-default ${brandAccent === 'default' ? 'active' : ''}`}
                onClick={() => setBrandAccent('default')}
                title="Default"
              >
                <span className="brand-swatch-check">✓</span>
              </button>
              {Object.keys(COLORS).map((key) => (
                <button
                  key={`brand-${key}`}
                  type="button"
                  className={`brand-swatch ${brandAccent === key ? 'active' : ''}`}
                  style={{ backgroundColor: COLORS[key].value }}
                  onClick={() => setBrandAccent(key)}
                  title={COLOR_LABELS[key] || key}
                >
                  <span className="brand-swatch-check">✓</span>
                </button>
                ))}
              </div>
              <p className="settings-panel-note">Цей параметр змінює верхній бренд-блок у сайдбарі та його логотипну підкладку.</p>
            </section>

          <section className="settings-panel">
            <div className="settings-panel-label">Загальна тема</div>
            <div className="theme-choice-grid">
              <button type="button" className={`theme-card ${theme === 'dark' ? 'active' : ''}`} onClick={() => setTheme('dark')}>
                <span className="theme-card-icon">☾</span>
                <span>Темна</span>
              </button>
              <button type="button" className={`theme-card ${theme === 'light' ? 'active' : ''}`} onClick={() => setTheme('light')}>
                <span className="theme-card-icon">☼</span>
                <span>Світла</span>
              </button>
            </div>
          </section>

          <section className="settings-panel">
            <div className="settings-panel-label">Режим sidebar</div>
            <div className="sidebar-mode-strip">
              <button type="button" className={sidebarMode === 'themed' ? 'active' : ''} onClick={() => setSidebarMode('themed')}>
                ⟳ Тематичний
              </button>
              <button type="button" className={sidebarMode === 'standard' ? 'active' : ''} onClick={() => setSidebarMode('standard')}>
                ◫ Стандартний
              </button>
            </div>
            <p className="settings-panel-note">У тематичному режимі сайдбар адаптується до обраного кольору бренду та системної теми.</p>
          </section>

          <section className="settings-panel">
            <div className="settings-panel-label">Колір кнопок (акцент)</div>
            <div className="accent-pill-row">
              {Object.keys(COLORS).map((key) => (
                <button key={key} type="button" className={`accent-pill ${accent === key ? 'active' : ''}`} onClick={() => setAccent(key)}>
                  {COLOR_LABELS[key] || key}
                </button>
              ))}
            </div>
          </section>
        </div>
      </div>

      <div className="settings-grid">
        <div className="card side-card">
          <div className="section-kicker">Документи</div>
          <h2>PDF-журнал з паролями</h2>

          <div className="card-header-row">
            <div className="muted">Історія згенерованих документів</div>
            <button className="btn btn-ghost" type="button" onClick={() => setPdfLogs([])}>
              Очистити історію
            </button>
          </div>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Назва файлу</th>
                  <th>Дата</th>
                  <th>Користувачів</th>
                  <th>Розмір</th>
                  <th>Дії</th>
                </tr>
              </thead>
              <tbody>
                {pdfLogs.map((file) => (
                  <tr key={file.id}>
                    <td>{file.name}</td>
                    <td>{file.date}</td>
                    <td>{file.users}</td>
                    <td>{file.size}</td>
                    <td>
                      <div className="table-actions">
                        <button className="table-action" type="button">
                          Переглянути
                        </button>
                        <button className="table-action" type="button">
                          Завантажити
                        </button>
                        <button className="table-action danger-text" type="button" onClick={() => setPdfLogs((prev) => prev.filter((item) => item.id !== file.id))}>
                          Видалити
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
                {!pdfLogs.length && (
                  <tr>
                    <td colSpan={5} className="empty-cell">
                      Історія PDF поки порожня
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </section>
  )
}
