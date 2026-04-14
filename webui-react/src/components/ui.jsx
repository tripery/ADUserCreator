import React from 'react'

export function NavItem({ label, icon, active, onClick }) {
  return (
    <button className={`menu-item ${active ? 'active' : ''}`} type="button" onClick={onClick}>
      <span className="menu-icon">{icon}</span>
      <span>{label}</span>
    </button>
  )
}

export function FeatureRow({ title, desc }) {
  return (
    <div className="checklist-item">
      <div className="checklist-title">{title}</div>
      <div className="checklist-text">{desc}</div>
    </div>
  )
}

export function PreviewCard({ previewRows, previewErrors }) {
  return (
    <div className="card">
      <div className="card-header-row">
        <div>
          <div className="section-kicker">Попередній перегляд</div>
          <h2>Користувачі перед створенням</h2>
        </div>
        <div className="muted">Показано {previewRows.length ? `1-${previewRows.length}` : 0} з {previewRows.length} користувачів</div>
      </div>

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>ПІБ</th>
              <th>Логін</th>
              <th>E-mail</th>
              <th>Підрозділ</th>
            </tr>
          </thead>
          <tbody>
            {previewRows.map((row, index) => (
              <tr key={`${row.login}-${index}`}>
                <td>{row.fullName}</td>
                <td>{row.login}</td>
                <td>{row.email}</td>
                <td>{row.unit || '—'}</td>
              </tr>
            ))}
            {!previewRows.length && (
              <tr>
                <td colSpan={4} className="empty-cell">
                  Немає даних для preview
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {previewErrors.length > 0 && (
        <div className="error-list">
          {previewErrors.map((error, index) => (
            <div className="error-item" key={`${error.fullName ?? 'row'}-${index}`}>
              {error.fullName || `Рядок ${error.sourceRow ?? '?'}`}: {error.error}
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

export function ResultsCard({ createResults, createErrors }) {
  return (
    <div className="card">
      <div className="card-header-row">
        <div>
          <div className="section-kicker">Результати</div>
          <h2>Створені облікові записи</h2>
        </div>
        <div className="muted">Локально, не передавайте назовні</div>
      </div>

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>ПІБ</th>
              <th>Логін</th>
              <th>E-mail</th>
              <th>Пароль</th>
              <th>Статус</th>
            </tr>
          </thead>
          <tbody>
            {createResults.map((row, index) => (
              <tr key={`${row.login}-${index}`}>
                <td>{row.fullName}</td>
                <td>{row.login}</td>
                <td>{row.email}</td>
                <td>{row.password || '—'}</td>
                <td>{row.status}</td>
              </tr>
            ))}
            {!createResults.length && (
              <tr>
                <td colSpan={5} className="empty-cell">
                  Результатів ще немає
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {createErrors.length > 0 && (
        <div className="error-list">
          {createErrors.map((error, index) => (
            <div className="error-item" key={`${error.fullName ?? 'create'}-${index}`}>
              {error.fullName || `Рядок ${error.sourceRow ?? '?'}`}: {error.error}
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

export function SessionCard({ apiStatus, selectedOuName, domainSuffix, latestLog }) {
  const controllerName = domainSuffix ? `DC-01.${domainSuffix}` : 'Невідомо'
  const apiLabel = apiStatus === 'online' ? 'Connected' : apiStatus === 'offline' ? 'Offline' : 'Checking'

  return (
    <div className="card side-card domain-card">
      <div className="section-kicker">Параметри домену</div>

      <div className="domain-pill">Служби каталогів</div>

      <div className="domain-metrics">
        <div className="domain-row">
          <span>Контролер</span>
          <strong>{controllerName}</strong>
        </div>
        <div className="domain-row">
          <span>DNS домен</span>
          <strong>{domainSuffix || 'Не вказано'}</strong>
        </div>
        <div className="domain-row">
          <span>Functional Level</span>
          <strong>Windows Server 2022</strong>
        </div>
        <div className="domain-row">
          <span>Статус API</span>
          <strong className={`status-inline ${apiStatus}`}>{apiLabel}</strong>
        </div>
      </div>

      <div className="domain-note">
        <div className="context-box-label">Нагадування</div>
        <div className="context-box-text">
          "{latestLog?.message || 'Адміністрування - це не про кнопки, а про порядок у системі.'}"
        </div>
        {selectedOuName ? <div className="domain-note-meta">Активний OU: {selectedOuName}</div> : null}
      </div>
    </div>
  )
}

export function LogCard({ logs, onClear }) {
  return (
    <div className="card side-card">
      <div className="section-kicker">Журнал виконання</div>
      <div className="card-header-row">
        <h2>Події сесії</h2>
        <button className="btn btn-ghost" type="button" onClick={onClear}>
          Очистити
        </button>
      </div>
      <pre className="log-box">{logs.map((log) => `[${log.stamp}] [${log.level}] ${log.message}`).join('\n')}</pre>
    </div>
  )
}

export function MetricCard({ label, value, tone }) {
  return (
    <div className={`mini-card ${tone}`}>
      <div className="mini-card-label">{label}</div>
      <div className="mini-card-value">{value}</div>
    </div>
  )
}

export function OuTreeNode({ node, depth, expandedOuNodes, selectedOu, isSearchMode, onToggle, onSelect }) {
  const hasChildren = node.children.length > 0
  const isExpanded = expandedOuNodes.has(node.dn)
  const shouldShowChildren = hasChildren && (isSearchMode || isExpanded)

  return (
    <>
      <div className={`ou-tree-node ${selectedOu === node.dn ? 'selected' : ''}`} style={{ paddingLeft: `${10 + depth * 16}px` }}>
        <button
          type="button"
          className={`ou-tree-expander ${hasChildren ? '' : 'leaf'}`}
          onClick={(event) => {
            event.stopPropagation()
            if (hasChildren) onToggle(node.dn)
          }}
        >
          {hasChildren ? (isSearchMode || isExpanded ? '▾' : '▸') : '•'}
        </button>

        <button type="button" className="ou-tree-pick" onClick={() => onSelect(node.dn)} title={node.dn}>
          <span className="ou-tree-name">{node.label}</span>
          <span className="ou-tree-dn">{node.dn}</span>
        </button>
      </div>

      {shouldShowChildren &&
        node.children.map((child) => (
          <OuTreeNode
            key={child.dn}
            node={child}
            depth={depth + 1}
            expandedOuNodes={expandedOuNodes}
            selectedOu={selectedOu}
            isSearchMode={isSearchMode}
            onToggle={onToggle}
            onSelect={onSelect}
          />
        ))}
    </>
  )
}
