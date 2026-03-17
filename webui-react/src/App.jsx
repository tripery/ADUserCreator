import React, { startTransition, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'

const API_BASE = ''
const THEME_STORAGE_KEY = 'adusercreator-theme'
const COLORS = {
  blue: { value: '#2d68d6', soft: 'rgba(45, 104, 214, 0.12)' },
  emerald: { value: '#21916a', soft: 'rgba(33, 145, 106, 0.13)' },
  purple: { value: '#7c55d8', soft: 'rgba(124, 85, 216, 0.14)' },
  amber: { value: '#bc7a19', soft: 'rgba(188, 122, 25, 0.14)' },
}

const INITIAL_PDF_LOGS = [
  { id: 'pdf-1', name: 'passwords_credentials_2026-03-17.pdf', size: '142 KB', date: '2026-03-17 13:30', users: 12 },
  { id: 'pdf-2', name: 'it_department_logins.pdf', size: '89 KB', date: '2026-03-16 10:15', users: 5 },
]

export default function App() {
  const [activeTab, setActiveTab] = useState('creator')
  const [theme, setTheme] = useState(getInitialTheme)
  const [accent, setAccent] = useState('blue')
  const [apiStatus, setApiStatus] = useState('checking')
  const [fileName, setFileName] = useState('Файл не вибрано')
  const [domainSuffix, setDomainSuffix] = useState('')
  const [selectedOu, setSelectedOu] = useState('')
  const [ouOptions, setOuOptions] = useState([])
  const [groupOptions, setGroupOptions] = useState([])
  const [selectedGroups, setSelectedGroups] = useState([])
  const [groupInput, setGroupInput] = useState('')
  const [passwordNeverExpires, setPasswordNeverExpires] = useState(true)
  const [sourceUsers, setSourceUsers] = useState([])
  const [previewRows, setPreviewRows] = useState([])
  const [previewErrors, setPreviewErrors] = useState([])
  const [createResults, setCreateResults] = useState([])
  const [logs, setLogs] = useState([
    makeLog('INFO', 'React/Vite UI ініціалізовано'),
    makeLog('INFO', 'Очікування підключення до локального PowerShell API...'),
  ])
  const [pdfLogs, setPdfLogs] = useState(INITIAL_PDF_LOGS)
  const [isPreviewLoading, setIsPreviewLoading] = useState(false)
  const [isCreating, setIsCreating] = useState(false)
  const [isOuDropdownOpen, setIsOuDropdownOpen] = useState(false)
  const [ouSearch, setOuSearch] = useState('')
  const [expandedOuNodes, setExpandedOuNodes] = useState(new Set())

  const fileInputRef = useRef(null)
  const ouDropdownRef = useRef(null)
  const ouSearchInputRef = useRef(null)

  const accentColors = COLORS[accent] || COLORS.blue
  const latestLog = logs[logs.length - 1]
  const workflowSteps = [
    { title: '1. Імпорт Excel', description: 'Завантажте файл зі списком співробітників або студентів.' },
    { title: '2. Налаштування AD', description: 'Оберіть OU, домен і додаткові групи без переходів по інших вікнах.' },
    { title: '3. Контроль результату', description: 'Перевірте preview, dry-run і фінальне створення користувачів.' },
  ]

  const ouTree = useMemo(() => buildOuTree(ouOptions), [ouOptions])
  const filteredOuTree = useMemo(() => filterOuTree(ouTree, ouSearch), [ouTree, ouSearch])
  const selectedOuName = useMemo(() => {
    const match = ouOptions.find((item) => item.distinguishedName === selectedOu)
    return match?.name || selectedOu
  }, [ouOptions, selectedOu])
  const summaryCards = useMemo(
    () => [
      { label: 'Записів Excel', value: sourceUsers.length, tone: 'blue' },
      { label: 'У preview', value: previewRows.length, tone: 'green' },
      { label: 'Помилок', value: previewErrors.length, tone: 'red' },
      { label: 'Створено', value: createResults.length, tone: 'amber' },
    ],
    [sourceUsers.length, previewRows.length, previewErrors.length, createResults.length],
  )

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    try {
      window.localStorage.setItem(THEME_STORAGE_KEY, theme)
    } catch {}
  }, [theme])

  useEffect(() => {
    let disposed = false

    async function loadOptions() {
      try {
        const health = await fetch(`${API_BASE}/api/health`)
        if (!health.ok) throw new Error(`health check failed (HTTP ${health.status})`)
        const optionsRes = await fetch(`${API_BASE}/api/ad/options`)
        const options = await readApiResponse(optionsRes, 'AD options')
        if (disposed) return

        startTransition(() => {
          setApiStatus('online')
          setOuOptions(options.ous ?? [])
          setGroupOptions(options.groups ?? [])
          if (options.domain) setDomainSuffix((prev) => prev || options.domain)
          if (options.ous?.length) setSelectedOu((prev) => prev || options.ous[0].distinguishedName)
        })

        addLog(setLogs, 'OK', 'Підключено локальний PowerShell API та завантажено OU/групи')
      } catch (error) {
        if (disposed) return
        setApiStatus('offline')
        addLog(setLogs, 'ERROR', `API недоступний: ${error.message}. Запустіть webapi/server.ps1`)
      }
    }

    loadOptions()
    return () => {
      disposed = true
    }
  }, [])

  useEffect(() => {
    if (!ouTree.roots.length) return
    setExpandedOuNodes((prev) => {
      const next = new Set(prev)
      for (const root of ouTree.roots) next.add(root.dn)
      let cursor = selectedOu
      while (cursor) {
        const parent = ouTree.parentByDn.get(cursor)
        if (!parent) break
        next.add(parent)
        cursor = parent
      }
      return next
    })
  }, [ouTree, selectedOu])

  useEffect(() => {
    if (!isOuDropdownOpen) return undefined
    function handleClickOutside(event) {
      if (ouDropdownRef.current && !ouDropdownRef.current.contains(event.target)) {
        setIsOuDropdownOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [isOuDropdownOpen])

  useEffect(() => {
    if (isOuDropdownOpen && ouSearchInputRef.current) ouSearchInputRef.current.focus()
  }, [isOuDropdownOpen])

  async function handleFileSelected(event) {
    const file = event.target.files?.[0]
    if (!file) return
    try {
      setFileName(file.name)
      addLog(setLogs, 'INFO', `Читання Excel через SheetJS: ${file.name}`)
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array' })
      const sheetName = findSheetWithColumn(workbook, 'Вступник') ?? workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]
      if (!sheet) throw new Error('Не знайдено лист у книзі Excel')
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false })
      const users = normalizeExcelRows(rawRows)
      setSourceUsers(users)
      setCreateResults([])
      addLog(setLogs, 'OK', `Знайдено ${users.length} користувачів на листі '${sheetName}'`)
      if (!users.length) {
        setPreviewRows([])
        setPreviewErrors([])
        return
      }
      if (domainSuffix) await requestPreview(users, domainSuffix, selectedOu)
      else addLog(setLogs, 'WARN', 'Немає domainSuffix для побудови preview')
    } catch (error) {
      addLog(setLogs, 'ERROR', `Помилка читання Excel: ${error.message}`)
    }
  }

  function clearFile() {
    if (fileInputRef.current) fileInputRef.current.value = ''
    setFileName('Файл не вибрано')
    setSourceUsers([])
    setPreviewRows([])
    setPreviewErrors([])
    setCreateResults([])
    addLog(setLogs, 'INFO', 'Вибір файлу очищено')
  }

  async function requestPreview(users = sourceUsers, domain = domainSuffix, ou = selectedOu) {
    if (!users.length) return
    if (!domain) {
      addLog(setLogs, 'WARN', 'Немає домену для preview')
      return
    }

    setIsPreviewLoading(true)
    try {
      const response = await fetch(`${API_BASE}/api/users/preview`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ users, domainSuffix: domain, ou }),
      })
      const data = await readApiResponse(response, 'Preview')
      setPreviewRows(data.preview ?? [])
      setPreviewErrors(data.errors ?? [])
      addLog(setLogs, 'OK', `Оновлено preview (${(data.preview ?? []).length} записів)`)
      if ((data.errors ?? []).length) addLog(setLogs, 'WARN', `Є ${(data.errors ?? []).length} помилок розбору або preview`)
    } catch (error) {
      addLog(setLogs, 'ERROR', `Помилка preview: ${error.message}`)
    } finally {
      setIsPreviewLoading(false)
    }
  }

  function addGroup(value) {
    const next = String(value ?? '').trim()
    if (!next) return
    setSelectedGroups((prev) => {
      if (prev.includes(next)) {
        addLog(setLogs, 'WARN', `Група вже додана: ${next}`)
        return prev
      }
      addLog(setLogs, 'OK', `Групу додано: ${next}`)
      return [...prev, next]
    })
  }

  function removeGroup(group) {
    setSelectedGroups((prev) => prev.filter((item) => item !== group))
    addLog(setLogs, 'INFO', `Групу видалено: ${group}`)
  }

  function toggleOuNode(dn) {
    setExpandedOuNodes((prev) => {
      const next = new Set(prev)
      if (next.has(dn)) next.delete(dn)
      else next.add(dn)
      return next
    })
  }

  async function createUsers({ dryRun = false } = {}) {
    if (!sourceUsers.length) return addLog(setLogs, 'ERROR', 'Спочатку виберіть Excel-файл')
    if (!selectedOu) return addLog(setLogs, 'ERROR', 'Оберіть OU')
    if (!domainSuffix) return addLog(setLogs, 'ERROR', 'Вкажіть домен для UPN і E-mail')

    setIsCreating(true)
    try {
      const response = await fetch(`${API_BASE}/api/users/create`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({
          users: sourceUsers,
          ou: selectedOu,
          domainSuffix,
          groupsToAdd: selectedGroups,
          passwordNeverExpires,
          dryRun,
        }),
      })
      const data = await readApiResponse(response, 'Create')
      setCreateResults(data.created ?? [])
      setPreviewErrors(data.errors ?? [])
      addLog(setLogs, 'OK', `${dryRun ? 'Dry-run' : 'Створення'} завершено: ${(data.created ?? []).length} успішно, ${(data.errors ?? []).length} помилок`)
      if (!dryRun) setActiveTab('monitoring')
    } catch (error) {
      addLog(setLogs, 'ERROR', `Помилка create: ${error.message}`)
    } finally {
      setIsCreating(false)
    }
  }

  const appStyle = { '--blue': accentColors.value, '--blue-soft': accentColors.soft }

  return (
    <div className="app-shell" style={appStyle}>
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-icon">AD</div>
          <div>
            <div className="brand-kicker">ADUSERCREATOR</div>
            <div className="brand-text">Масове створення користувачів</div>
          </div>
        </div>
        <nav className="menu">
          <MenuItem label="Головна" icon="⌂" active={activeTab === 'creator'} onClick={() => setActiveTab('creator')} />
          <MenuItem label="Користувачі" icon="👤" active={activeTab === 'users'} onClick={() => setActiveTab('users')} />
          <MenuItem label="Моніторинг" icon="▣" active={activeTab === 'monitoring'} onClick={() => setActiveTab('monitoring')} />
          <MenuItem label="Налаштування" icon="⚙" active={activeTab === 'settings'} onClick={() => setActiveTab('settings')} />
        </nav>
        <div className="sidebar-note">
          <div className="sidebar-note-label">Активний контур</div>
          <div className="sidebar-note-value">{selectedOuName || 'OU не обрано'}</div>
          <div className="sidebar-note-muted">{domainSuffix || 'Домен не вказано'}</div>
        </div>
      </aside>
      <main className="main-area">
        <header className="topbar">
          <div className="topbar-left">
            <h1>{activeTab === 'settings' ? 'Конфігурація системи' : activeTab === 'monitoring' ? 'Моніторинг виконання' : activeTab === 'users' ? 'Перевірка користувачів' : 'Створення користувачів AD'}</h1>
            <p>Один робочий простір для імпорту Excel, перевірки preview, dry-run та фінального створення облікових записів.</p>
          </div>
          <div className="topbar-right">
            <div className={`status-badge ${apiStatus}`}>{apiStatus === 'online' ? 'API online' : apiStatus === 'offline' ? 'API offline' : 'API checking'}</div>
            <button className="theme-toggle" type="button" onClick={() => setTheme((prev) => prev === 'dark' ? 'light' : 'dark')}>
              <span className="theme-toggle-icon">{theme === 'dark' ? '☀' : '☾'}</span>
              <span>{theme === 'dark' ? 'Світла тема' : 'Темна тема'}</span>
            </button>
            <button className="profile" type="button">
              <span className="avatar">A</span>
              <span className="profile-name">admin</span>
            </button>
          </div>
        </header>
        <section className="content">
          <section className="hero-panel">
            <div className="hero-copy">
              <div className="eyebrow">Панель керування Active Directory</div>
              <h2>Акуратний фронтенд для масового створення користувачів без ручної рутини.</h2>
              <p>Завантажуйте Excel, одразу бачте готові логіни та пошту, запускайте dry-run і лише потім створюйте облікові записи в AD.</p>
              <div className="hero-actions">
                <button className="btn btn-success" type="button" disabled={isCreating} onClick={() => createUsers()}>{isCreating ? 'Створення...' : 'Створити користувачів'}</button>
                <button className="btn btn-ghost strong" type="button" onClick={() => createUsers({ dryRun: true })} disabled={isCreating || !sourceUsers.length}>Dry-run create</button>
              </div>
            </div>
            <div className="hero-stats">
              {summaryCards.map((card) => <div key={card.label} className={`hero-stat ${card.tone}`}><div className="hero-stat-value">{card.value}</div><div className="hero-stat-label">{card.label}</div></div>)}
            </div>
          </section>
          {renderTab({
            activeTab, fileName, fileInputRef, handleFileSelected, clearFile, domainSuffix, setDomainSuffix,
            selectedOu, selectedOuName, isOuDropdownOpen, setIsOuDropdownOpen, ouDropdownRef, ouSearchInputRef,
            ouSearch, setOuSearch, filteredOuTree, expandedOuNodes, toggleOuNode, sourceUsers, requestPreview,
            setSelectedOu, groupOptions, addGroup, groupInput, setGroupInput, selectedGroups, removeGroup,
            passwordNeverExpires, setPasswordNeverExpires, isPreviewLoading, isCreating, createUsers, previewRows,
            previewErrors, createResults, apiStatus, latestLog, workflowSteps, logs, setLogs, accent, setAccent,
            theme, setTheme, pdfLogs, setPdfLogs,
          })}
        </section>
      </main>
    </div>
  )
}

function renderTab(props) {
  if (props.activeTab === 'users') {
    return <section className="tab-grid"><PreviewCard previewRows={props.previewRows} previewErrors={props.previewErrors} /><ResultsCard createResults={props.createResults} /></section>
  }

  if (props.activeTab === 'monitoring') {
    return (
      <section className="tab-grid">
        <div className="metric-grid">
          <MetricCard label="Записів Excel" value={props.sourceUsers.length} tone="blue" />
          <MetricCard label="У preview" value={props.previewRows.length} tone="green" />
          <MetricCard label="Помилок" value={props.previewErrors.length} tone="red" />
          <MetricCard label="Створено" value={props.createResults.length} tone="amber" />
        </div>
        <LogCard logs={props.logs} onClear={() => props.setLogs([])} />
        <SessionCard apiStatus={props.apiStatus} selectedOuName={props.selectedOuName} domainSuffix={props.domainSuffix} latestLog={props.latestLog} />
      </section>
    )
  }

  if (props.activeTab === 'settings') {
    return (
      <section className="tab-grid">
        <div className="settings-grid">
          <div className="card side-card">
            <div className="section-kicker">Тема інтерфейсу</div>
            <h2>Вигляд системи</h2>
            <div className="settings-group">
              <div className="settings-row">
                <span>Режим</span>
                <div className="segmented-control">
                  <button type="button" className={props.theme === 'dark' ? 'active' : ''} onClick={() => props.setTheme('dark')}>Темна</button>
                  <button type="button" className={props.theme === 'light' ? 'active' : ''} onClick={() => props.setTheme('light')}>Світла</button>
                </div>
              </div>
              <div className="settings-row settings-row-stack">
                <span>Акцент</span>
                <div className="swatch-row">
                  {Object.keys(COLORS).map((key) => <button key={key} type="button" className={`swatch ${props.accent === key ? 'active' : ''}`} style={{ backgroundColor: COLORS[key].value }} onClick={() => props.setAccent(key)} />)}
                </div>
              </div>
            </div>
          </div>
          <div className="card side-card">
            <div className="section-kicker">PDF журнал</div>
            <h2>Лог-файли з паролями</h2>
            <div className="card-header-row">
              <div className="muted">Історія згенерованих документів</div>
              <button className="btn btn-ghost" type="button" onClick={() => props.setPdfLogs([])}>Очистити історію</button>
            </div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>Назва файлу</th><th>Дата</th><th>Користувачів</th><th>Розмір</th><th>Дії</th></tr></thead>
                <tbody>
                  {props.pdfLogs.map((file) => <tr key={file.id}><td>{file.name}</td><td>{file.date}</td><td>{file.users}</td><td>{file.size}</td><td><div className="table-actions"><button className="table-action" type="button">Переглянути</button><button className="table-action" type="button">Завантажити</button><button className="table-action danger-text" type="button" onClick={() => props.setPdfLogs((prev) => prev.filter((item) => item.id !== file.id))}>Видалити</button></div></td></tr>)}
                  {!props.pdfLogs.length && <tr><td colSpan={5} className="empty-cell">Історія PDF поки порожня</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </section>
    )
  }

  return (
    <section className="workspace-grid">
      <div className="workspace-main">
        <div className="card card-elevated">
          <div className="card-lead">
            <div><div className="section-kicker">Налаштування пакета</div><h2>Завантажте Excel і налаштуйте контекст створення</h2></div>
            <div className="pill-info">{props.fileName}</div>
          </div>
          <div className="field-block">
            <label className="label">Excel файл (*.xlsx)</label>
            <div className="file-row">
              <label className="btn btn-primary file-pick" htmlFor="excelFile">Вибрати файл</label>
              <input id="excelFile" ref={props.fileInputRef} type="file" accept=".xlsx" hidden onChange={props.handleFileSelected} />
              <div className="file-pill">{props.fileName}</div>
              <button className="btn btn-danger" type="button" onClick={props.clearFile}>Видалити</button>
            </div>
          </div>
          <div className="grid-2">
            <div className="field-block">
              <label className="label" htmlFor="domainSuffix">Домен для UPN / E-mail</label>
              <input id="domainSuffix" className="text-input" value={props.domainSuffix} onChange={(event) => props.setDomainSuffix(event.target.value)} onBlur={() => props.requestPreview()} placeholder="donnu.edu.ua" />
            </div>
            <div className="field-block">
              <label className="label" htmlFor="ouSelect">Виберіть OU</label>
              <div id="ouSelect" ref={props.ouDropdownRef} className="ou-tree-select">
                <button type="button" className="ou-tree-trigger" onClick={() => props.setIsOuDropdownOpen((prev) => !prev)}><span className="ou-tree-trigger-text">{props.selectedOuName || 'Оберіть OU...'}</span><span className="ou-tree-caret">{props.isOuDropdownOpen ? '▴' : '▾'}</span></button>
                <div className={`ou-tree-dropdown ${props.isOuDropdownOpen ? 'open' : ''}`}>
                  <input ref={props.ouSearchInputRef} className="text-input ou-tree-search" value={props.ouSearch} onChange={(event) => props.setOuSearch(event.target.value)} placeholder="Пошук OU або DN..." />
                  <div className="ou-tree-list" role="tree">
                    {!props.filteredOuTree.roots.length && <div className="ou-tree-empty">{props.ouSearch.trim() ? 'Нічого не знайдено' : 'Список OU порожній'}</div>}
                    {props.filteredOuTree.roots.map((node) => <OuTreeNode key={node.dn} node={node} depth={0} expandedOuNodes={props.expandedOuNodes} selectedOu={props.selectedOu} isSearchMode={Boolean(props.ouSearch.trim())} onToggle={props.toggleOuNode} onSelect={(dn) => { props.setSelectedOu(dn); props.setIsOuDropdownOpen(false); if (props.sourceUsers.length) props.requestPreview(props.sourceUsers, props.domainSuffix, dn) }} />)}
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="field-block">
            <label className="label">Додати користувачів до груп (опціонально)</label>
            <div className="chips-input">
              <div className="chips-left">
                <select className="chip-select" defaultValue="" onChange={(event) => { if (event.target.value) props.addGroup(event.target.value); event.target.value = '' }}>
                  <option value="">Вибрати групу з AD</option>
                  {props.groupOptions.slice(0, 400).map((group) => <option key={group.samAccountName} value={group.samAccountName}>{group.name} ({group.samAccountName})</option>)}
                </select>
                <input className="chip-text" value={props.groupInput} onChange={(event) => props.setGroupInput(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); props.addGroup(props.groupInput); props.setGroupInput('') } }} placeholder="Введіть SamAccountName групи" />
              </div>
              <button className="add-chip" type="button" onClick={() => { props.addGroup(props.groupInput); props.setGroupInput('') }}>❯</button>
            </div>
            <div className="chip-list">{props.selectedGroups.map((group) => <div className="chip" key={group}><span>{group}</span><button type="button" onClick={() => props.removeGroup(group)}>×</button></div>)}</div>
          </div>
          <div className="field-block inline-actions">
            <label className="checkbox-row"><input type="checkbox" checked={props.passwordNeverExpires} onChange={(event) => props.setPasswordNeverExpires(event.target.checked)} /><span>Пароль не має терміну дії</span></label>
            <button className="btn btn-ghost strong" type="button" onClick={() => props.requestPreview()} disabled={props.isPreviewLoading || !props.sourceUsers.length}>{props.isPreviewLoading ? 'Оновлення preview...' : 'Оновити preview'}</button>
            <button className="btn btn-ghost" type="button" onClick={() => props.createUsers({ dryRun: true })} disabled={props.isCreating || !props.sourceUsers.length}>Dry-run create</button>
          </div>
        </div>
        <div className="card action-bar"><div><div className="section-kicker">Фінальна дія</div><h2 className="action-title">Після dry-run можна запускати створення в AD</h2></div><button className="btn btn-success btn-lg" type="button" disabled={props.isCreating} onClick={() => props.createUsers()}>{props.isCreating ? 'Створення...' : 'Створити користувачів'}</button></div>
      </div>
      <aside className="workspace-side">
        <SessionCard apiStatus={props.apiStatus} selectedOuName={props.selectedOuName} domainSuffix={props.domainSuffix} latestLog={props.latestLog} />
        <div className="card side-card"><div className="section-kicker">Робочий сценарій</div><h2>Що робити далі</h2><div className="checklist">{props.workflowSteps.map((step) => <div key={step.title} className="checklist-item"><div className="checklist-title">{step.title}</div><div className="checklist-text">{step.description}</div></div>)}</div></div>
      </aside>
    </section>
  )
}

function MenuItem({ label, icon, active, onClick }) {
  return <button className={`menu-item ${active ? 'active' : ''}`} type="button" onClick={onClick}><span className="menu-icon">{icon}</span><span>{label}</span></button>
}

function PreviewCard({ previewRows, previewErrors }) {
  return <div className="card"><div className="card-header-row"><div><div className="section-kicker">Попередній перегляд</div><h2>Користувачі перед створенням</h2></div><div className="muted">Показано {previewRows.length ? `1-${previewRows.length}` : 0} з {previewRows.length} користувачів</div></div><div className="table-wrap"><table><thead><tr><th>ПІБ</th><th>Логін</th><th>E-mail</th><th>Підрозділ</th></tr></thead><tbody>{previewRows.map((row, index) => <tr key={`${row.login}-${index}`}><td>{row.fullName}</td><td>{row.login}</td><td>{row.email}</td><td>{row.unit || '—'}</td></tr>)}{!previewRows.length && <tr><td colSpan={4} className="empty-cell">Немає даних для preview</td></tr>}</tbody></table></div>{previewErrors.length > 0 && <div className="error-list">{previewErrors.map((error, index) => <div className="error-item" key={`${error.fullName ?? 'row'}-${index}`}>{error.fullName || `Рядок ${error.sourceRow ?? '?'}`}: {error.error}</div>)}</div>}</div>
}

function ResultsCard({ createResults }) {
  return <div className="card"><div className="card-header-row"><div><div className="section-kicker">Результати</div><h2>Створені облікові записи</h2></div><div className="muted">Локально, не передавайте назовні</div></div><div className="table-wrap"><table><thead><tr><th>ПІБ</th><th>Логін</th><th>E-mail</th><th>Пароль</th><th>Статус</th></tr></thead><tbody>{createResults.map((row, index) => <tr key={`${row.login}-${index}`}><td>{row.fullName}</td><td>{row.login}</td><td>{row.email}</td><td>{row.password || '—'}</td><td>{row.status}</td></tr>)}{!createResults.length && <tr><td colSpan={5} className="empty-cell">Результатів ще немає</td></tr>}</tbody></table></div></div>
}

function SessionCard({ apiStatus, selectedOuName, domainSuffix, latestLog }) {
  return <div className="card side-card accent-card"><div className="section-kicker">Поточний стан</div><h2>Сесія запуску</h2><div className="status-stack"><div className="status-row"><span>API</span><strong className={`status-inline ${apiStatus}`}>{apiStatus === 'online' ? 'Доступний' : apiStatus === 'offline' ? 'Недоступний' : 'Перевірка'}</strong></div><div className="status-row"><span>OU</span><strong>{selectedOuName || 'Не обрано'}</strong></div><div className="status-row"><span>Домен</span><strong>{domainSuffix || 'Не вказано'}</strong></div><div className="status-row"><span>Останній лог</span><strong>{latestLog?.level || '—'}</strong></div></div><div className="context-box"><div className="context-box-label">Остання подія</div><div className="context-box-text">{latestLog?.message || 'Журнал ще порожній.'}</div></div></div>
}

function LogCard({ logs, onClear }) {
  return <div className="card side-card"><div className="section-kicker">Журнал виконання</div><div className="card-header-row"><h2>Події сесії</h2><button className="btn btn-ghost" type="button" onClick={onClear}>Очистити</button></div><pre className="log-box">{logs.map((log) => `[${log.stamp}] [${log.level}] ${log.message}`).join('\n')}</pre></div>
}

function MetricCard({ label, value, tone }) {
  return <div className={`mini-card ${tone}`}><div className="mini-card-label">{label}</div><div className="mini-card-value">{value}</div></div>
}

function OuTreeNode({ node, depth, expandedOuNodes, selectedOu, isSearchMode, onToggle, onSelect }) {
  const hasChildren = node.children.length > 0
  const isExpanded = expandedOuNodes.has(node.dn)
  const shouldShowChildren = hasChildren && (isSearchMode || isExpanded)

  return (
    <>
      <div className={`ou-tree-node ${selectedOu === node.dn ? 'selected' : ''}`} style={{ paddingLeft: `${10 + depth * 16}px` }}>
        <button type="button" className={`ou-tree-expander ${hasChildren ? '' : 'leaf'}`} onClick={(event) => { event.stopPropagation(); if (hasChildren) onToggle(node.dn) }}>
          {hasChildren ? (isSearchMode || isExpanded ? '▾' : '▸') : '•'}
        </button>
        <button type="button" className="ou-tree-pick" onClick={() => onSelect(node.dn)} title={node.dn}>
          <span className="ou-tree-name">{node.label}</span>
          <span className="ou-tree-dn">{node.dn}</span>
        </button>
      </div>
      {shouldShowChildren && node.children.map((child) => <OuTreeNode key={child.dn} node={child} depth={depth + 1} expandedOuNodes={expandedOuNodes} selectedOu={selectedOu} isSearchMode={isSearchMode} onToggle={onToggle} onSelect={onSelect} />)}
    </>
  )
}

function makeLog(level, message) {
  return { level, message, stamp: new Date().toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit', second: '2-digit' }) }
}

function addLog(setLogs, level, message) {
  setLogs((prev) => [...prev, makeLog(level, message)])
}

async function readApiResponse(response, operationName) {
  const contentType = String(response.headers.get('content-type') || '').toLowerCase()
  const rawBody = await response.text()
  if (!contentType.includes('application/json')) {
    const snippet = rawBody.replace(/\s+/g, ' ').trim().slice(0, 200)
    throw new Error(`${operationName}: HTTP ${response.status}. ${snippet || 'Сервер повернув не-JSON відповідь.'}`)
  }

  let data
  try {
    data = rawBody ? JSON.parse(rawBody) : {}
  } catch {
    throw new Error(`${operationName}: HTTP ${response.status}. Некоректний JSON у відповіді.`)
  }

  if (!response.ok || data?.ok === false) throw new Error(data?.error || `${operationName}: HTTP ${response.status}`)
  return data
}

function getInitialTheme() {
  if (typeof window === 'undefined') return 'light'
  try {
    const savedTheme = window.localStorage.getItem(THEME_STORAGE_KEY)
    if (savedTheme === 'light' || savedTheme === 'dark') return savedTheme
  } catch {}
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

function buildOuTree(ouOptions) {
  const nodesByDn = new Map()
  const parentByDn = new Map()
  const dnSet = new Set()

  for (const ou of ouOptions) {
    const dn = String(ou?.distinguishedName || '').trim()
    if (!dn || dnSet.has(dn)) continue
    dnSet.add(dn)
    nodesByDn.set(dn, { dn, label: String(ou?.name || '').trim() || getLabelFromDn(dn), children: [] })
  }

  for (const dn of dnSet) {
    const parentDn = findClosestExistingParentDn(dn, dnSet)
    if (!parentDn) continue
    parentByDn.set(dn, parentDn)
    nodesByDn.get(parentDn).children.push(nodesByDn.get(dn))
  }

  const roots = []
  for (const [dn, node] of nodesByDn.entries()) {
    if (!parentByDn.has(dn)) roots.push(node)
  }
  sortTreeNodes(roots)
  return { roots, parentByDn }
}

function splitDnParts(dn) {
  const parts = []
  let token = ''
  let escaped = false

  for (const ch of String(dn || '')) {
    if (escaped) {
      token += ch
      escaped = false
      continue
    }
    if (ch === '\\') {
      token += ch
      escaped = true
      continue
    }
    if (ch === ',') {
      if (token.trim()) parts.push(token.trim())
      token = ''
      continue
    }
    token += ch
  }

  if (token.trim()) parts.push(token.trim())
  return parts
}

function removeFirstRdn(dn) {
  const parts = splitDnParts(dn)
  return parts.length <= 1 ? null : parts.slice(1).join(',')
}

function findClosestExistingParentDn(dn, dnSet) {
  let cursor = dn
  while (true) {
    const parent = removeFirstRdn(cursor)
    if (!parent) return null
    if (dnSet.has(parent)) return parent
    cursor = parent
  }
}

function getLabelFromDn(dn) {
  const firstPart = splitDnParts(dn)[0] || ''
  const eq = firstPart.indexOf('=')
  return eq === -1 ? (firstPart || dn) : `${firstPart.slice(eq + 1)} (${firstPart.slice(0, eq).toUpperCase()})`
}

function sortTreeNodes(nodes) {
  nodes.sort((a, b) => a.label.localeCompare(b.label, 'uk-UA', { sensitivity: 'base' }))
  for (const node of nodes) if (node.children.length) sortTreeNodes(node.children)
}

function filterOuTree(tree, search) {
  const query = normalizeSearchValue(search)
  if (!query) return tree
  return { roots: tree.roots.map((node) => filterOuNode(node, query)).filter(Boolean), parentByDn: tree.parentByDn }
}

function filterOuNode(node, query) {
  const matchedChildren = node.children.map((child) => filterOuNode(child, query)).filter(Boolean)
  if (!matchesSearch(node.label, node.dn, query) && matchedChildren.length === 0) return null
  return { ...node, children: matchedChildren }
}

function matchesSearch(label, dn, query) {
  const haystack = [
    normalizeSearchValue(label),
    normalizeSearchValue(dn),
    translitUaToLat(normalizeSearchValue(label)),
    translitUaToLat(normalizeSearchValue(dn)),
    translitLatToUa(normalizeSearchValue(label)),
    translitLatToUa(normalizeSearchValue(dn)),
  ].join(' ')
  return Array.from(new Set([query, translitUaToLat(query), translitLatToUa(query)])).some((variant) => variant && haystack.includes(variant))
}

function normalizeSearchValue(value) {
  return String(value || '').toLocaleLowerCase('uk-UA').normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/[’'`"]/g, '').replace(/\s+/g, ' ').trim()
}

function findSheetWithColumn(workbook, expectedHeader) {
  const target = normalizeHeader(expectedHeader)
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet || !sheet['!ref']) continue
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
    const header = (rows[0] ?? []).map((value) => normalizeHeader(String(value)))
    if (header.includes(target)) return sheetName
  }
  return null
}

function normalizeHeader(value) {
  return String(value).replace(/\u00A0/g, ' ').trim().toLowerCase()
}

function normalizeExcelRows(rows) {
  return rows.map((row, index) => ({ ...buildUserFromExcelRow(row), sourceRow: index + 2 })).filter((row) => row.fullName)
}

function buildUserFromExcelRow(row) {
  const fullName = String(row['Вступник'] || row['ПІБ'] || row['ПIБ'] || row['П.І.Б.'] || row['ПІП'] || '').trim()
  const unit = String(row['Структурний підрозділ'] || row.OU || row['Підрозділ'] || '').trim()
  const position = String(row.Должность || row.Посада || '').trim()
  const organization = String(row.Организация || row.Організація || '').trim()
  const name = splitUkrainianFullName(fullName)
  return {
    fullName,
    unit,
    position,
    organization,
    firstName: name.firstName,
    lastName: name.lastName,
    middleName: name.middleName,
    samAccountName: generateSamAccountName(name.firstName, name.lastName),
    password: generateTempPassword(),
  }
}

function splitUkrainianFullName(fullName) {
  const parts = String(fullName || '').trim().split(/\s+/).filter(Boolean)
  return parts.length === 0 ? { firstName: '', lastName: '', middleName: '' } : { lastName: parts[0] || '', firstName: parts[1] || '', middleName: parts.slice(2).join(' ') }
}

function generateSamAccountName(firstName, lastName) {
  const first = translitUaToLat(firstName || '')
  const last = translitUaToLat(lastName || '')
  const cleaned = `${first ? first[0] : ''}.${last}`.toLowerCase().replace(/[^a-z0-9._-]/g, '').replace(/^[._-]+|[._-]+$/g, '')
  return (cleaned || `user${Math.floor(1000 + Math.random() * 9000)}`).slice(0, 20)
}

function generateTempPassword(length = 12) {
  const lower = 'abcdefghijkmnpqrstuvwxyz'
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
  const digits = '23456789'
  const symbols = '!@#$%*?'
  const all = `${lower}${upper}${digits}${symbols}`
  const required = [
    lower[Math.floor(Math.random() * lower.length)],
    upper[Math.floor(Math.random() * upper.length)],
    digits[Math.floor(Math.random() * digits.length)],
    symbols[Math.floor(Math.random() * symbols.length)],
  ]
  while (required.length < length) required.push(all[Math.floor(Math.random() * all.length)])
  for (let index = required.length - 1; index > 0; index -= 1) {
    const target = Math.floor(Math.random() * (index + 1))
    const temp = required[index]
    required[index] = required[target]
    required[target] = temp
  }
  return required.join('')
}

function translitUaToLat(input) {
  const map = { а: 'a', б: 'b', в: 'v', г: 'h', ґ: 'g', д: 'd', е: 'e', є: 'ie', ж: 'zh', з: 'z', и: 'y', і: 'i', ї: 'i', й: 'i', к: 'k', л: 'l', м: 'm', н: 'n', о: 'o', п: 'p', р: 'r', с: 's', т: 't', у: 'u', ф: 'f', х: 'kh', ц: 'ts', ч: 'ch', ш: 'sh', щ: 'shch', ь: '', ю: 'iu', я: 'ia', "'": '', '’': '', '-': '-', ' ': '' }
  return String(input || '').toLowerCase().split('').map((ch) => map[ch] ?? ch).join('')
}

function translitLatToUa(input) {
  let value = String(input || '').toLowerCase()
  for (const [from, to] of [['shch', 'щ'], ['zh', 'ж'], ['kh', 'х'], ['ts', 'ц'], ['ch', 'ч'], ['sh', 'ш'], ['yu', 'ю'], ['iu', 'ю'], ['ya', 'я'], ['ia', 'я'], ['ye', 'є'], ['ie', 'є'], ['yi', 'ї'], ['yo', 'йо']]) {
    value = value.split(from).join(to)
  }
  const chars = { a: 'а', b: 'б', c: 'к', d: 'д', e: 'е', f: 'ф', g: 'г', h: 'х', i: 'і', j: 'й', k: 'к', l: 'л', m: 'м', n: 'н', o: 'о', p: 'п', q: 'к', r: 'р', s: 'с', t: 'т', u: 'у', v: 'в', w: 'в', x: 'кс', y: 'и', z: 'з' }
  return value.split('').map((ch) => chars[ch] ?? ch).join('')
}
