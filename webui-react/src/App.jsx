import { startTransition, useEffect, useRef, useState } from 'react'
import * as XLSX from 'xlsx'

const API_BASE = '/api'

const initialLogs = [
  { level: 'OK', message: 'React/Vite UI ініціалізовано' },
  { level: 'INFO', message: 'Очікування підключення до локального PowerShell API...' },
]

export default function App() {
  const [apiStatus, setApiStatus] = useState('checking')
  const [domainSuffix, setDomainSuffix] = useState('')
  const [ouOptions, setOuOptions] = useState([])
  const [groupOptions, setGroupOptions] = useState([])
  const [selectedOu, setSelectedOu] = useState('')
  const [selectedGroups, setSelectedGroups] = useState([])
  const [groupInput, setGroupInput] = useState('')
  const [passwordNeverExpires, setPasswordNeverExpires] = useState(false)
  const [sourceUsers, setSourceUsers] = useState([])
  const [previewRows, setPreviewRows] = useState([])
  const [previewErrors, setPreviewErrors] = useState([])
  const [fileName, setFileName] = useState('Файл не вибрано')
  const [isPreviewLoading, setIsPreviewLoading] = useState(false)
  const [isCreating, setIsCreating] = useState(false)
  const [createResults, setCreateResults] = useState([])
  const [logs, setLogs] = useState(initialLogs)
  const fileInputRef = useRef(null)

  function addLog(level, message) {
    const stamp = new Date().toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
    setLogs((prev) => [...prev, { level, message, stamp }])
  }

  useEffect(() => {
    let disposed = false

    async function loadOptions() {
      try {
        const health = await fetch(`${API_BASE}/health`)
        if (!health.ok) throw new Error('health check failed')

        const optionsRes = await fetch(`${API_BASE}/ad/options`)
        if (!optionsRes.ok) throw new Error(`AD options HTTP ${optionsRes.status}`)
        const options = await optionsRes.json()

        if (disposed) return

        startTransition(() => {
          setApiStatus('online')
          setOuOptions(options.ous ?? [])
          setGroupOptions(options.groups ?? [])
          if (options.domain) setDomainSuffix((prev) => prev || options.domain)
          if (options.ous?.length) setSelectedOu((prev) => prev || options.ous[0].distinguishedName)
        })

        addLog('OK', 'Підключено локальний PowerShell API та завантажено OU/групи')
      } catch (error) {
        if (disposed) return
        setApiStatus('offline')
        addLog('ERROR', `API недоступний: ${error.message}. Запустіть webapi/server.ps1`)
      }
    }

    loadOptions()
    return () => {
      disposed = true
    }
  }, [])

  async function handleFileSelected(event) {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      setFileName(file.name)
      addLog('INFO', `Читання Excel через SheetJS: ${file.name}`)

      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array' })
      const sheetName = findSheetWithColumn(workbook, 'Вступник') ?? workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]
      if (!sheet) throw new Error('Не знайдено лист у книзі Excel')

      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false })
      const users = normalizeExcelRows(rawRows)
      setSourceUsers(users)

      addLog('OK', `Знайдено ${users.length} користувачів на листі '${sheetName}'`)

      if (!users.length) {
        setPreviewRows([])
        setPreviewErrors([])
        return
      }

      if (domainSuffix) {
        await requestPreview(users, domainSuffix, selectedOu)
      } else {
        addLog('WARN', 'Немає domainSuffix для побудови preview')
      }
    } catch (error) {
      addLog('ERROR', `Помилка читання Excel: ${error.message}`)
    }
  }

  function clearFile() {
    if (fileInputRef.current) fileInputRef.current.value = ''
    setFileName('Файл не вибрано')
    setSourceUsers([])
    setPreviewRows([])
    setPreviewErrors([])
    addLog('INFO', 'Вибір файлу очищено')
  }

  async function requestPreview(users = sourceUsers, domain = domainSuffix, ou = selectedOu) {
    if (!users.length) return
    if (!domain) {
      addLog('WARN', 'Немає домену для preview')
      return
    }

    setIsPreviewLoading(true)
    try {
      const response = await fetch(`${API_BASE}/users/preview`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ users, domainSuffix: domain, ou }),
      })
      const data = await response.json()
      if (!response.ok || !data.ok) throw new Error(data.error ?? `HTTP ${response.status}`)

      setPreviewRows(data.preview ?? [])
      setPreviewErrors(data.errors ?? [])
      addLog('OK', `Оновлено preview (${(data.preview ?? []).length} записів)`)
      if ((data.errors ?? []).length) addLog('WARN', `Є ${(data.errors ?? []).length} помилок розбору/preview`)
    } catch (error) {
      addLog('ERROR', `Помилка preview: ${error.message}`)
    } finally {
      setIsPreviewLoading(false)
    }
  }

  function addGroup(value) {
    const next = String(value ?? '').trim()
    if (!next) return
    setSelectedGroups((prev) => {
      if (prev.includes(next)) {
        addLog('WARN', `Група вже додана: ${next}`)
        return prev
      }
      addLog('OK', `Групу додано: ${next}`)
      return [...prev, next]
    })
  }

  function removeGroup(group) {
    setSelectedGroups((prev) => prev.filter((g) => g !== group))
    addLog('INFO', `Групу видалено: ${group}`)
  }

  async function createUsers({ dryRun = false } = {}) {
    if (!sourceUsers.length) return addLog('ERROR', 'Спочатку виберіть Excel-файл')
    if (!selectedOu) return addLog('ERROR', 'Оберіть OU')
    if (!domainSuffix) return addLog('ERROR', 'Вкажіть домен (domainSuffix)')

    setIsCreating(true)
    try {
      const response = await fetch(`${API_BASE}/users/create`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          users: sourceUsers,
          ou: selectedOu,
          domainSuffix,
          groupsToAdd: selectedGroups,
          passwordNeverExpires,
          dryRun,
        }),
      })
      const data = await response.json()
      if (!response.ok || !data.ok) throw new Error(data.error ?? `HTTP ${response.status}`)

      setCreateResults(data.created ?? [])
      setPreviewErrors(data.errors ?? [])
      addLog('OK', `${dryRun ? 'Dry-run' : 'Створення'} завершено: ${(data.created ?? []).length} успішно, ${(data.errors ?? []).length} помилок`)
    } catch (error) {
      addLog('ERROR', `Помилка create: ${error.message}`)
    } finally {
      setIsCreating(false)
    }
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand"><div className="brand-icon">👥</div><div className="brand-text">Створення користувачів AD</div></div>
        <nav className="menu">
          <button className="menu-item active" type="button"><span className="menu-icon">⌂</span><span>Головна</span></button>
          <button className="menu-item" type="button"><span className="menu-icon">👤</span><span>Користувачі</span></button>
          <button className="menu-item" type="button"><span className="menu-icon">👥</span><span>Групи</span></button>
        </nav>
        <div className="menu-footer"><button className="menu-item" type="button"><span className="menu-icon">⚙</span><span>Налаштування</span></button></div>
      </aside>

      <main className="main-area">
        <header className="topbar">
          <div className="topbar-left">
            <h1>Створення користувачів AD</h1>
            <p>React + Vite + SheetJS + PowerShell HTTP API</p>
          </div>
          <div className="topbar-right">
            <div className={`status-badge ${apiStatus}`}>{apiStatus === 'online' ? 'API online' : apiStatus === 'offline' ? 'API offline' : 'API...'}</div>
            <button className="profile" type="button"><span className="avatar">A</span><span className="profile-name">admin</span><span className="caret">⌄</span></button>
          </div>
        </header>

        <section className="content">
          <div className="card">
            <h2>Завантажте Excel-файл зі списком користувачів</h2>
            <div className="field-block">
              <label className="label">Excel файл (*.xlsx)</label>
              <div className="file-row">
                <label className="btn btn-primary file-pick" htmlFor="excelFile">📄 Вибрати файл</label>
                <input id="excelFile" ref={fileInputRef} type="file" accept=".xlsx" hidden onChange={handleFileSelected} />
                <div className="file-pill">{fileName}</div>
                <button className="btn btn-danger" type="button" onClick={clearFile}>Видалити</button>
              </div>
            </div>

            <div className="grid-2">
              <div className="field-block">
                <label className="label" htmlFor="domainSuffix">Домен для UPN / E-mail</label>
                <input id="domainSuffix" className="text-input" value={domainSuffix} onChange={(e) => setDomainSuffix(e.target.value)} onBlur={() => requestPreview()} placeholder="donnu.edu.ua" />
              </div>
              <div className="field-block">
                <label className="label" htmlFor="ouSelect">Виберіть OU</label>
                <select id="ouSelect" className="input-select" value={selectedOu} onChange={(e) => { setSelectedOu(e.target.value); if (sourceUsers.length) requestPreview(sourceUsers, domainSuffix, e.target.value) }}>
                  <option value="">Оберіть OU...</option>
                  {ouOptions.map((ou) => <option key={ou.distinguishedName} value={ou.distinguishedName}>{ou.name} | {ou.distinguishedName}</option>)}
                </select>
              </div>
            </div>

            <div className="field-block">
              <label className="label">Додати користувачів до груп (опціонально)</label>
              <div className="chips-input">
                <div className="chips-left">
                  <select className="chip-select" defaultValue="" onChange={(e) => { if (e.target.value) addGroup(e.target.value); e.target.value = '' }}>
                    <option value="">Вибрати групу з AD</option>
                    {groupOptions.slice(0, 400).map((g) => <option key={g.samAccountName} value={g.samAccountName}>{g.name} ({g.samAccountName})</option>)}
                  </select>
                  <input className="chip-text" value={groupInput} onChange={(e) => setGroupInput(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); addGroup(groupInput); setGroupInput('') } }} placeholder="Введіть SamAccountName групи" />
                </div>
                <button className="add-chip" type="button" onClick={() => { addGroup(groupInput); setGroupInput('') }}>❯</button>
              </div>
              <div className="chip-list">{selectedGroups.map((group) => <div className="chip" key={group}><span>{group}</span><button type="button" onClick={() => removeGroup(group)}>×</button></div>)}</div>
            </div>

            <div className="field-block inline-actions">
              <label className="checkbox-row"><input type="checkbox" checked={passwordNeverExpires} onChange={(e) => setPasswordNeverExpires(e.target.checked)} /><span>Пароль не має терміну дії</span></label>
              <button className="btn btn-ghost" type="button" onClick={() => requestPreview()} disabled={isPreviewLoading || !sourceUsers.length}>{isPreviewLoading ? 'Оновлення preview...' : 'Оновити preview'}</button>
              <button className="btn btn-ghost" type="button" onClick={() => createUsers({ dryRun: true })} disabled={isCreating || !sourceUsers.length}>Dry-run create</button>
            </div>
          </div>

          <div className="card">
            <div className="card-header-row"><h2>Попередній перегляд користувачів</h2><div className="muted">Показано 1-{previewRows.length} з {previewRows.length} користувачів</div></div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>ПІБ</th><th>Логін</th><th>E-mail</th><th>Підрозділ</th></tr></thead>
                <tbody>
                  {previewRows.map((row, index) => <tr key={`${row.login}-${index}`}><td>{row.fullName}</td><td>{row.login}</td><td>{row.email}</td><td>{row.unit || '—'}</td></tr>)}
                  {!previewRows.length && <tr><td colSpan={4} className="empty-cell">Немає даних для preview</td></tr>}
                </tbody>
              </table>
            </div>
            {previewErrors.length > 0 && <div className="error-list">{previewErrors.map((err, idx) => <div className="error-item" key={`${err.fullName ?? 'row'}-${idx}`}>{err.fullName || `Рядок ${err.sourceRow ?? '?'}`}: {err.error}</div>)}</div>}
          </div>

          <div className="card action-bar"><div className="action-bar-spacer" /><button className="btn btn-success btn-lg" type="button" disabled={isCreating} onClick={() => createUsers()}>{isCreating ? 'Створення...' : '👤＋ Створити користувачів'}</button></div>

          <div className="card">
            <div className="card-header-row"><h2>Результати створення (включно з паролями)</h2><div className="muted">Локально, не передавайте назовні</div></div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>ПІБ</th><th>Логін</th><th>E-mail</th><th>Пароль</th><th>Статус</th></tr></thead>
                <tbody>
                  {createResults.map((row, index) => <tr key={`${row.login}-${index}`}><td>{row.fullName}</td><td>{row.login}</td><td>{row.email}</td><td>{row.password || '—'}</td><td>{row.status}</td></tr>)}
                  {!createResults.length && <tr><td colSpan={5} className="empty-cell">Результатів ще немає</td></tr>}
                </tbody>
              </table>
            </div>
          </div>

          <div className="card">
            <div className="card-header-row"><h2>Журнал виконання</h2><button className="btn btn-ghost" type="button" onClick={() => setLogs([])}>Очистити</button></div>
            <pre className="log-box">{logs.map((l) => `[${l.stamp ?? '--:--:--'}] [${l.level}] ${l.message}`).join('\n')}</pre>
          </div>
        </section>
      </main>
    </div>
  )
}

function findSheetWithColumn(workbook, expectedHeader) {
  const target = normalizeHeader(expectedHeader)
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet || !sheet['!ref']) continue
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
    const header = (rows[0] ?? []).map((v) => normalizeHeader(String(v)))
    if (header.includes(target)) return sheetName
  }
  return null
}

function normalizeHeader(value) {
  return String(value).replace(/\u00A0/g, ' ').trim().toLowerCase()
}

function normalizeExcelRows(rows) {
  return rows
    .map((row, index) => ({
      fullName: String(row['Вступник'] || row['ПІБ'] || row['ПIБ'] || row['П.І.Б.'] || row['ПІП'] || '').trim(),
      unit: String(row['Структурний підрозділ'] || row['OU'] || row['Підрозділ'] || '').trim(),
      sourceRow: index + 2,
    }))
    .filter((row) => row.fullName)
}
