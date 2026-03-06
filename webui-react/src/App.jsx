import { startTransition, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'

const API_BASE = import.meta.env.VITE_API_BASE_URL || ''

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
  const [ouSearch, setOuSearch] = useState('')
  const [expandedOuNodes, setExpandedOuNodes] = useState(() => new Set())
  const [isOuDropdownOpen, setIsOuDropdownOpen] = useState(false)
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
  const ouDropdownRef = useRef(null)
  const ouSearchInputRef = useRef(null)
  const ouTree = useMemo(() => buildOuTree(ouOptions), [ouOptions])
  const filteredOuTree = useMemo(() => filterOuTree(ouTree, ouSearch), [ouTree, ouSearch])

  function addLog(level, message) {
    const stamp = new Date().toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
    setLogs((prev) => [...prev, { level, message, stamp }])
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

    if (!response.ok || data?.ok === false) {
      throw new Error(data?.error || `${operationName}: HTTP ${response.status}`)
    }

    return data
  }

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
    if (!isOuDropdownOpen) return
    function handleClickOutside(event) {
      if (ouDropdownRef.current && !ouDropdownRef.current.contains(event.target)) {
        setIsOuDropdownOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [isOuDropdownOpen])

  useEffect(() => {
    if (isOuDropdownOpen && ouSearchInputRef.current) {
      ouSearchInputRef.current.focus()
    }
  }, [isOuDropdownOpen])
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
      const response = await fetch(`${API_BASE}/api/users/preview`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ users, domainSuffix: domain, ou }),
      })
      const data = await readApiResponse(response, 'Preview')

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

  function toggleOuNode(dn) {
    setExpandedOuNodes((prev) => {
      const next = new Set(prev)
      if (next.has(dn)) next.delete(dn)
      else next.add(dn)
      return next
    })
  }
  async function createUsers({ dryRun = false } = {}) {
    if (!sourceUsers.length) return addLog('ERROR', 'Спочатку виберіть Excel-файл')
    if (!selectedOu) return addLog('ERROR', 'Оберіть OU')
    if (!domainSuffix) return addLog('ERROR', 'Вкажіть домен (domainSuffix)')

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
                <div id="ouSelect" ref={ouDropdownRef} className="ou-tree-select" aria-label="Вибір OU деревом">
                  <button type="button" className="ou-tree-trigger" onClick={() => setIsOuDropdownOpen((prev) => !prev)}>
                    <span className="ou-tree-trigger-text" title={selectedOu || 'Оберіть OU...'}>{selectedOu || 'Оберіть OU...'}</span>
                    <span className="ou-tree-caret">{isOuDropdownOpen ? '▴' : '▾'}</span>
                  </button>
                  <div className={`ou-tree-dropdown ${isOuDropdownOpen ? 'open' : ''}`}>
                    <input
                      ref={ouSearchInputRef}
                      className="text-input ou-tree-search"
                      value={ouSearch}
                      onChange={(e) => setOuSearch(e.target.value)}
                      placeholder="Пошук OU або DN..."
                    />
                    <div className="ou-tree-list" role="tree">
                      {!filteredOuTree.roots.length && <div className="ou-tree-empty">{ouSearch.trim() ? 'Нічого не знайдено' : 'Список OU порожній'}</div>}
                      {filteredOuTree.roots.map((node) => (
                        <OuTreeNode
                          key={node.dn}
                          node={node}
                          depth={0}
                          expandedOuNodes={expandedOuNodes}
                          selectedOu={selectedOu}
                          isSearchMode={Boolean(ouSearch.trim())}
                          onToggle={toggleOuNode}
                          onSelect={(dn) => {
                            setSelectedOu(dn)
                            setIsOuDropdownOpen(false)
                            if (sourceUsers.length) requestPreview(sourceUsers, domainSuffix, dn)
                          }}
                        />
                      ))}
                    </div>
                  </div>
                </div>
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

function OuTreeNode({ node, depth, expandedOuNodes, selectedOu, isSearchMode, onToggle, onSelect }) {
  const hasChildren = node.children.length > 0
  const isExpanded = expandedOuNodes.has(node.dn)
  const isSelected = selectedOu === node.dn
  const shouldShowChildren = hasChildren && (isSearchMode || isExpanded)

  return (
    <>
      <div className={`ou-tree-node ${isSelected ? 'selected' : ''}`} style={{ paddingLeft: `${10 + depth * 16}px` }}>
        <button
          type="button"
          className={`ou-tree-expander ${hasChildren ? '' : 'leaf'}`}
          onClick={(event) => {
            event.stopPropagation()
            if (hasChildren) onToggle(node.dn)
          }}
          aria-label={hasChildren ? (isExpanded ? 'Згорнути' : 'Розгорнути') : 'Лист'}
        >
          {hasChildren ? (isSearchMode || isExpanded ? '▾' : '▸') : '•'}
        </button>
        <button type="button" className="ou-tree-pick" onClick={() => onSelect(node.dn)} title={node.dn}>
          <span className="ou-tree-name">{node.label}</span>
          <span className="ou-tree-dn">{node.dn}</span>
        </button>
      </div>
      {shouldShowChildren && node.children.map((child) => (
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

function buildOuTree(ouOptions) {
  const nodesByDn = new Map()
  const parentByDn = new Map()
  const dnSet = new Set()

  for (const ou of ouOptions) {
    const dn = String(ou?.distinguishedName || '').trim()
    if (!dn || dnSet.has(dn)) continue
    dnSet.add(dn)
    nodesByDn.set(dn, {
      dn,
      label: String(ou?.name || '').trim() || getLabelFromDn(dn),
      children: []
    })
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
  if (parts.length <= 1) return null
  return parts.slice(1).join(',')
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
  if (eq === -1) return firstPart || dn

  const type = firstPart.slice(0, eq).toUpperCase()
  const value = firstPart.slice(eq + 1)
  return `${value} (${type})`
}

function sortTreeNodes(nodes) {
  nodes.sort((a, b) => a.label.localeCompare(b.label, 'uk-UA', { sensitivity: 'base' }))
  for (const node of nodes) {
    if (node.children.length > 0) sortTreeNodes(node.children)
  }
}

function filterOuTree(tree, search) {
  const query = normalizeSearchValue(search)
  if (!query) return tree

  const roots = tree.roots
    .map((node) => filterOuNode(node, query))
    .filter(Boolean)

  return { roots, parentByDn: tree.parentByDn }
}

function filterOuNode(node, query) {
  const matchSelf = matchesSearch(node.label, node.dn, query)
  const matchedChildren = node.children
    .map((child) => filterOuNode(child, query))
    .filter(Boolean)

  if (!matchSelf && matchedChildren.length === 0) return null
  return { ...node, children: matchedChildren }
}

function matchesSearch(label, dn, query) {
  const normalizedLabel = normalizeSearchValue(label)
  const normalizedDn = normalizeSearchValue(dn)
  const queryVariants = buildSearchVariants(query)
  const haystack = [normalizedLabel, normalizedDn, translitUaToLat(normalizedLabel), translitUaToLat(normalizedDn), translitLatToUa(normalizedLabel), translitLatToUa(normalizedDn)].join(' ')
  return queryVariants.some((variant) => variant && haystack.includes(variant))
}

function buildSearchVariants(query) {
  const normalized = normalizeSearchValue(query)
  return Array.from(new Set([
    normalized,
    translitUaToLat(normalized),
    translitLatToUa(normalized)
  ]))
}

function normalizeSearchValue(value) {
  return String(value || '')
    .toLocaleLowerCase('uk-UA')
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[’'`"]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
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
      ...buildUserFromExcelRow(row),
      sourceRow: index + 2
    }))
    .filter((row) => row.fullName)
}

function buildUserFromExcelRow(row) {
  const fullName = String(row['Вступник'] || row['ПІБ'] || row['ПIБ'] || row['П.І.Б.'] || row['ПІП'] || '').trim()
  const unit = String(row['Структурний підрозділ'] || row['OU'] || row['Підрозділ'] || '').trim()
  const position = String(row['Должность'] || row['Посада'] || '').trim()
  const organization = String(row['Организация'] || row['Організація'] || '').trim()
  const name = splitUkrainianFullName(fullName)
  const samAccountName = generateSamAccountName(name.firstName, name.lastName)
  const password = generateTempPassword()

  return {
    fullName,
    unit,
    position,
    organization,
    firstName: name.firstName,
    lastName: name.lastName,
    middleName: name.middleName,
    samAccountName,
    password
  }
}

function splitUkrainianFullName(fullName) {
  const parts = String(fullName || '').trim().split(/\s+/).filter(Boolean)
  if (parts.length === 0) {
    return { firstName: '', lastName: '', middleName: '' }
  }

  const lastName = parts[0] || ''
  const firstName = parts[1] || ''
  const middleName = parts.slice(2).join(' ')
  return { firstName, lastName, middleName }
}

function generateSamAccountName(firstName, lastName) {
  const first = translitUaToLat(firstName || '')
  const last = translitUaToLat(lastName || '')
  const base = `${first ? first[0] : ''}.${last}`.toLowerCase().replace(/[^a-z0-9._-]/g, '')
  const cleaned = base.replace(/^[._-]+|[._-]+$/g, '')
  if (cleaned) return cleaned.slice(0, 20)

  const fallback = `user${Math.floor(1000 + Math.random() * 9000)}`
  return fallback.slice(0, 20)
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
    symbols[Math.floor(Math.random() * symbols.length)]
  ]

  while (required.length < length) {
    required.push(all[Math.floor(Math.random() * all.length)])
  }

  for (let i = required.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1))
    const tmp = required[i]
    required[i] = required[j]
    required[j] = tmp
  }

  return required.join('')
}

function translitUaToLat(input) {
  const map = {
    а: 'a', б: 'b', в: 'v', г: 'h', ґ: 'g', д: 'd', е: 'e', є: 'ie', ж: 'zh', з: 'z',
    и: 'y', і: 'i', ї: 'i', й: 'i', к: 'k', л: 'l', м: 'm', н: 'n', о: 'o', п: 'p',
    р: 'r', с: 's', т: 't', у: 'u', ф: 'f', х: 'kh', ц: 'ts', ч: 'ch', ш: 'sh', щ: 'shch',
    ь: '', ю: 'iu', я: 'ia', "'": '', '’': '', '-': '-', ' ': ''
  }

  return String(input || '')
    .toLowerCase()
    .split('')
    .map((ch) => map[ch] ?? ch)
    .join('')
}

function translitLatToUa(input) {
  let value = String(input || '').toLowerCase()
  const digraphs = [
    ['shch', 'щ'],
    ['zh', 'ж'],
    ['kh', 'х'],
    ['ts', 'ц'],
    ['ch', 'ч'],
    ['sh', 'ш'],
    ['yu', 'ю'],
    ['iu', 'ю'],
    ['ya', 'я'],
    ['ia', 'я'],
    ['ye', 'є'],
    ['ie', 'є'],
    ['yi', 'ї'],
    ['yo', 'йо']
  ]

  for (const [from, to] of digraphs) {
    value = value.split(from).join(to)
  }

  const chars = {
    a: 'а', b: 'б', c: 'к', d: 'д', e: 'е', f: 'ф', g: 'г', h: 'х', i: 'і', j: 'й',
    k: 'к', l: 'л', m: 'м', n: 'н', o: 'о', p: 'п', q: 'к', r: 'р', s: 'с', t: 'т',
    u: 'у', v: 'в', w: 'в', x: 'кс', y: 'и', z: 'з'
  }

  return value
    .split('')
    .map((ch) => chars[ch] ?? ch)
    .join('')
}



