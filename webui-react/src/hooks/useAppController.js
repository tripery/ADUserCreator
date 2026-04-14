import { startTransition, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { API_BASE, BRAND_STORAGE_KEY, COLORS, INITIAL_PDF_LOGS, SIDEBAR_MODE_STORAGE_KEY } from '../config/appConfig'
import { WORKFLOW_STEPS } from '../config/uiContent'
import {
  addLog,
  buildOuTree,
  filterOuTree,
  findSheetWithColumn,
  getBrandPalette,
  getHeroPalette,
  getInitialChoice,
  getInitialTheme,
  getSidebarPalette,
  makeLog,
  normalizeExcelRows,
  readApiResponse,
} from '../utils/appUtils'

export function useAppController() {
  const [activeTab, setActiveTab] = useState('creator')
  const [theme, setTheme] = useState(getInitialTheme)
  const [accent, setAccent] = useState('blue')
  const [sidebarMode, setSidebarMode] = useState(() =>
    getInitialChoice(SIDEBAR_MODE_STORAGE_KEY, 'themed', ['themed', 'standard']),
  )
  const [brandAccent, setBrandAccent] = useState(() =>
    getInitialChoice(BRAND_STORAGE_KEY, 'default', ['default', ...Object.keys(COLORS)]),
  )
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
  const [createErrors, setCreateErrors] = useState([])
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
  const sidebarPalette = getSidebarPalette(theme, brandAccent, sidebarMode)
  const brandPalette = getBrandPalette(theme, brandAccent)
  const heroPalette = getHeroPalette(theme, accent)
  const latestLog = logs[logs.length - 1]
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
      { label: 'Помилок preview', value: previewErrors.length, tone: 'red' },
      { label: 'Створено', value: createResults.length, tone: 'amber' },
    ],
    [sourceUsers.length, previewRows.length, previewErrors.length, createResults.length],
  )

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    try {
      window.localStorage.setItem('adusercreator-theme', theme)
    } catch {}
  }, [theme])

  useEffect(() => {
    try {
      window.localStorage.setItem(SIDEBAR_MODE_STORAGE_KEY, sidebarMode)
    } catch {}
  }, [sidebarMode])

  useEffect(() => {
    try {
      window.localStorage.setItem(BRAND_STORAGE_KEY, brandAccent)
    } catch {}
  }, [brandAccent])

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
          if (options.ous?.length) {
            setSelectedOu((prev) => prev || options.ous[0].distinguishedName)
          }
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
    if (isOuDropdownOpen && ouSearchInputRef.current) {
      ouSearchInputRef.current.focus()
    }
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
      setCreateErrors([])
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
    setCreateErrors([])
    addLog(setLogs, 'INFO', 'Вибір файлу очищено')
  }

  async function requestPreview(users = sourceUsers, domain = domainSuffix, ou = selectedOu) {
    if (!users.length) return addLog(setLogs, 'WARN', 'Немає даних для preview')
    if (!domain) return addLog(setLogs, 'WARN', 'Немає домену для preview')
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
      setCreateErrors(data.errors ?? [])
      addLog(setLogs, 'OK', `${dryRun ? 'Dry-run' : 'Створення'} завершено: ${(data.created ?? []).length} успішно, ${(data.errors ?? []).length} помилок`)
      if (!dryRun) setActiveTab('monitoring')
    } catch (error) {
      addLog(setLogs, 'ERROR', `Помилка create: ${error.message}`)
    } finally {
      setIsCreating(false)
    }
  }

  const appStyle = {
    '--blue': accentColors.value,
    '--blue-soft': accentColors.soft,
    '--sidebar-bg': sidebarPalette.bg,
    '--sidebar-hover': sidebarPalette.hover,
    '--sidebar-active': sidebarPalette.active,
    '--sidebar-active-border': sidebarPalette.activeBorder,
    '--sidebar-active-text': sidebarPalette.activeText,
    '--sidebar-note-bg': sidebarPalette.noteBg,
    '--sidebar-note-border': sidebarPalette.noteBorder,
    '--sidebar-text': sidebarPalette.text,
    '--sidebar-muted': sidebarPalette.muted,
    '--sidebar-divider': sidebarPalette.divider,
    '--brand-bg': brandPalette.bg,
    '--hero-bg': heroPalette.bg,
    '--hero-text': heroPalette.text,
    '--hero-muted': heroPalette.muted,
    '--hero-glow': heroPalette.glow,
  }

  const sharedPageProps = {
    activeTab,
    apiStatus,
    createErrors,
    createResults,
    domainSuffix,
    expandedOuNodes,
    fileInputRef,
    fileName,
    filteredOuTree,
    groupInput,
    groupOptions,
    handleFileSelected,
    isCreating,
    isOuDropdownOpen,
    isPreviewLoading,
    latestLog,
    logs,
    ouDropdownRef,
    ouSearch,
    ouSearchInputRef,
    passwordNeverExpires,
    previewErrors,
    previewRows,
    removeGroup,
    requestPreview,
    selectedGroups,
    selectedOu,
    selectedOuName,
    setDomainSuffix,
    setGroupInput,
    setIsOuDropdownOpen,
    setLogs,
    setPasswordNeverExpires,
    setSelectedOu,
    setOuSearch,
    sourceUsers,
    toggleOuNode,
    workflowSteps: WORKFLOW_STEPS,
    addGroup,
    clearFile,
    createUsers,
  }

  return {
    activeTab,
    appStyle,
    apiStatus,
    brandAccent,
    brandPalette,
    pdfLogs,
    setAccent,
    accent,
    setBrandAccent,
    setPdfLogs,
    setSidebarMode,
    setTheme,
    sharedPageProps,
    sidebarMode,
    summaryCards,
    theme,
    latestLog,
    logs,
    sourceUsers,
    selectedOuName,
    domainSuffix,
    setActiveTab,
    createUsers,
    isCreating,
    setLogs,
  }
}
