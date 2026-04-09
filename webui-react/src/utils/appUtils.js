import * as XLSX from 'xlsx'
import { BRAND_STORAGE_KEY, COLORS, SIDEBAR_MODE_STORAGE_KEY, THEME_STORAGE_KEY } from '../config/appConfig'

export function makeLog(level, message) {
  return {
    level,
    message,
    stamp: new Date().toLocaleTimeString('uk-UA', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    }),
  }
}

export function addLog(setLogs, level, message) {
  setLogs((prev) => [...prev, makeLog(level, message)])
}

export async function readApiResponse(response, operationName) {
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

export function getInitialTheme() {
  if (typeof window === 'undefined') return 'light'
  try {
    const savedTheme = window.localStorage.getItem(THEME_STORAGE_KEY)
    if (savedTheme === 'light' || savedTheme === 'dark') return savedTheme
  } catch {}
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

export function getInitialChoice(storageKey, fallback, allowedValues) {
  if (typeof window === 'undefined') return fallback
  try {
    const savedValue = window.localStorage.getItem(storageKey)
    if (savedValue && allowedValues.includes(savedValue)) return savedValue
  } catch {}
  return fallback
}

export function getSidebarPalette(theme, accentKey, mode) {
  const isDark = theme === 'dark'
  const isThemed = mode !== 'standard'

  if (!isThemed) {
    return isDark
      ? {
          bg: 'linear-gradient(180deg, #07162f 0%, #081b3f 100%)',
          text: '#e7efff',
          muted: 'rgba(231, 239, 255, 0.72)',
          divider: 'rgba(255, 255, 255, 0.08)',
          hover: 'rgba(255, 255, 255, 0.08)',
          active: 'linear-gradient(90deg, rgba(103, 156, 255, 0.32), rgba(103, 156, 255, 0.12))',
          activeBorder: 'rgba(255, 255, 255, 0.1)',
          activeText: '#f7fbff',
          noteBg: 'rgba(255, 255, 255, 0.06)',
          noteBorder: 'rgba(255, 255, 255, 0.08)',
        }
      : {
          bg: 'linear-gradient(180deg, #eef4ff 0%, #dde7f6 100%)',
          text: '#25405f',
          muted: '#5f7391',
          divider: 'rgba(152, 170, 197, 0.35)',
          hover: 'rgba(45, 104, 214, 0.08)',
          active: 'linear-gradient(90deg, rgba(45, 104, 214, 0.2), rgba(45, 104, 214, 0.08))',
          activeBorder: 'rgba(45, 104, 214, 0.18)',
          activeText: '#163563',
          noteBg: 'rgba(255, 255, 255, 0.58)',
          noteBorder: 'rgba(152, 170, 197, 0.24)',
        }
  }

  const resolvedAccent = COLORS[accentKey] ? accentKey : 'blue'
  const base = COLORS[resolvedAccent].value

  if (isDark) {
    return {
      bg: `linear-gradient(180deg, ${shadeHex(base, -60)} 0%, ${shadeHex(base, -38)} 100%)`,
      text: '#edf6ff',
      muted: 'rgba(237, 246, 255, 0.74)',
      divider: 'rgba(255, 255, 255, 0.08)',
      hover: 'rgba(255, 255, 255, 0.1)',
      active: 'linear-gradient(90deg, rgba(255, 255, 255, 0.24), rgba(255, 255, 255, 0.1))',
      activeBorder: 'rgba(255, 255, 255, 0.18)',
      activeText: '#ffffff',
      noteBg: 'rgba(255, 255, 255, 0.08)',
      noteBorder: 'rgba(255, 255, 255, 0.12)',
    }
  }

  const lightStart = mixHex(base, '#ffffff', 0.82)
  const lightEnd = mixHex(base, '#ffffff', 0.68)

  return {
    bg: `linear-gradient(180deg, ${lightStart} 0%, ${lightEnd} 100%)`,
    text: '#24415f',
    muted: '#5d7291',
    divider: hexToRgba(base, 0.16),
    hover: 'rgba(255, 255, 255, 0.38)',
    active: `linear-gradient(90deg, ${base} 0%, ${shadeHex(base, 12)} 100%)`,
    activeBorder: hexToRgba(base, 0.18),
    activeText: '#ffffff',
    noteBg: 'rgba(255, 255, 255, 0.56)',
    noteBorder: hexToRgba(base, 0.18),
  }
}

export function getBrandPalette(theme, accentKey) {
  if (!accentKey || accentKey === 'default' || !COLORS[accentKey]) {
    return {
      bg:
        theme === 'dark'
          ? 'linear-gradient(135deg, #0d2a57 0%, #1b4f9a 100%)'
          : 'linear-gradient(135deg, #17438f 0%, #2d6ce1 100%)',
    }
  }

  const base = COLORS[accentKey].value
  const start = shadeHex(base, theme === 'dark' ? -26 : -18)
  const end = shadeHex(base, theme === 'dark' ? 6 : 0)

  return {
    bg: `linear-gradient(135deg, ${start} 0%, ${end} 100%)`,
  }
}

export function getHeroPalette(theme, accentKey) {
  const resolvedAccent = COLORS[accentKey] ? accentKey : 'blue'
  const base = COLORS[resolvedAccent].value

  if (theme === 'dark') {
    return {
      bg: `linear-gradient(135deg, ${shadeHex(base, -62)} 0%, ${shadeHex(base, -28)} 56%, ${shadeHex(base, -6)} 100%)`,
      text: '#edf6ff',
      muted: 'rgba(237, 246, 255, 0.76)',
      glow: hexToRgba(base, 0.2),
    }
  }

  return {
    bg: `linear-gradient(135deg, ${shadeHex(base, -18)} 0%, ${base} 56%, ${shadeHex(base, 22)} 100%)`,
    text: '#eef6ff',
    muted: 'rgba(238, 246, 255, 0.82)',
    glow: hexToRgba(base, 0.18),
  }
}

function shadeHex(hex, amount) {
  const normalized = String(hex || '').replace('#', '')
  if (normalized.length !== 6) return hex

  const parts = [0, 2, 4].map((index) => {
    const value = parseInt(normalized.slice(index, index + 2), 16)
    const next = Math.max(0, Math.min(255, value + amount))
    return next.toString(16).padStart(2, '0')
  })

  return `#${parts.join('')}`
}

function mixHex(hexA, hexB, weight) {
  const rgbA = hexToRgb(hexA)
  const rgbB = hexToRgb(hexB)
  if (!rgbA || !rgbB) return hexA

  const blend = ['r', 'g', 'b'].map((channel) => {
    const value = Math.round(rgbA[channel] * (1 - weight) + rgbB[channel] * weight)
    return value.toString(16).padStart(2, '0')
  })

  return `#${blend.join('')}`
}

function hexToRgb(hex) {
  const normalized = String(hex || '').replace('#', '')
  if (normalized.length !== 6) return null

  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16),
  }
}

function hexToRgba(hex, alpha) {
  const rgb = hexToRgb(hex)
  if (!rgb) return `rgba(0, 0, 0, ${alpha})`
  return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})`
}

export function buildOuTree(ouOptions) {
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
      children: [],
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
  return eq === -1 ? firstPart || dn : `${firstPart.slice(eq + 1)} (${firstPart.slice(0, eq).toUpperCase()})`
}

function sortTreeNodes(nodes) {
  nodes.sort((a, b) => a.label.localeCompare(b.label, 'uk-UA', { sensitivity: 'base' }))
  for (const node of nodes) if (node.children.length) sortTreeNodes(node.children)
}

export function filterOuTree(tree, search) {
  const query = normalizeSearchValue(search)
  if (!query) return tree
  return {
    roots: tree.roots.map((node) => filterOuNode(node, query)).filter(Boolean),
    parentByDn: tree.parentByDn,
  }
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

  return Array.from(new Set([query, translitUaToLat(query), translitLatToUa(query)])).some(
    (variant) => variant && haystack.includes(variant),
  )
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

export function findSheetWithColumn(workbook, expectedHeader) {
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

export function normalizeExcelRows(rows) {
  return rows
    .map((row, index) => ({ ...buildUserFromExcelRow(row), sourceRow: index + 2 }))
    .filter((row) => row.fullName)
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
  return parts.length === 0
    ? { firstName: '', lastName: '', middleName: '' }
    : {
        lastName: parts[0] || '',
        firstName: parts[1] || '',
        middleName: parts.slice(2).join(' '),
      }
}

function generateSamAccountName(firstName, lastName) {
  const first = translitUaToLat(firstName || '')
  const last = translitUaToLat(lastName || '')
  const cleaned = `${first ? first[0] : ''}.${last}`
    .toLowerCase()
    .replace(/[^a-z0-9._-]/g, '')
    .replace(/^[._-]+|[._-]+$/g, '')

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

  while (required.length < length) {
    required.push(all[Math.floor(Math.random() * all.length)])
  }

  for (let index = required.length - 1; index > 0; index -= 1) {
    const target = Math.floor(Math.random() * (index + 1))
    const temp = required[index]
    required[index] = required[target]
    required[target] = temp
  }

  return required.join('')
}

function translitUaToLat(input) {
  const map = {
    а: 'a',
    б: 'b',
    в: 'v',
    г: 'h',
    ґ: 'g',
    д: 'd',
    е: 'e',
    є: 'ie',
    ж: 'zh',
    з: 'z',
    и: 'y',
    і: 'i',
    ї: 'i',
    й: 'i',
    к: 'k',
    л: 'l',
    м: 'm',
    н: 'n',
    о: 'o',
    п: 'p',
    р: 'r',
    с: 's',
    т: 't',
    у: 'u',
    ф: 'f',
    х: 'kh',
    ц: 'ts',
    ч: 'ch',
    ш: 'sh',
    щ: 'shch',
    ь: '',
    ю: 'iu',
    я: 'ia',
    "'": '',
    '’': '',
    '-': '-',
    ' ': '',
  }

  return String(input || '')
    .toLowerCase()
    .split('')
    .map((ch) => map[ch] ?? ch)
    .join('')
}

function translitLatToUa(input) {
  let value = String(input || '').toLowerCase()

  for (const [from, to] of [
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
    ['yo', 'йо'],
  ]) {
    value = value.split(from).join(to)
  }

  const chars = {
    a: 'а',
    b: 'б',
    c: 'к',
    d: 'д',
    e: 'е',
    f: 'ф',
    g: 'г',
    h: 'х',
    i: 'і',
    j: 'й',
    k: 'к',
    l: 'л',
    m: 'м',
    n: 'н',
    o: 'о',
    p: 'п',
    q: 'к',
    r: 'р',
    s: 'с',
    t: 'т',
    u: 'у',
    v: 'в',
    w: 'в',
    x: 'кс',
    y: 'и',
    z: 'з',
  }

  return value
    .split('')
    .map((ch) => chars[ch] ?? ch)
    .join('')
}

export const INITIAL_THEME_OPTIONS = ['light', 'dark']
export const INITIAL_SIDEBAR_OPTIONS = ['themed', 'standard']
export const INITIAL_BRAND_OPTIONS = ['default', ...Object.keys(COLORS)]
export const STORAGE_KEYS = {
  THEME_STORAGE_KEY,
  SIDEBAR_MODE_STORAGE_KEY,
  BRAND_STORAGE_KEY,
}
