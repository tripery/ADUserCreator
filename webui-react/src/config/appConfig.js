export const API_BASE = import.meta.env.VITE_API_BASE || ''
export const THEME_STORAGE_KEY = 'adusercreator-theme'
export const SIDEBAR_MODE_STORAGE_KEY = 'adusercreator-sidebar-mode'
export const BRAND_STORAGE_KEY = 'adusercreator-brand-accent'

export const COLORS = {
  blue: { value: '#2d68d6', soft: 'rgba(45, 104, 214, 0.12)' },
  emerald: { value: '#21916a', soft: 'rgba(33, 145, 106, 0.13)' },
  purple: { value: '#7c55d8', soft: 'rgba(124, 85, 216, 0.14)' },
  amber: { value: '#bc7a19', soft: 'rgba(188, 122, 25, 0.14)' },
  rose: { value: '#f23d68', soft: 'rgba(242, 61, 104, 0.14)' },
}

export const COLOR_LABELS = {
  blue: 'Blue',
  emerald: 'Emerald',
  purple: 'Purple',
  amber: 'Amber',
  rose: 'Rose',
}

export const INITIAL_PDF_LOGS = [
  { id: 'pdf-1', name: 'passwords_credentials_2026-03-17.pdf', size: '142 KB', date: '2026-03-17 13:30', users: 12 },
  { id: 'pdf-2', name: 'it_department_logins.pdf', size: '89 KB', date: '2026-03-16 10:15', users: 5 },
]
