import React from 'react'

export function BrandMarkIcon({ className = '' }) {
  return (
    <svg viewBox="0 0 48 48" aria-hidden="true" className={className}>
      <defs>
        <linearGradient id="brandMarkGlow" x1="0%" x2="100%" y1="0%" y2="100%">
          <stop offset="0%" stopColor="currentColor" stopOpacity="0.95" />
          <stop offset="100%" stopColor="currentColor" stopOpacity="0.55" />
        </linearGradient>
      </defs>
      <rect x="7" y="9" width="14" height="14" rx="4" fill="url(#brandMarkGlow)" />
      <rect x="27" y="9" width="14" height="14" rx="4" fill="currentColor" opacity="0.85" />
      <rect x="7" y="27" width="14" height="14" rx="4" fill="currentColor" opacity="0.75" />
      <path d="M28 31.5c0-2.5 2.1-4.5 4.6-4.5 1.7 0 3.2.9 4 2.3A4.5 4.5 0 1 1 34 39h-4.2A1.8 1.8 0 0 1 28 37.2v-5.7Z" fill="currentColor" />
      <path d="M17.5 17.5h13M17.5 30.5h7" stroke="currentColor" strokeLinecap="round" strokeOpacity="0.45" strokeWidth="2" />
    </svg>
  )
}

export function ThemeSparkIcon({ className = '' }) {
  return (
    <svg viewBox="0 0 48 48" aria-hidden="true" className={className}>
      <circle cx="24" cy="24" r="18" fill="none" stroke="currentColor" strokeOpacity="0.18" strokeWidth="2" />
      <path
        d="M24 11c2.7 0 5 2.2 5 5 0 1.1-.4 2.2-1 3.1l-2.2 3 2.3 2.3a5.4 5.4 0 0 1 1.6 3.8c0 3-2.4 5.5-5.4 5.5A5.4 5.4 0 0 1 19 28.2c0-1.4.5-2.7 1.4-3.7l2.2-2.3-2.1-2.9a5.1 5.1 0 0 1-1-3.2c0-2.8 2.2-5.1 5-5.1Z"
        fill="currentColor"
      />
      <path d="M34.5 13.5 36 10m2.5 8.5 3.5-1.5M11.5 18.5 8 17m4-7 1.5 3.5" stroke="currentColor" strokeLinecap="round" strokeWidth="2" />
    </svg>
  )
}
