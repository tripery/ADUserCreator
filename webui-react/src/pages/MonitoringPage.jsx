import React from 'react'
import { LogCard, SessionCard } from '../components/ui'

export function MonitoringPage({ apiStatus, logs, selectedOuName, domainSuffix, latestLog, setLogs, summaryCards }) {
  return (
    <section className="tab-grid">
      <section className="hero-panel">
        <div className="hero-copy">
          <h2>Акуратний фронтенд для масового створення користувачів без ручної рутини.</h2>
          <p>Завантажуйте Excel, одразу бачте готові логіни та пошту, запускайте dry-run і лише потім створюйте облікові записи в AD.</p>
        </div>

        <div className="hero-stats">
          {summaryCards.map((card) => (
            <div key={card.label} className={`hero-stat ${card.tone}`}>
              <div className="hero-stat-value">{card.value}</div>
              <div className="hero-stat-label">{card.label}</div>
            </div>
          ))}
        </div>
      </section>
      <LogCard logs={logs} onClear={() => setLogs([])} />
      <SessionCard apiStatus={apiStatus} selectedOuName={selectedOuName} domainSuffix={domainSuffix} latestLog={latestLog} />
    </section>
  )
}
