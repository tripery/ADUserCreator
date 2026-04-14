import React from 'react'
import { SessionCard } from '../components/ui'

const TOPICS = [
  {
    title: 'Архітектура локальної мережі',
    description:
      'LAN у середовищі MS Windows базується на ролях серверів, централізованому керуванні користувачами, політиках доступу та надійній взаємодії клієнтів із доменними службами.',
  },
  {
    title: 'Active Directory Domain Services',
    description:
      'AD DS формує ієрархію домену, OU, користувачів, комп’ютерів і груп. Це основа для делегування прав, централізованої автентифікації та впорядкованого адміністрування.',
  },
  {
    title: 'DNS і DHCP',
    description:
      'DNS забезпечує резолюцію імен, а DHCP автоматизує видачу IP-конфігурацій. Разом ці служби зменшують ручні помилки та спрощують підтримку мережевої інфраструктури.',
  },
  {
    title: 'Group Policy',
    description:
      'Групові політики дозволяють централізовано задавати правила безпеки, параметри середовища користувача, конфігурацію робочих станцій і типові адміністративні обмеження.',
  },
  {
    title: 'Безпека та контроль змін',
    description:
      'Якісне адміністрування Windows-мережі поєднує облік подій, резервування, контроль доступу, аудит змін і стандартизовані сценарії супроводу домену.',
  },
  {
    title: 'Лабораторні стенди й віртуалізація',
    description:
      'Hyper-V та інші платформи віртуалізації дають змогу відпрацьовувати сценарії LAN-адміністрування без ризику для продуктивного середовища та швидко моделювати інфраструктуру.',
  },
]

const SOURCES = [
  {
    label: 'Основи адміністрування LAN у середовищі MS Windows',
    meta: 'Навчальний посібник, Видавництво Львівської політехніки',
    href: 'https://vlp.com.ua/node/11348',
  },
  {
    label: 'Active Directory Domain Services overview',
    meta: 'Microsoft Learn',
    href: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/get-started/virtual-dc/active-directory-domain-services-overview',
  },
  {
    label: 'Group Policy overview',
    meta: 'Microsoft Learn',
    href: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/group-policy/group-policy-overview',
  },
  {
    label: 'DNS Architecture in Windows Server',
    meta: 'Microsoft Learn',
    href: 'https://learn.microsoft.com/en-us/windows-server/networking/dns/dns-architecture',
  },
  {
    label: 'DHCP overview for Windows Server',
    meta: 'Microsoft Learn',
    href: 'https://learn.microsoft.com/en-us/windows-server/networking/technologies/dhcp/dhcp-top',
  },
  {
    label: 'Hyper-V overview',
    meta: 'Microsoft Learn',
    href: 'https://learn.microsoft.com/en-us/virtualization/hyper-v-on-windows/about/',
  },
]

export function HomePage(props) {
  return (
    <section className="workspace-grid">
      <div className="workspace-main">
        <div className="card card-elevated landing-card-shell">
          <div className="card-lead">
            <div>
              <div className="section-kicker">Візитка системи</div>
              <h2>Основи адміністрування LAN у середовищі MS Windows</h2>
            </div>
          </div>

          <p className="landing-intro">
            Головна сторінка подає стислий огляд тем, які формують базу адміністрування локальних мереж у середовищі Windows: служби
            каталогу, мережеві ролі, політики, безпека та кероване впровадження змін.
          </p>

          <div className="landing-grid">
            {TOPICS.map((topic) => (
              <article key={topic.title} className="landing-topic-card">
                <h3>{topic.title}</h3>
                <p>{topic.description}</p>
              </article>
            ))}
          </div>

          <div className="landing-sources">
            <div className="section-kicker">Джерела</div>
            <h3>Матеріали для поглибленого вивчення</h3>
            <div className="sources-list">
              {SOURCES.map((source) => (
                <a key={source.href} className="source-link-card" href={source.href} target="_blank" rel="noreferrer">
                  <strong>{source.label}</strong>
                  <span>{source.meta}</span>
                </a>
              ))}
            </div>
          </div>
        </div>
      </div>

      <aside className="workspace-side">
        <SessionCard
          apiStatus={props.apiStatus}
          selectedOuName={props.selectedOuName}
          domainSuffix={props.domainSuffix}
          latestLog={props.latestLog}
        />
      </aside>
    </section>
  )
}
