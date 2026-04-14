import React from 'react'
import { OuTreeNode, PreviewCard, ResultsCard } from '../components/ui'

export function UsersPage(props) {
  return (
    <section className="tab-grid">
      <div className="card card-elevated">
        <div className="card-lead">
          <div>
            <div className="section-kicker">Панель керування Active Directory</div>
            <h2>Завантажте Excel і налаштуйте контекст створення</h2>
          </div>
          <div className="card-lead-actions">
            <div className="pill-info">{props.fileName}</div>
            <button className="btn btn-success" type="button" disabled={props.isCreating} onClick={() => props.createUsers()}>
              {props.isCreating ? 'Створення...' : 'Створити користувачів'}
            </button>
          </div>
        </div>

        <div className="field-block">
          <label className="label">Excel файл (*.xlsx)</label>
          <div className="file-row">
            <label className="btn btn-primary file-pick" htmlFor="excelUsersFile">
              Вибрати файл
            </label>
            <input id="excelUsersFile" ref={props.fileInputRef} type="file" accept=".xlsx,.xls" hidden onChange={props.handleFileSelected} />
            <div className="file-pill">{props.fileName}</div>
            <button className="btn btn-danger" type="button" onClick={props.clearFile}>
              Видалити
            </button>
          </div>
        </div>

        <div className="grid-2">
          <div className="field-block">
            <label className="label" htmlFor="usersDomainSuffix">
              Домен для UPN / E-mail
            </label>
            <input
              id="usersDomainSuffix"
              className="text-input"
              value={props.domainSuffix}
              onChange={(event) => props.setDomainSuffix(event.target.value)}
              placeholder="donnu.edu.ua"
            />
          </div>

          <div className="field-block">
            <label className="label" htmlFor="usersOuSelect">
              Виберіть OU
            </label>
            <div id="usersOuSelect" ref={props.ouDropdownRef} className="ou-tree-select">
              <button type="button" className="ou-tree-trigger" onClick={() => props.setIsOuDropdownOpen((prev) => !prev)}>
                <span className="ou-tree-trigger-text">{props.selectedOuName || 'Оберіть OU...'}</span>
                <span className="ou-tree-caret">{props.isOuDropdownOpen ? '▴' : '▾'}</span>
              </button>

              <div className={`ou-tree-dropdown ${props.isOuDropdownOpen ? 'open' : ''}`}>
                <input
                  ref={props.ouSearchInputRef}
                  className="text-input ou-tree-search"
                  value={props.ouSearch}
                  onChange={(event) => props.setOuSearch(event.target.value)}
                  placeholder="Пошук OU або DN..."
                />

                <div className="ou-tree-list" role="tree">
                  {!props.filteredOuTree.roots.length && (
                    <div className="ou-tree-empty">{props.ouSearch.trim() ? 'Нічого не знайдено' : 'Список OU порожній'}</div>
                  )}

                  {props.filteredOuTree.roots.map((node) => (
                    <OuTreeNode
                      key={node.dn}
                      node={node}
                      depth={0}
                      expandedOuNodes={props.expandedOuNodes}
                      selectedOu={props.selectedOu}
                      isSearchMode={Boolean(props.ouSearch.trim())}
                      onToggle={props.toggleOuNode}
                      onSelect={(dn) => {
                        props.setSelectedOu(dn)
                        props.setIsOuDropdownOpen(false)
                        if (props.sourceUsers.length) props.requestPreview(props.sourceUsers, props.domainSuffix, dn)
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
              <select
                className="chip-select"
                defaultValue=""
                onChange={(event) => {
                  if (event.target.value) props.addGroup(event.target.value)
                  event.target.value = ''
                }}
              >
                <option value="">Вибрати групу з AD</option>
                {props.groupOptions.slice(0, 400).map((group) => (
                  <option key={group.samAccountName} value={group.samAccountName}>
                    {group.name} ({group.samAccountName})
                  </option>
                ))}
              </select>

              <input
                className="chip-text"
                value={props.groupInput}
                onChange={(event) => props.setGroupInput(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault()
                    props.addGroup(props.groupInput)
                    props.setGroupInput('')
                  }
                }}
                placeholder="Введіть SamAccountName групи"
              />
            </div>

            <button
              className="add-chip"
              type="button"
              onClick={() => {
                props.addGroup(props.groupInput)
                props.setGroupInput('')
              }}
            >
              ❯
            </button>
          </div>

          <div className="chip-list">
            {props.selectedGroups.map((group) => (
              <div className="chip" key={group}>
                <span>{group}</span>
                <button type="button" onClick={() => props.removeGroup(group)}>
                  ×
                </button>
              </div>
            ))}
          </div>
        </div>

        <div className="field-block inline-actions">
          <label className="checkbox-row">
            <input
              type="checkbox"
              checked={props.passwordNeverExpires}
              onChange={(event) => props.setPasswordNeverExpires(event.target.checked)}
            />
            <span>Пароль не має терміну дії</span>
          </label>

          <button
            className="btn btn-ghost strong"
            type="button"
            onClick={() => props.requestPreview()}
            disabled={props.isPreviewLoading || !props.sourceUsers.length}
          >
            {props.isPreviewLoading ? 'Оновлення preview...' : 'Оновити preview'}
          </button>

          <button
            className="btn btn-ghost"
            type="button"
            onClick={() => props.createUsers({ dryRun: true })}
            disabled={props.isCreating || !props.sourceUsers.length}
          >
            Dry-run create
          </button>
        </div>
      </div>

      <PreviewCard previewRows={props.previewRows} previewErrors={props.previewErrors} />
      <ResultsCard createResults={props.createResults} createErrors={props.createErrors} />
    </section>
  )
}
