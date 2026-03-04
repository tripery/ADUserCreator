const sampleUsers = [
  {
    fullName: "Іванчук Владислава Анатоліївна",
    login: "ivvanchuk",
    email: "ivanchuk@donnu.edu.ua",
    groups: ["Студенти"]
  },
  {
    fullName: "Козленко Андрій Миколайович",
    login: "a.kozlenko",
    email: "kozlenko.a@donnu.edu.ua",
    groups: ["Студенти"]
  },
  {
    fullName: "Кучер Олександр Павлович",
    login: "o.kucher",
    email: "kucher.o@donnu.edu.ua",
    groups: ["Студенти", "Група КН-22"]
  },
  {
    fullName: "Осадчук Наталія Ігорівна",
    login: "n.osadchuk",
    email: "n.osadch.n@donnu.edu.ua",
    groups: ["Студенти", "Група КН-22"]
  }
];

const state = {
  groups: ["Студенти", "Група КН-22"],
  fileName: ""
};

const el = {
  fileInput: document.getElementById("excelFile"),
  fileName: document.getElementById("fileName"),
  clearFile: document.getElementById("clearFile"),
  groupInput: document.getElementById("groupInput"),
  groupPreset: document.getElementById("groupPreset"),
  addGroupBtn: document.getElementById("addGroupBtn"),
  chipList: document.getElementById("chipList"),
  previewTableBody: document.getElementById("previewTableBody"),
  previewMeta: document.getElementById("previewMeta"),
  createUsersBtn: document.getElementById("createUsersBtn"),
  clearLogBtn: document.getElementById("clearLogBtn"),
  logBox: document.getElementById("logBox"),
  ouSelect: document.getElementById("ouSelect"),
  neverExpire: document.getElementById("neverExpire")
};

function log(message, level = "INFO") {
  const stamp = new Date().toLocaleTimeString("uk-UA", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  const line = `[${stamp}] [${level}] ${message}`;
  el.logBox.textContent += (el.logBox.textContent ? "\n" : "") + line;
  el.logBox.scrollTop = el.logBox.scrollHeight;
}

function renderPreview(users) {
  el.previewTableBody.innerHTML = "";

  users.forEach((user) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(user.fullName)}</td>
      <td>${escapeHtml(user.login)}</td>
      <td>${escapeHtml(user.email)}</td>
      <td>${escapeHtml(user.groups.join(", "))}</td>
    `;
    el.previewTableBody.appendChild(tr);
  });

  el.previewMeta.textContent = `Показано 1-${users.length} з ${users.length} користувачів`;
}

function renderChips() {
  el.chipList.innerHTML = "";

  state.groups.forEach((group) => {
    const chip = document.createElement("div");
    chip.className = "chip";
    chip.innerHTML = `<span>${escapeHtml(group)}</span><button type="button" aria-label="Видалити ${escapeHtml(group)}">×</button>`;
    chip.querySelector("button").addEventListener("click", () => {
      state.groups = state.groups.filter((g) => g !== group);
      renderChips();
      log(`Групу видалено: ${group}`, "OK");
    });
    el.chipList.appendChild(chip);
  });
}

function addGroup(name) {
  const trimmed = (name || "").trim();
  if (!trimmed) return;
  if (state.groups.includes(trimmed)) {
    log(`Група вже додана: ${trimmed}`, "WARN");
    return;
  }

  state.groups.push(trimmed);
  renderChips();
  log(`Групу додано: ${trimmed}`, "OK");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

el.fileInput.addEventListener("change", (event) => {
  const file = event.target.files?.[0];
  state.fileName = file ? file.name : "";
  el.fileName.textContent = state.fileName || "Файл не вибрано";

  if (file) {
    log(`Вибрано Excel файл: ${file.name}`, "OK");
    log("Зараз використовується демо-прев’ю. Наступний крок: підв’язати реальний parser/API.", "INFO");
  }
});

el.clearFile.addEventListener("click", () => {
  el.fileInput.value = "";
  state.fileName = "";
  el.fileName.textContent = "Файл не вибрано";
  log("Вибір файлу очищено", "INFO");
});

el.addGroupBtn.addEventListener("click", () => {
  addGroup(el.groupInput.value);
  el.groupInput.value = "";
  el.groupInput.focus();
});

el.groupInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    addGroup(el.groupInput.value);
    el.groupInput.value = "";
  }
});

el.groupPreset.addEventListener("change", () => {
  const value = el.groupPreset.value;
  if (value === "students") addGroup("Студенти");
  if (value === "staff") addGroup("Працівники");
});

el.createUsersBtn.addEventListener("click", () => {
  if (!state.fileName) {
    log("Спочатку виберіть Excel-файл", "ERROR");
    return;
  }

  log(`СТАРТ (demo): OU='${el.ouSelect.value}', групи='${state.groups.join(", ")}', пароль без строку дії='${el.neverExpire.checked}'`, "INFO");
  log("Це UI-прототип. Інтеграція з PowerShell/AD ще не підключена.", "WARN");
});

el.clearLogBtn.addEventListener("click", () => {
  el.logBox.textContent = "";
});

renderPreview(sampleUsers);
renderChips();
log("Web UI прототип ініціалізовано", "OK");
log("Відкрийте webui/index.html у браузері для перегляду", "INFO");
