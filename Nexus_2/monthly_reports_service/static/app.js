"use strict";

const state = {
  wordId: null, wordName: null,
  excelId: null, excelName: null,
  receptionId: null, receptionName: null,
  months: [],
  fillReception: false,
  dataQuality: null,
};

const SESSION_KEY = 'sonatrach_report_session_v2';

function $(id) { return document.getElementById(id); }

function showToast(msg, ms = 3000) {
  const t = $("toast");
  t.textContent = msg;
  t.classList.remove("hidden");
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => t.classList.add("hidden"), ms);
}

function showGlobalError(msg) {
  const el = $("global-error");
  if (!msg) { el.classList.add("hidden"); el.textContent = ""; return; }
  el.textContent = msg;
  el.classList.remove("hidden");
}

function normalizeStatus(status) {
  return status === "warn" ? "warning" : status;
}

function statusDotClass(status) {
  const s = normalizeStatus(status);
  return { ok: "dot-ok", warning: "dot-warning", error: "dot-error", missing: "dot-error", info: "dot-info" }[s] || "dot-info";
}

function badgeClass(status) {
  const s = normalizeStatus(status);
  return { ok: "badge-ok", warning: "badge-warning", error: "badge-error", missing: "badge-error", info: "badge-info" }[s] || "badge-info";
}

async function postJSON(url, body) {
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body || {}),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Erreur (${res.status})`);
  return data;
}

async function postForm(url, formData) {
  const res = await fetch(url, { method: "POST", body: formData });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Erreur (${res.status})`);
  return data;
}

function setStepState(stepNum, cls) {
  const el = document.querySelector(`.step[data-step="${stepNum}"]`);
  if (!el) return;
  el.classList.remove("active", "done");
  if (cls) el.classList.add(cls);
}

function updateActionButtons() {
  const ready = !!state.wordId && !!state.excelId && (!state.fillReception || !!state.receptionId);
  $("btn-preview").disabled = !state.excelId;
  $("btn-generate").disabled = !ready;
  $("btn-validate").disabled = !state.excelId;
}

function saveSession() {
  try {
    const sessionData = {
      wordId: state.wordId,
      wordName: state.wordName,
      excelId: state.excelId,
      excelName: state.excelName,
      receptionId: state.receptionId,
      receptionName: state.receptionName,
      months: state.months,
      fillReception: state.fillReception,
      selectedMonth: $("select-month").value,
      selectedCategories: selectedCategories(),
      timestamp: Date.now()
    };
    localStorage.setItem(SESSION_KEY, JSON.stringify(sessionData));
  } catch (e) {
    console.warn("Could not save session:", e);
  }
}

function restoreSession() {
  const saved = localStorage.getItem(SESSION_KEY);
  if (!saved) return;
  
  try {
    const sessionData = JSON.parse(saved);
    const age = Date.now() - (sessionData.timestamp || 0);
    if (age > 24 * 60 * 60 * 1000) {
      localStorage.removeItem(SESSION_KEY);
      return;
    }
    
    if (sessionData.wordId) {
      state.wordId = sessionData.wordId;
      state.wordName = sessionData.wordName;
      $("word-status").innerHTML = `<span class="file-ok">✓ <span class="file-name">${sessionData.wordName}</span> (restauré)</span>`;
      setStepState(1, "done");
    }
    
    if (sessionData.excelId) {
      state.excelId = sessionData.excelId;
      state.excelName = sessionData.excelName;
      state.months = sessionData.months || [];
      
      const sel = $("select-month");
      sel.innerHTML = "";
      state.months.forEach(m => {
        const opt = document.createElement("option");
        opt.value = `${m.year}-${m.month}`;
        opt.textContent = m.label;
        sel.appendChild(opt);
      });
      
      if (sessionData.selectedMonth) {
        sel.value = sessionData.selectedMonth;
      } else if (state.months.length > 0) {
        sel.selectedIndex = state.months.length - 1;
      }
      
      sel.disabled = state.months.length === 0;
      $("excel-status").innerHTML = `<span class="file-ok">✓ <span class="file-name">${sessionData.excelName}</span> (restauré)</span>`;
      setStepState(2, "done");
    }
    
    if (sessionData.receptionId) {
      state.receptionId = sessionData.receptionId;
      state.receptionName = sessionData.receptionName;
      state.fillReception = sessionData.fillReception;
      $("toggle-reception").checked = state.fillReception;
      $("reception-upload-wrap").classList.toggle("hidden", !state.fillReception);
      $("reception-status").innerHTML = `<span class="file-ok">✓ <span class="file-name">${sessionData.receptionName}</span> (restauré)</span>`;
      setStepState(3, "done");
    }
    
    if (sessionData.selectedCategories) {
      document.querySelectorAll(".cat-checkbox").forEach(cb => {
        cb.checked = sessionData.selectedCategories.includes(cb.value);
      });
    }
    
    updateActionButtons();
    showToast("✓ Session précédente restaurée");
  } catch (e) {
    console.error("Failed to restore session:", e);
    localStorage.removeItem(SESSION_KEY);
  }
}

function wireDropzone(zoneId, inputId, onFile) {
  const zone = $(zoneId);
  const input = $(inputId);
  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => {
    if (input.files && input.files[0]) onFile(input.files[0]);
  });
  ["dragenter", "dragover"].forEach(evt =>
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add("dragover"); })
  );
  ["dragleave", "drop"].forEach(evt =>
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove("dragover"); })
  );
  zone.addEventListener("drop", e => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) onFile(file);
  });
}

function renderValidationReport(report) {
  $("template-report").classList.remove("hidden");

  const badge = $("overall-badge");
  badge.textContent = { ok: "OK", warning: "Avertissements", error: "Erreurs" }[report.overall_status] || report.overall_status;
  badge.className = "badge " + badgeClass(report.overall_status);

  const vi1List = $("vi1-list");
  vi1List.innerHTML = "";
  Object.values(report.vi1).forEach(entry => {
    const li = document.createElement("li");
    const dot = `<span class="dot ${statusDotClass(entry.status)}"></span>`;
    let msg;
    if (entry.status === "missing") msg = "en-tête introuvable dans le document";
    else if (entry.status === "error") msg = "en-tête trouvée mais tableau introuvable";
    else if (entry.status === "warning") msg = `colonnes manquantes : ${entry.missing_columns.join(", ")}`;
    else msg = `OK — ${entry.data_rows} ligne(s) de données`;
    li.innerHTML = `${dot}<span><strong>${entry.label}</strong> — ${msg}</span>`;
    vi1List.appendChild(li);
  });

  const vi3List = $("vi3-list");
  vi3List.innerHTML = "";
  const vi3Entries = [
    ["VI-3-1 Terrain", report.vi3_terrain],
    ["VI-3-2/3 Stack", report.vi3_stack],
  ];
  vi3Entries.forEach(([label, entry]) => {
    const li = document.createElement("li");
    const dot = `<span class="dot ${statusDotClass(entry.status)}"></span>`;
    let msg;
    if (entry.status === "missing") msg = "en-tête introuvable dans le document";
    else if (entry.status === "error") msg = "en-tête trouvée mais aucun tableau associé";
    else if (entry.status === "warning") {
      const missing = entry.tables.flatMap(t => t.missing_columns);
      msg = `colonnes manquantes sur au moins un tableau : ${[...new Set(missing)].join(", ")}`;
    } else {
      const rows = entry.tables.reduce((sum, t) => sum + t.data_rows, 0);
      msg = `OK — ${entry.tables_found} tableau(x), ${rows} ligne(s) au total`;
    }
    li.innerHTML = `${dot}<span><strong>${label}</strong> — ${msg}</span>`;
    vi3List.appendChild(li);
  });

  $("btn-revalidate").classList.remove("hidden");
}

function renderDataQuality(quality) {
  state.dataQuality = quality;
  const panel = $("card-data-quality");
  panel.classList.remove("hidden");
  
  const scoreClass = quality.quality_score >= 80 ? "quality-good" : 
                     quality.quality_score >= 50 ? "quality-warn" : "quality-bad";
  const totalIssues = quality.empty_critical_fields.length + 
                      quality.date_anomalies.length + 
                      quality.duplicates.length + 
                      quality.warnings.length;
  
  let html = `
    <div class="quality-header">
      <div class="quality-score ${scoreClass}">
        <span class="score-value">${quality.quality_score}</span>
        <span class="score-label">Score qualité</span>
      </div>
      <div class="quality-stats">
        <div><strong>${quality.total_rows}</strong> lignes totales</div>
        <div><span class="issue-count ${totalIssues > 0 ? 'issue-count-red' : 'issue-count-green'}">${totalIssues}</span> problème(s) détecté(s)</div>
      </div>
    </div>
    <div class="quality-summary">
  `;
  
  if (quality.empty_critical_fields.length > 0) {
    html += `<div class="quality-item quality-bad">❌ <strong>${quality.empty_critical_fields.length}</strong> champ(s) critique(s) vide(s)</div>`;
  }
  if (quality.date_anomalies.length > 0) {
    html += `<div class="quality-item quality-warn">⚠️ <strong>${quality.date_anomalies.length}</strong> date(s) manquante(s)</div>`;
  }
  if (quality.duplicates.length > 0) {
    html += `<div class="quality-item quality-warn">🔁 <strong>${quality.duplicates.length}</strong> étude(s) dupliquée(s)</div>`;
  }
  if (quality.warnings.length > 0) {
    html += `<div class="quality-item quality-info">ℹ️ ${quality.warnings.length} avertissement(s) divers</div>`;
  }
  if (totalIssues === 0) {
    html += `<div class="quality-item quality-good">✅ Aucune anomalie détectée</div>`;
  }
  
  html += `</div>`;
  
  $("quality-summary").innerHTML = html;
  $("btn-view-issues").classList.toggle("hidden", totalIssues === 0);
}

function renderIssuesModal() {
  if (!state.dataQuality) return;
  const q = state.dataQuality;
  let html = "";
  
  if (q.empty_critical_fields.length > 0) {
    html += `<div class="issue-section">
      <h4>❌ Champs critiques vides (${q.empty_critical_fields.length})</h4>
      <table class="issue-table">
        <thead><tr><th>Ligne</th><th>Catégorie</th><th>Champ</th></tr></thead>
        <tbody>
          ${q.empty_critical_fields.map(i => `<tr><td>${i.row}</td><td>${i.category}</td><td>${i.field}</td></tr>`).join("")}
        </tbody>
      </table>
      <p class="issue-help">⚠️ Ces lignes ne peuvent pas être utilisées. Veuillez compléter le champ "Data" dans votre Excel.</p>
    </div>`;
  }
  
  if (q.date_anomalies.length > 0) {
    html += `<div class="issue-section">
      <h4>⚠️ Dates manquantes (${q.date_anomalies.length})</h4>
      <table class="issue-table">
        <thead><tr><th>Ligne</th><th>Catégorie</th><th>Problème</th></tr></thead>
        <tbody>
          ${q.date_anomalies.map(i => `<tr><td>${i.row}</td><td>${i.category}</td><td>${i.issue}</td></tr>`).join("")}
        </tbody>
      </table>
    </div>`;
  }
  
  if (q.duplicates.length > 0) {
    html += `<div class="issue-section">
      <h4>🔁 Études dupliquées (${q.duplicates.length})</h4>
      <table class="issue-table">
        <thead><tr><th>Étude</th><th>Catégorie</th><th>Lignes</th></tr></thead>
        <tbody>
          ${q.duplicates.map(d => `<tr><td>${d.study}</td><td>${d.category}</td><td>${d.rows.join(", ")}</td></tr>`).join("")}
        </tbody>
      </table>
    </div>`;
  }
  
  if (q.warnings.length > 0) {
    html += `<div class="issue-section">
      <h4>ℹ️ Avertissements</h4>
      <ul>${q.warnings.map(w => `<li>${w}</li>`).join("")}</ul>
    </div>`;
  }
  
  $("issues-body").innerHTML = html;
  $("issues-modal").classList.remove("hidden");
}

async function handleWordFile(file) {
  $("dz-word-text").textContent = `Envoi de "${file.name}"…`;
  $("word-status").innerHTML = "";
  const fd = new FormData();
  fd.append("word", file);
  try {
    const data = await postForm("/api/upload", fd);
    state.wordId = data.word_id;
    state.wordName = data.word_name;
    $("dz-word-text").textContent = "Cliquez ou glissez-déposez le fichier .docx ici";
    $("word-status").innerHTML = `<span class="file-ok">✓ <span class="file-name">${data.word_name}</span> chargé</span>`;
    if (data.template_error) {
      showGlobalError(data.template_error);
    } else if (data.template_validation) {
      showGlobalError(null);
      renderValidationReport(data.template_validation);
    }
    setStepState(1, "done");
    updateActionButtons();
    saveSession();
  } catch (err) {
    $("dz-word-text").textContent = "Cliquez ou glissez-déposez le fichier .docx ici";
    $("word-status").innerHTML = `<span class="file-err">✕ ${err.message}</span>`;
  }
}

$("btn-revalidate").addEventListener("click", async () => {
  if (!state.wordId) return;
  $("btn-revalidate").disabled = true;
  try {
    const report = await postJSON("/api/validate_template", { word_id: state.wordId });
    renderValidationReport(report);
    showToast("Modèle revérifié.");
  } catch (err) {
    showGlobalError(err.message);
  } finally {
    $("btn-revalidate").disabled = false;
  }
});

async function handleExcelFile(file) {
  $("dz-excel-text").textContent = `Envoi de "${file.name}"…`;
  $("excel-status").innerHTML = "";
  const fd = new FormData();
  fd.append("excel", file);
  try {
    const data = await postForm("/api/upload", fd);
    $("dz-excel-text").textContent = "Cliquez ou glissez-déposez le fichier Excel ici";
    if (data.excel_error) {
      $("excel-status").innerHTML = `<span class="file-err">✕ ${data.excel_error}</span>`;
      return;
    }
    state.excelId = data.excel_id;
    state.excelName = data.excel_name;
    state.months = data.months || [];
    $("excel-status").innerHTML =
      `<span class="file-ok">✓ <span class="file-name">${data.excel_name}</span> chargé — ` +
      `${state.months.length} mois disponible(s), ${(data.categories || []).length} section(s) détectée(s)</span>`;

    const sel = $("select-month");
    sel.innerHTML = "";
    if (state.months.length === 0) {
      sel.innerHTML = `<option value="">Aucune donnée datée trouvée</option>`;
      sel.disabled = true;
    } else {
      state.months.forEach(m => {
        const opt = document.createElement("option");
        opt.value = `${m.year}-${m.month}`;
        opt.textContent = m.label;
        sel.appendChild(opt);
      });
      sel.selectedIndex = state.months.length - 1;
      sel.disabled = false;
    }
    
    if (data.data_quality) {
      renderDataQuality(data.data_quality);
    }
    
    setStepState(2, "done");
    updateActionButtons();
    saveSession();
  } catch (err) {
    $("dz-excel-text").textContent = "Cliquez ou glissez-déposez le fichier Excel ici";
    $("excel-status").innerHTML = `<span class="file-err">✕ ${err.message}</span>`;
  }
}

$("toggle-reception").addEventListener("change", (e) => {
  state.fillReception = e.target.checked;
  $("reception-upload-wrap").classList.toggle("hidden", !e.target.checked);
  updateActionButtons();
  saveSession();
});

async function handleReceptionFile(file) {
  $("dz-reception-text").textContent = `Envoi de "${file.name}"…`;
  $("reception-status").innerHTML = "";
  const fd = new FormData();
  fd.append("reception", file);
  try {
    const data = await postForm("/api/upload", fd);
    $("dz-reception-text").textContent = "Cliquez ou glissez-déposez le fichier Excel de réception ici";
    if (data.reception_error) {
      $("reception-status").innerHTML = `<span class="file-err">✕ ${data.reception_error}</span>`;
      return;
    }
    state.receptionId = data.reception_id;
    state.receptionName = data.reception_name;
    $("reception-status").innerHTML =
      `<span class="file-ok">✓ <span class="file-name">${data.reception_name}</span> chargé — ` +
      `${data.reception_terrain_count} étude(s) terrain, ${data.reception_stack_count} réception(s) stack</span>`;
    setStepState(3, "done");
    updateActionButtons();
    saveSession();
  } catch (err) {
    $("dz-reception-text").textContent = "Cliquez ou glissez-déposez le fichier Excel de réception ici";
    $("reception-status").innerHTML = `<span class="file-err">✕ ${err.message}</span>`;
  }
}

function selectedCategories() {
  return Array.from(document.querySelectorAll(".cat-checkbox:checked")).map(cb => cb.value);
}

function renderPreviewTable(rows) {
  if (!rows.length) return `<p class="preview-empty">Aucune ligne pour cette section / période.</p>`;
  const cols = [
    ["etudes", "Étude(s)"], ["type", "Type"], ["date", "Date"], ["source", "Source"],
    ["capacite", "Capacité"], ["realisation", "Réalisation"], ["remarque", "Remarque"],
  ];
  let html = `<table class="preview-table"><thead><tr>${cols.map(c => `<th>${c[1]}</th>`).join("")}</tr></thead><tbody>`;
  rows.forEach(r => {
    html += `<tr>${cols.map(c => `<td>${(r[c[0]] || "").toString().replace(/\n/g, "<br>")}</td>`).join("")}</tr>`;
  });
  html += "</tbody></table>";
  return html;
}

$("btn-validate").addEventListener("click", async () => {
  if (!state.excelId) return;
  const sel = $("select-month");
  const [year, month] = (sel.value || "").split("-");
  if (!year || !month) { showGlobalError("Sélectionnez d'abord un mois."); return; }
  showGlobalError(null);
  
  $("btn-validate").disabled = true;
  $("btn-validate").innerHTML = `<span class="spin"></span> Validation…`;
  try {
    const quality = await postJSON("/api/validate_data", { 
      excel_id: state.excelId, 
      year: +year, 
      month: +month 
    });
    renderDataQuality(quality);
    
    if (quality.ready_to_generate && !quality.has_warnings) {
      showToast("✅ Données validées — prêtes à générer !");
    } else if (quality.ready_to_generate) {
      showToast("⚠️ Données utilisables mais avec avertissements");
      renderIssuesModal();
    } else {
      showToast("❌ Données invalides — voir les détails");
      renderIssuesModal();
    }
  } catch (err) {
    showGlobalError(err.message);
  } finally {
    $("btn-validate").disabled = false;
    $("btn-validate").textContent = "🔍 Valider les données";
  }
});

$("btn-view-issues").addEventListener("click", renderIssuesModal);
$("btn-close-issues").addEventListener("click", () => $("issues-modal").classList.add("hidden"));
$("issues-modal").addEventListener("click", (e) => { if (e.target.id === "issues-modal") $("issues-modal").classList.add("hidden"); });

$("btn-preview").addEventListener("click", async () => {
  if (!state.excelId) return;
  const sel = $("select-month");
  const [year, month] = (sel.value || "").split("-");
  if (!year || !month) { showGlobalError("Sélectionnez d'abord un mois."); return; }
  showGlobalError(null);

  $("btn-preview").disabled = true;
  $("btn-preview").innerHTML = `<span class="spin"></span> Chargement…`;
  try {
    const preview = await postJSON("/api/preview", { excel_id: state.excelId, year: +year, month: +month });
    const selected = new Set(selectedCategories());

    let body = "";
    Object.entries(preview).forEach(([code, section]) => {
      if (!selected.has(code)) return;
      body += `<div class="preview-section"><h4>${section.label} (${section.rows.length})</h4>${renderPreviewTable(section.rows)}</div>`;
    });

    if (state.fillReception && state.receptionId) {
      const rec = await postJSON("/api/preview_reception", { reception_id: state.receptionId });
      body += `<div class="preview-section"><h4>VI-3-1 Terrain — étude(s) trouvée(s) (${rec.terrain.length})</h4>`;
      body += rec.terrain.length
        ? `<table class="preview-table"><thead><tr><th>Étude</th><th>+ Profils/Swaths</th><th>+ Nombre Sup 3592</th></tr></thead><tbody>` +
          rec.terrain.map(t => `<tr><td>${t.etude}</td><td>${t.profile_append}</td><td>${t.cassette_append}</td></tr>`).join("") +
          `</tbody></table>`
        : `<p class="preview-empty">Aucune donnée terrain.</p>`;
      body += `</div>`;

      body += `<div class="preview-section"><h4>VI-3-2/3 Stack — réception(s) trouvée(s) (${rec.stack.length})</h4>`;
      body += rec.stack.length
        ? `<table class="preview-table"><thead><tr><th>Direction</th><th>Étude</th><th>Date</th><th>Centre</th><th>Données reçues</th></tr></thead><tbody>` +
          rec.stack.map(s => `<tr><td>${s.direction}</td><td>${s.etude}</td><td>${s.date}</td><td>${s.centre}</td><td>${s.donnees}</td></tr>`).join("") +
          `</tbody></table>`
        : `<p class="preview-empty">Aucune réception stack.</p>`;
      body += `</div>`;
    }

    $("preview-body").innerHTML = body || `<p class="preview-empty">Rien à afficher pour la sélection actuelle.</p>`;
    $("preview-modal").classList.remove("hidden");
  } catch (err) {
    showGlobalError(err.message);
  } finally {
    $("btn-preview").disabled = false;
    $("btn-preview").textContent = "👁 Aperçu";
  }
});

$("btn-close-preview").addEventListener("click", () => $("preview-modal").classList.add("hidden"));
$("preview-modal").addEventListener("click", (e) => { if (e.target.id === "preview-modal") $("preview-modal").classList.add("hidden"); });

function renderResult(summary, filename, outputId) {
  const list = $("result-list");
  list.innerHTML = "";
  Object.entries(summary).forEach(([label, entry]) => {
    const li = document.createElement("li");
    li.innerHTML = `<span class="dot ${statusDotClass(entry.status)}"></span>` +
      `<span class="r-label">${label}</span><span>${entry.msg}</span>`;
    list.appendChild(li);
  });

  const dl = $("btn-download");
  dl.href = `/api/download/${outputId}?filename=${encodeURIComponent(filename)}`;
  dl.setAttribute("download", filename);

  $("card-result").classList.remove("hidden");
  $("card-result").scrollIntoView({ behavior: "smooth", block: "start" });
}

$("btn-generate").addEventListener("click", async () => {
  const sel = $("select-month");
  const [year, month] = (sel.value || "").split("-");
  if (!year || !month) { showGlobalError("Sélectionnez d'abord un mois."); return; }
  showGlobalError(null);

  const btn = $("btn-generate");
  btn.disabled = true;
  const originalLabel = btn.textContent;
  btn.innerHTML = `<span class="spin"></span> Génération…`;

  try {
    const payload = {
      excel_id: state.excelId,
      word_id: state.wordId,
      year: +year,
      month: +month,
      categories: selectedCategories(),
      fill_reception: !!(state.fillReception && state.receptionId),
      reception_id: state.receptionId,
    };
    const data = await postJSON("/api/generate", payload);
    renderResult(data.summary, data.filename, data.output_id);
    setStepState(4, "done");
    showToast("✅ Rapport généré avec succès.");
  } catch (err) {
    showGlobalError(err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = originalLabel;
    updateActionButtons();
  }
});

$("btn-new-report").addEventListener("click", () => {
  if (confirm("Démarrer un nouveau rapport ? Les fichiers uploadés resteront disponibles.")) {
    $("card-result").classList.add("hidden");
    setStepState(4, "active");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
});

function initDarkMode() {
  const saved = localStorage.getItem("darkMode");
  if (saved === "true") {
    document.body.classList.add("dark-mode");
    $("btn-dark-mode").textContent = "☀️";
  }
}

$("btn-dark-mode").addEventListener("click", () => {
  document.body.classList.toggle("dark-mode");
  const isDark = document.body.classList.contains("dark-mode");
  $("btn-dark-mode").textContent = isDark ? "☀️" : "🌙";
  localStorage.setItem("darkMode", isDark);
});

const shortcuts = {
  '1': () => $("input-word").click(),
  '2': () => $("input-excel").click(),
  '3': () => $("input-reception").click(),
  'p': () => !$("btn-preview").disabled && $("btn-preview").click(),
  'g': () => !$("btn-generate").disabled && $("btn-generate").click(),
};

document.addEventListener('keydown', (e) => {
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
  
  if (e.key === 'Escape') {
    $("preview-modal").classList.add("hidden");
    $("issues-modal").classList.add("hidden");
    return;
  }
  
  if (e.ctrlKey || e.metaKey) {
    const key = e.key.toLowerCase();
    if (shortcuts[key]) {
      e.preventDefault();
      shortcuts[key]();
    }
  }
});

wireDropzone("dz-word", "input-word", handleWordFile);
wireDropzone("dz-excel", "input-excel", handleExcelFile);
wireDropzone("dz-reception", "input-reception", handleReceptionFile);

$("cat-all").addEventListener("click", (e) => {
  e.preventDefault();
  document.querySelectorAll(".cat-checkbox").forEach(cb => cb.checked = true);
  saveSession();
});
$("cat-none").addEventListener("click", (e) => {
  e.preventDefault();
  document.querySelectorAll(".cat-checkbox").forEach(cb => cb.checked = false);
  saveSession();
});

$("select-month").addEventListener("change", saveSession);
document.querySelectorAll(".cat-checkbox").forEach(cb => {
  cb.addEventListener("change", saveSession);
});

initDarkMode();
restoreSession();
updateActionButtons();