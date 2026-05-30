const { invoke } = window.__TAURI__.core;
const { load } = window.__TAURI__.store;
const { listen } = window.__TAURI__.event;

let store;
let automations = [];

window.addEventListener("DOMContentLoaded", async () => {
  store = await load("config.json", { autoSave: true });

  const emptyState = document.querySelector("#empty-state");
  const wizardContainer = document.querySelector("#wizard-container");
  const automationList = document.querySelector("#automation-list");
  const automationCards = document.querySelector("#automation-cards");
  const contentSubtitle = document.querySelector("#content-subtitle");

  let selectedMailboxPath = "";
  let selectedDestinationMailboxPath = "";

  function renderAutomationList() {
    automationCards.innerHTML = "";
    if (automations.length === 0) {
      applyViewState();
      return;
    }
    automations.forEach((config, index) => {
      const freqLabel = config.analysisFrequency === "weekly" ? "Semanal" : "Mensual";
      const criterionLabel = config.criterion === "copy" ? "Copiar" : "Mover";
      const card = document.createElement("div");
      card.className = "automation-card";
      card.innerHTML = `
        <div class="card-row">
          <span class="card-label">Origen</span>
          <span class="card-value">${config.source}</span>
        </div>
        <div class="card-row">
          <span class="card-label">Destino</span>
          <span class="card-value">${config.destination}</span>
        </div>
        <div class="card-row">
          <span class="card-label">Frecuencia</span>
          <span class="card-value">${freqLabel}</span>
        </div>
        <div class="card-row">
          <span class="card-label">Criterio</span>
          <span class="card-value">${criterionLabel}</span>
        </div>
        <div class="card-actions">
          <button class="delete-btn" data-index="${index}">Eliminar</button>
        </div>
      `;
      card.querySelector(".delete-btn").addEventListener("click", async () => {
        automations.splice(index, 1);
        await store.set("automations", automations);
        renderAutomationList();
      });
      automationCards.appendChild(card);
    });
  }

  function applyViewState() {
    if (automations.length > 0) {
      emptyState.classList.add("hidden");
      wizardContainer.classList.add("hidden");
      automationList.classList.remove("hidden");
      contentSubtitle.textContent = "Automatizaciones configuradas y activas.";
      renderAutomationList();
    } else {
      emptyState.classList.remove("hidden");
      wizardContainer.classList.add("hidden");
      automationList.classList.add("hidden");
      contentSubtitle.textContent = "Gestione la automatización de sus respaldos de Outlook.";
    }
  }

  async function loadWizard() {
    const response = await fetch("wizard.html");
    wizardContainer.innerHTML = await response.text();

    let currentStep = 1;
    const wizardSteps = [
      document.querySelector("#wizard-step-1"),
      document.querySelector("#wizard-step-2"),
      document.querySelector("#wizard-step-3")
    ];
    const stepIndicators = document.querySelectorAll(".step-indicator");
    const stepLines = document.querySelectorAll(".step-line");

    const wizardAnalysisFrequencySelect = document.querySelector("#wizard-analysis-frequency");

    const wizardCriterionSelect = document.querySelector("#wizard-criterion");
    const wizardTriggerConditionContainer = document.querySelector("#wizard-trigger-condition-container");
    const wizardTriggerConditionSelect = document.querySelector("#wizard-trigger-condition");

    const summarySource = document.querySelector("#summary-source");
    const summaryAnalysisFrequency = document.querySelector("#summary-analysis-frequency");
    const summaryDestination = document.querySelector("#summary-destination");
    const summaryCriterion = document.querySelector("#summary-criterion");

    const btnWizardBack = document.querySelector("#btn-wizard-back");
    const btnWizardNext = document.querySelector("#btn-wizard-next");
    const btnWizardCancel = document.querySelector("#btn-wizard-cancel");

    const btnSearchMailbox = document.querySelector("#btn-search-mailbox");
    const mailboxLoading = document.querySelector("#mailbox-loading");
    const mailboxListContainer = document.querySelector("#mailbox-list-container");
    const mailboxList = document.querySelector("#mailbox-list");

    const btnSearchDestinationMailbox = document.querySelector("#btn-search-destination-mailbox");
    const mailboxDestinationLoading = document.querySelector("#mailbox-destination-loading");
    const mailboxDestinationListContainer = document.querySelector("#mailbox-destination-list-container");
    const mailboxDestinationList = document.querySelector("#mailbox-destination-list");

    function resetWizard() {
      currentStep = 1;
      selectedMailboxPath = "";
      selectedDestinationMailboxPath = "";

      wizardAnalysisFrequencySelect.value = "weekly";
      wizardCriterionSelect.value = "copy";
      wizardTriggerConditionContainer.classList.add("hidden");
      wizardTriggerConditionSelect.value = "80";

      mailboxListContainer.classList.add("hidden");
      mailboxLoading.classList.add("hidden");
      mailboxList.innerHTML = "";

      mailboxDestinationListContainer.classList.add("hidden");
      mailboxDestinationLoading.classList.add("hidden");
      mailboxDestinationList.innerHTML = "";

      showStep(1);
    }

    wizardCriterionSelect.addEventListener("change", (e) => {
      if (e.target.value === "move") {
        wizardTriggerConditionContainer.classList.remove("hidden");
      } else {
        wizardTriggerConditionContainer.classList.add("hidden");
      }
    });

    btnSearchMailbox.addEventListener("click", async () => {
      btnSearchMailbox.disabled = true;
      mailboxLoading.classList.remove("hidden");
      mailboxListContainer.classList.add("hidden");
      mailboxList.innerHTML = "";
      selectedMailboxPath = "";

      try {
        const stores = await invoke("find_outlook_stores");

        mailboxLoading.classList.add("hidden");
        btnSearchMailbox.disabled = false;

        if (!stores || stores.length === 0) {
          mailboxList.innerHTML = `<div style="padding: 10px; font-size:12px; color: #666; text-align: center;">No se encontraron buzones OST activos en Outlook. Asegúrese de que Outlook está configurado correctamente.</div>`;
          mailboxListContainer.classList.remove("hidden");
          return;
        }

        stores.forEach(store => {
          const card = document.createElement("div");
          card.className = "mailbox-card";
          card.innerHTML = `
            <div class="mailbox-title">${store.displayName}</div>
            <div class="mailbox-path" title="${store.filePath}">${store.filePath}</div>
            <div class="mailbox-size">Tamaño: ${store.fileSize}</div>
          `;

          card.addEventListener("click", () => {
            document.querySelectorAll(".mailbox-card").forEach(c => c.classList.remove("selected"));
            card.classList.add("selected");
            selectedMailboxPath = store.filePath;
          });

          mailboxList.appendChild(card);
        });

        mailboxListContainer.classList.remove("hidden");

      } catch (error) {
        console.error("Error al buscar buzones:", error);
        mailboxLoading.classList.add("hidden");
        btnSearchMailbox.disabled = false;
        alert("Error al comunicarse con Outlook: " + error);
      }
    });

    btnSearchDestinationMailbox.addEventListener("click", async () => {
      btnSearchDestinationMailbox.disabled = true;
      mailboxDestinationLoading.classList.remove("hidden");
      mailboxDestinationListContainer.classList.add("hidden");
      mailboxDestinationList.innerHTML = "";
      selectedDestinationMailboxPath = "";

      try {
        const stores = await invoke("find_outlook_stores");

        mailboxDestinationLoading.classList.add("hidden");
        btnSearchDestinationMailbox.disabled = false;

        const filteredStores = stores.filter(store => store.filePath !== selectedMailboxPath);

        if (!filteredStores || filteredStores.length === 0) {
          mailboxDestinationList.innerHTML = `<div style="padding: 10px; font-size:12px; color: #666; text-align: center;">No se encontraron otros buzones OST disponibles en Outlook que difieran del de origen.</div>`;
          mailboxDestinationListContainer.classList.remove("hidden");
          return;
        }

        filteredStores.forEach(store => {
          const card = document.createElement("div");
          card.className = "mailbox-card";
          card.innerHTML = `
            <div class="mailbox-title">${store.displayName}</div>
            <div class="mailbox-path" title="${store.filePath}">${store.filePath}</div>
            <div class="mailbox-size">Tamaño: ${store.fileSize}</div>
          `;

          card.addEventListener("click", () => {
            mailboxDestinationList.querySelectorAll(".mailbox-card").forEach(c => c.classList.remove("selected"));
            card.classList.add("selected");
            selectedDestinationMailboxPath = store.filePath;
          });

          mailboxDestinationList.appendChild(card);
        });

        mailboxDestinationListContainer.classList.remove("hidden");

      } catch (error) {
        console.error("Error al buscar buzones de destino:", error);
        mailboxDestinationLoading.classList.add("hidden");
        btnSearchDestinationMailbox.disabled = false;
        alert("Error al comunicarse con Outlook para buscar destino: " + error);
      }
    });

    function showStep(step) {
      currentStep = step;

      wizardSteps.forEach((stepEl, idx) => {
        if (idx + 1 === step) {
          stepEl.classList.remove("hidden");
        } else {
          stepEl.classList.add("hidden");
        }
      });

      stepIndicators.forEach((indicator, idx) => {
        const stepNum = idx + 1;
        indicator.className = "step-indicator";
        if (stepNum < step) {
          indicator.classList.add("completed");
        } else if (stepNum === step) {
          indicator.classList.add("active");
        }
      });

      stepLines.forEach((line, idx) => {
        const lineNum = idx + 1;
        line.className = "step-line";
        if (lineNum < step) {
          line.classList.add("completed");
        }
      });

      if (step === 1) {
        btnWizardBack.classList.add("hidden");
        btnWizardNext.textContent = "Siguiente";
      } else if (step === 3) {
        btnWizardBack.classList.remove("hidden");
        btnWizardNext.textContent = "Finalizar y Activar";
      } else {
        btnWizardBack.classList.remove("hidden");
        btnWizardNext.textContent = "Siguiente";
      }
    }

    btnWizardNext.addEventListener("click", async () => {
      if (currentStep === 1) {
        if (!selectedMailboxPath) {
          alert("Por favor busque y seleccione el buzón de Outlook que desea respaldar de la lista.");
          return;
        }
        showStep(2);
      } else if (currentStep === 2) {
        if (!selectedDestinationMailboxPath) {
          alert("Por favor busque y seleccione un buzón de destino para almacenar el respaldo.");
          return;
        }

        summarySource.textContent = selectedMailboxPath;

        const analysisFreqVal = wizardAnalysisFrequencySelect.value;
        summaryAnalysisFrequency.textContent = analysisFreqVal === "weekly" ? "1 vez a la semana" : "1 vez al mes";

        summaryDestination.textContent = selectedDestinationMailboxPath;

        const criterionVal = wizardCriterionSelect.value;
        if (criterionVal === "copy") {
          summaryCriterion.textContent = "Copiar archivo (Mantener original)";
        } else {
          const triggerPercent = wizardTriggerConditionSelect.value;
          summaryCriterion.textContent = `Mover archivo (Liberar espacio) - Activar cuando buzón sea mayor al ${triggerPercent}%`;
        }

        showStep(3);
      } else if (currentStep === 3) {
        automations.push({
          source: selectedMailboxPath,
          analysisFrequency: wizardAnalysisFrequencySelect.value,
          destination: selectedDestinationMailboxPath,
          criterion: wizardCriterionSelect.value,
          triggerCondition: wizardCriterionSelect.value === "move" ? wizardTriggerConditionSelect.value : null
        });

        await store.set("automations", automations);
        applyViewState();
      }
    });

    btnWizardBack.addEventListener("click", () => {
      if (currentStep > 1) {
        showStep(currentStep - 1);
      }
    });

    btnWizardCancel.addEventListener("click", () => {
      if (confirm("¿Está seguro de que desea cancelar el asistente? Se perderán los datos introducidos.")) {
        applyViewState();
      }
    });

    resetWizard();
  }

  document.querySelector("#btn-create-automation").addEventListener("click", async () => {
    emptyState.classList.add("hidden");
    wizardContainer.classList.remove("hidden");
    contentSubtitle.textContent = "Siga los pasos para configurar una nueva automatización de respaldo.";
    await loadWizard();
  });

  document.querySelector("#btn-create-from-list")?.addEventListener("click", async () => {
    automationList.classList.add("hidden");
    wizardContainer.classList.remove("hidden");
    contentSubtitle.textContent = "Siga los pasos para configurar una nueva automatización de respaldo.";
    await loadWizard();
  });

  // ------------------------------------------------------------------
  // Navegación entre secciones (sidebar)
  // ------------------------------------------------------------------
  const contentTitle = document.querySelector("#content-title");
  const automationSection = document.querySelector("#automation-section");
  const scanSection = document.querySelector("#scan-section");
  const btnAutomatizar = document.querySelector("#btn-automatizar");
  const btnEscanear = document.querySelector("#btn-escanear");

  function activateSection(section) {
    [btnAutomatizar, btnEscanear].forEach(b => b.classList.remove("active"));

    if (section === "scan") {
      btnEscanear.classList.add("active");
      automationSection.classList.add("hidden");
      scanSection.classList.remove("hidden");
      contentTitle.textContent = "Escanear Buzón";
      contentSubtitle.textContent = "Analice el contenido de un buzón de Outlook por carpetas y años.";
    } else {
      btnAutomatizar.classList.add("active");
      scanSection.classList.add("hidden");
      automationSection.classList.remove("hidden");
      contentTitle.textContent = "Automatizar Respaldo";
      applyViewState();
    }
  }

  btnAutomatizar.addEventListener("click", () => activateSection("automation"));
  btnEscanear.addEventListener("click", () => activateSection("scan"));

  let scanLoaded = false;

  function activateSection(section) {
    [btnAutomatizar, btnEscanear].forEach(b => b.classList.remove("active"));

    if (section === "scan") {
      btnEscanear.classList.add("active");
      automationSection.classList.add("hidden");
      scanSection.classList.remove("hidden");
      contentTitle.textContent = "Escanear Buzón";
      contentSubtitle.textContent = "Analice el contenido de un buzón de Outlook por carpetas y años.";
      if (!scanLoaded) loadScanSection();
    } else {
      btnAutomatizar.classList.add("active");
      scanSection.classList.add("hidden");
      automationSection.classList.remove("hidden");
      contentTitle.textContent = "Automatizar Respaldo";
      applyViewState();
    }
  }

  async function loadScanSection() {
    scanLoaded = true;
    const response = await fetch("scan.html");
    scanSection.innerHTML = await response.text();

    const btnScanSearchMailbox = document.querySelector("#btn-scan-search-mailbox");
    const scanMailboxLoading = document.querySelector("#scan-mailbox-loading");
    const scanMailboxListContainer = document.querySelector("#scan-mailbox-list-container");
    const scanMailboxList = document.querySelector("#scan-mailbox-list");
    const btnAnalyze = document.querySelector("#btn-analyze");

    const scanProgressContainer = document.querySelector("#scan-progress-container");
    const scanProgressStatus = document.querySelector("#scan-progress-status");
    const scanProgressPercent = document.querySelector("#scan-progress-percent");
    const scanProgressFill = document.querySelector("#scan-progress-fill");
    const scanProgressFolders = document.querySelector("#scan-progress-folders");
    const scanProgressItems = document.querySelector("#scan-progress-items");
    const scanProgressCurrent = document.querySelector("#scan-progress-current");
    const scanResultContainer = document.querySelector("#scan-result-container");

    let selectedScanStoreId = "";
    let scanRunning = false;

    btnScanSearchMailbox.addEventListener("click", async () => {
      btnScanSearchMailbox.disabled = true;
      scanMailboxLoading.classList.remove("hidden");
      scanMailboxListContainer.classList.add("hidden");
      scanMailboxList.innerHTML = "";
      selectedScanStoreId = "";
      btnAnalyze.disabled = true;

      try {
        const stores = await invoke("find_outlook_stores");

        scanMailboxLoading.classList.add("hidden");
        btnScanSearchMailbox.disabled = false;

        if (!stores || stores.length === 0) {
          scanMailboxList.innerHTML = `<div style="padding: 10px; font-size:12px; color: #666; text-align: center;">No se encontraron buzones OST activos en Outlook.</div>`;
          scanMailboxListContainer.classList.remove("hidden");
          return;
        }

        stores.forEach(store => {
          const card = document.createElement("div");
          card.className = "mailbox-card";
          card.innerHTML = `
            <div class="mailbox-title">${store.displayName}</div>
            <div class="mailbox-path" title="${store.filePath}">${store.filePath}</div>
            <div class="mailbox-size">Tamaño: ${store.fileSize}</div>
          `;
          card.addEventListener("click", () => {
            if (scanRunning) return;
            scanMailboxList.querySelectorAll(".mailbox-card").forEach(c => c.classList.remove("selected"));
            card.classList.add("selected");
            selectedScanStoreId = store.storeId || "";
            btnAnalyze.disabled = !selectedScanStoreId;
          });
          scanMailboxList.appendChild(card);
        });

        scanMailboxListContainer.classList.remove("hidden");
      } catch (error) {
        console.error("Error al buscar buzones:", error);
        scanMailboxLoading.classList.add("hidden");
        btnScanSearchMailbox.disabled = false;
        alert("Error al comunicarse con Outlook: " + error);
      }
    });

    function resetScanProgressUI() {
      scanProgressContainer.classList.remove("hidden");
      scanResultContainer.classList.add("hidden");
      scanResultContainer.innerHTML = "";
      scanProgressStatus.textContent = "Iniciando análisis...";
      scanProgressPercent.textContent = "0%";
      scanProgressFill.style.width = "0%";
      scanProgressFolders.textContent = "Carpetas: 0 / 0";
      scanProgressItems.textContent = "Elementos: 0";
      scanProgressCurrent.textContent = "";
    }

    function renderScanSummary(summary) {
      const totals = summary.totals || {};
      const source = summary.source || {};
      const scan = summary.scan || {};
      const yearRows = (summary.yearBreakdown || [])
        .map(y => `<tr><td>${y.year}</td><td>${y.count.toLocaleString()}</td></tr>`)
        .join("");
      const topFolders = (summary.topFoldersBySize && summary.topFoldersBySize.length
        ? summary.topFoldersBySize
        : summary.topFoldersByItems || [])
        .slice(0, 10)
        .map(f => `<tr><td title="${f.path}">${f.path}</td><td>${(f.itemCount || 0).toLocaleString()}</td><td>${f.sizeHuman || "-"}</td></tr>`)
        .join("");

      scanResultContainer.innerHTML = `
        <h3 style="margin-bottom:10px;">Resultado del análisis</h3>
        <div class="scan-summary-grid">
          <div class="scan-summary-item"><span class="card-label">Buzón</span><span class="card-value">${source.storeDisplayName || "-"}</span></div>
          <div class="scan-summary-item"><span class="card-label">Tamaño PST</span><span class="card-value">${source.pstSizeHuman || "-"}</span></div>
          <div class="scan-summary-item"><span class="card-label">Carpetas</span><span class="card-value">${scan.matchedFolders || 0}</span></div>
          <div class="scan-summary-item"><span class="card-label">Elementos</span><span class="card-value">${(totals.items || 0).toLocaleString()}</span></div>
          <div class="scan-summary-item"><span class="card-label">Con fecha</span><span class="card-value">${(totals.datedItems || 0).toLocaleString()}</span></div>
          <div class="scan-summary-item"><span class="card-label">Tamaño total</span><span class="card-value">${totals.sizeHuman || "-"}</span></div>
        </div>
        ${yearRows ? `<h4 style="margin-top:16px;">Por año</h4>
        <table class="scan-table"><thead><tr><th>Año</th><th>Elementos</th></tr></thead><tbody>${yearRows}</tbody></table>` : ""}
        ${topFolders ? `<h4 style="margin-top:16px;">Carpetas principales</h4>
        <table class="scan-table"><thead><tr><th>Carpeta</th><th>Elementos</th><th>Tamaño</th></tr></thead><tbody>${topFolders}</tbody></table>` : ""}
      `;
      scanResultContainer.classList.remove("hidden");
    }

    await listen("scan://event", (event) => {
      const payload = event.payload || {};
      if (payload.type === "scanMeta") {
        scanProgressStatus.textContent = "Analizando carpetas...";
        scanProgressFolders.textContent = `Carpetas: 0 / ${payload.totalFolders || 0}`;
      } else if (payload.type === "scanProgress") {
        const percent = payload.percent || 0;
        scanProgressPercent.textContent = `${percent}%`;
        scanProgressFill.style.width = `${percent}%`;
        scanProgressFolders.textContent = `Carpetas: ${payload.scannedFolders || 0} / ${payload.totalFolders || 0}`;
        scanProgressItems.textContent = `Elementos: ${(payload.accumulatedItems || 0).toLocaleString()}`;
        if (payload.folderPath) {
          scanProgressCurrent.textContent = payload.folderPath;
        }
        if (payload.phase === "completed") {
          scanProgressStatus.textContent = "Generando resumen...";
        }
      } else if (payload.type === "log") {
        scanProgressStatus.textContent = payload.message || scanProgressStatus.textContent;
      }
    });

    await listen("scan://complete", (event) => {
      scanRunning = false;
      btnAnalyze.disabled = false;
      btnScanSearchMailbox.disabled = false;
      scanProgressStatus.textContent = "Análisis completado";
      scanProgressPercent.textContent = "100%";
      scanProgressFill.style.width = "100%";
      renderScanSummary(event.payload || {});
    });

    await listen("scan://error", (event) => {
      scanRunning = false;
      btnAnalyze.disabled = false;
      btnScanSearchMailbox.disabled = false;
      scanProgressStatus.textContent = "Error durante el análisis";
      alert("Error en el escaneo: " + (event.payload || "desconocido"));
    });

    btnAnalyze.addEventListener("click", async () => {
      if (!selectedScanStoreId) {
        alert("Seleccione un buzón para analizar.");
        return;
      }
      if (scanRunning) return;

      scanRunning = true;
      btnAnalyze.disabled = true;
      btnScanSearchMailbox.disabled = true;
      resetScanProgressUI();

      try {
        await invoke("scan_mailbox", { storeId: selectedScanStoreId });
      } catch (error) {
        console.error("Error al escanear buzón:", error);
        scanRunning = false;
        btnAnalyze.disabled = false;
        btnScanSearchMailbox.disabled = false;
        scanProgressStatus.textContent = "Error durante el análisis";
      }
    });
  }

  automations = await store.get("automations") || [];
  applyViewState();
});
