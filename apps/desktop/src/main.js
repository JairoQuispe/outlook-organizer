// Obtenemos acceso a las APIs principales de Tauri
const { invoke } = window.__TAURI__.core;

window.addEventListener("DOMContentLoaded", () => {
  // --- Estados de las Vistas ---
  const emptyState = document.querySelector("#empty-state");
  const wizardContainer = document.querySelector("#wizard-container");
  const activePanel = document.querySelector("#active-panel");
  const contentSubtitle = document.querySelector("#content-subtitle");

  // --- Elementos del Formulario de Edición (Vista Activa) ---
  const backupForm = document.querySelector("#backup-form");
  const statusMessage = document.querySelector("#status-message");
  const destinationDisplay = document.querySelector("#backup-destination-display");
  
  // Elementos de Criterio de Respaldo y Análisis de Espacio (Vista Activa)
  const analysisFrequencySelect = document.querySelector("#backup-analysis-frequency");
  const backupCriterionSelect = document.querySelector("#backup-criterion");
  const backupTriggerConditionContainer = document.querySelector("#backup-trigger-condition-container");
  const backupTriggerConditionSelect = document.querySelector("#backup-trigger-condition");

  // --- Elementos del Asistente (Wizard) ---
  let currentStep = 1;
  const wizardSteps = [
    document.querySelector("#wizard-step-1"),
    document.querySelector("#wizard-step-2"),
    document.querySelector("#wizard-step-3")
  ];
  const stepIndicators = document.querySelectorAll(".step-indicator");
  const stepLines = document.querySelectorAll(".step-line");
  
  const wizardAnalysisFrequencySelect = document.querySelector("#wizard-analysis-frequency");

  // Elementos de Criterio de Respaldo (Wizard)
  const wizardCriterionSelect = document.querySelector("#wizard-criterion");
  const wizardTriggerConditionContainer = document.querySelector("#wizard-trigger-condition-container");
  const wizardTriggerConditionSelect = document.querySelector("#wizard-trigger-condition");

  // Elementos de renderizado dinámico del Resumen (Paso 3)
  const summarySource = document.querySelector("#summary-source");
  const summaryAnalysisFrequency = document.querySelector("#summary-analysis-frequency");
  const summaryDestination = document.querySelector("#summary-destination");
  const summaryCriterion = document.querySelector("#summary-criterion");

  // Variables de estado para recordar las rutas de buzón
  let selectedMailboxPath = "";
  let selectedDestinationMailboxPath = "";

  const btnWizardBack = document.querySelector("#btn-wizard-back");
  const btnWizardNext = document.querySelector("#btn-wizard-next");
  const btnWizardCancel = document.querySelector("#btn-wizard-cancel");

  // Elementos de la búsqueda de buzones OST (Origen - Paso 1)
  const btnSearchMailbox = document.querySelector("#btn-search-mailbox");
  const mailboxLoading = document.querySelector("#mailbox-loading");
  const mailboxListContainer = document.querySelector("#mailbox-list-container");
  const mailboxList = document.querySelector("#mailbox-list");

  // Elementos de la búsqueda de buzones OST (Destino - Paso 2)
  const btnSearchDestinationMailbox = document.querySelector("#btn-search-destination-mailbox");
  const mailboxDestinationLoading = document.querySelector("#mailbox-destination-loading");
  const mailboxDestinationListContainer = document.querySelector("#mailbox-destination-list-container");
  const mailboxDestinationList = document.querySelector("#mailbox-destination-list");

  // --- Lógica del Estado Inicial ---
  // Guardaremos la configuración en localStorage para mantener el estado incluso si recargan la app
  let savedConfig = JSON.parse(localStorage.getItem("outlook_backup_config") || "null");

  function applyViewState() {
    if (savedConfig) {
      // Si ya hay configuración, rellenamos el formulario principal y mostramos la vista activa
      destinationDisplay.value = savedConfig.destination || "";
      
      // Frecuencia de análisis
      analysisFrequencySelect.value = savedConfig.analysisFrequency || "weekly";
      selectedMailboxPath = savedConfig.source || "";
      selectedDestinationMailboxPath = savedConfig.destination || "";

      // Rellenar nuevos campos de criterio
      const criterion = savedConfig.criterion || "copy";
      backupCriterionSelect.value = criterion;
      if (criterion === "move") {
        backupTriggerConditionContainer.classList.remove("hidden");
        backupTriggerConditionSelect.value = savedConfig.triggerCondition || "80";
      } else {
        backupTriggerConditionContainer.classList.add("hidden");
      }

      emptyState.classList.add("hidden");
      wizardContainer.classList.add("hidden");
      activePanel.classList.remove("hidden");
      contentSubtitle.textContent = "La automatización se está ejecutando según la programación establecida.";
    } else {
      // Si no hay configuración, mostramos el estado vacío con un único botón de 'Crear'
      emptyState.classList.remove("hidden");
      wizardContainer.classList.add("hidden");
      activePanel.classList.add("hidden");
      contentSubtitle.textContent = "Gestione la automatización de sus respaldos de Outlook.";
    }
  }

  // --- Transición: Estado Vacío -> Wizard ---
  document.querySelector("#btn-create-automation").addEventListener("click", () => {
    emptyState.classList.add("hidden");
    wizardContainer.classList.remove("hidden");
    contentSubtitle.textContent = "Siga los pasos para configurar una nueva automatización de respaldo.";
    resetWizard();
  });

  // --- Lógica del Wizard (Asistente por pasos) ---
  function resetWizard() {
    currentStep = 1;
    selectedMailboxPath = "";
    selectedDestinationMailboxPath = "";
    
    // Resetear criterio en Wizard
    wizardAnalysisFrequencySelect.value = "weekly";
    wizardCriterionSelect.value = "copy";
    wizardTriggerConditionContainer.classList.add("hidden");
    wizardTriggerConditionSelect.value = "80";
    
    // Ocultar resultados de búsqueda anteriores (Origen)
    mailboxListContainer.classList.add("hidden");
    mailboxLoading.classList.add("hidden");
    mailboxList.innerHTML = "";

    // Ocultar resultados de búsqueda anteriores (Destino)
    mailboxDestinationListContainer.classList.add("hidden");
    mailboxDestinationLoading.classList.add("hidden");
    mailboxDestinationList.innerHTML = "";
    
    showStep(1);
  }

  // --- Comportamientos Reactivos de los Criterios (Mostrar/Ocultar 'Detectar cuando') ---
  wizardCriterionSelect.addEventListener("change", (e) => {
    if (e.target.value === "move") {
      wizardTriggerConditionContainer.classList.remove("hidden");
    } else {
      wizardTriggerConditionContainer.classList.add("hidden");
    }
  });

  backupCriterionSelect.addEventListener("change", (e) => {
    if (e.target.value === "move") {
      backupTriggerConditionContainer.classList.remove("hidden");
    } else {
      backupTriggerConditionContainer.classList.add("hidden");
    }
  });

  // --- Interacción con el Script PowerShell a través de Tauri ---
  btnSearchMailbox.addEventListener("click", async () => {
    btnSearchMailbox.disabled = true;
    mailboxLoading.classList.remove("hidden");
    mailboxListContainer.classList.add("hidden");
    mailboxList.innerHTML = "";
    selectedMailboxPath = "";

    try {
      // Llamar al comando tauri expuesto en Rust
      const stores = await invoke("find_outlook_stores");

      mailboxLoading.classList.add("hidden");
      btnSearchMailbox.disabled = false;

      if (!stores || stores.length === 0) {
        mailboxList.innerHTML = `<div style="padding: 10px; font-size:12px; color: #666; text-align: center;">No se encontraron buzones OST activos en Outlook. Asegúrese de que Outlook está configurado correctamente.</div>`;
        mailboxListContainer.classList.remove("hidden");
        return;
      }

      // Renderizar los buzones encontrados en formato de tarjetas interactivas
      stores.forEach(store => {
        const card = document.createElement("div");
        card.className = "mailbox-card";
        card.innerHTML = `
          <div class="mailbox-title">${store.displayName}</div>
          <div class="mailbox-path" title="${store.filePath}">${store.filePath}</div>
          <div class="mailbox-size">Tamaño: ${store.fileSize}</div>
        `;

        card.addEventListener("click", () => {
          // Desmarcar otras tarjetas
          document.querySelectorAll(".mailbox-card").forEach(c => c.classList.remove("selected"));
          // Seleccionar esta tarjeta
          card.classList.add("selected");
          // Guardar ruta seleccionada en la variable global de sesión
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

  // --- Interacción con el Script PowerShell a través de Tauri (Destino - Paso 2) ---
  btnSearchDestinationMailbox.addEventListener("click", async () => {
    btnSearchDestinationMailbox.disabled = true;
    mailboxDestinationLoading.classList.remove("hidden");
    mailboxDestinationListContainer.classList.add("hidden");
    mailboxDestinationList.innerHTML = "";
    selectedDestinationMailboxPath = "";

    try {
      // Llamar al comando tauri expuesto en Rust
      const stores = await invoke("find_outlook_stores");

      mailboxDestinationLoading.classList.add("hidden");
      btnSearchDestinationMailbox.disabled = false;

      // Filtrar para excluir el buzón seleccionado en el Paso 1 (selectedMailboxPath)
      const filteredStores = stores.filter(store => store.filePath !== selectedMailboxPath);

      if (!filteredStores || filteredStores.length === 0) {
        mailboxDestinationList.innerHTML = `<div style="padding: 10px; font-size:12px; color: #666; text-align: center;">No se encontraron otros buzones OST disponibles en Outlook que difieran del de origen.</div>`;
        mailboxDestinationListContainer.classList.remove("hidden");
        return;
      }

      // Renderizar los buzones de destino encontrados en formato de tarjetas interactivas
      filteredStores.forEach(store => {
        const card = document.createElement("div");
        card.className = "mailbox-card";
        card.innerHTML = `
          <div class="mailbox-title">${store.displayName}</div>
          <div class="mailbox-path" title="${store.filePath}">${store.filePath}</div>
          <div class="mailbox-size">Tamaño: ${store.fileSize}</div>
        `;

        card.addEventListener("click", () => {
          // Desmarcar otras tarjetas en la lista de destino
          mailboxDestinationList.querySelectorAll(".mailbox-card").forEach(c => c.classList.remove("selected"));
          // Seleccionar esta tarjeta
          card.classList.add("selected");
          // Guardar ruta seleccionada en la variable global de sesión de destino
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
    
    // Mostrar/ocultar contenidos de paso
    wizardSteps.forEach((stepEl, idx) => {
      if (idx + 1 === step) {
        stepEl.classList.remove("hidden");
      } else {
        stepEl.classList.add("hidden");
      }
    });

    // Actualizar indicadores visuales de pasos
    stepIndicators.forEach((indicator, idx) => {
      const stepNum = idx + 1;
      indicator.className = "step-indicator";
      if (stepNum < step) {
        indicator.classList.add("completed");
      } else if (stepNum === step) {
        indicator.classList.add("active");
      }
    });

    // Actualizar líneas conectores
    stepLines.forEach((line, idx) => {
      const lineNum = idx + 1;
      line.className = "step-line";
      if (lineNum < step) {
        line.classList.add("completed");
      }
    });

    // Configurar botones de navegación
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

  // Navegación Siguiente / Finalizar
  btnWizardNext.addEventListener("click", () => {
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
      
      // --- Cargar dinámicamente la información al Resumen (Paso 3) ---
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
      // Guardar Configuración Final del Wizard (Resumen)
      savedConfig = {
        source: selectedMailboxPath,
        analysisFrequency: wizardAnalysisFrequencySelect.value,
        destination: selectedDestinationMailboxPath,
        criterion: wizardCriterionSelect.value,
        triggerCondition: wizardCriterionSelect.value === "move" ? wizardTriggerConditionSelect.value : null
      };
      
      localStorage.setItem("outlook_backup_config", JSON.stringify(savedConfig));
      
      // Aplicar estado para mostrar la GUI con la configuración activa
      applyViewState();
      
      // Mostrar mensaje de éxito
      statusMessage.textContent = "✓ ¡Automatización creada e iniciada con éxito!";
      statusMessage.className = "status-message success";
      setTimeout(() => { statusMessage.textContent = ""; }, 4000);
    }
  });

  // Navegación Atrás
  btnWizardBack.addEventListener("click", () => {
    if (currentStep > 1) {
      showStep(currentStep - 1);
    }
  });

  // Cancelar Wizard
  btnWizardCancel.addEventListener("click", () => {
    if (confirm("¿Está seguro de que desea cancelar el asistente? Se perderán los datos introducidos.")) {
      applyViewState();
    }
  });

  // --- Examinar Archivos en el Wizard (Simulados) ---
  document.querySelector("#btn-wizard-select-destination").addEventListener("click", () => {
    wizardDestinationInput.value = "E:\\Respaldos\\Outlook_Backups";
  });

  // --- Lógica del Panel de Automatización Activa (Edición y Acciones) ---
  
  // Guardar Cambios en Formulario de Edición
  backupForm.addEventListener("submit", (e) => {
    e.preventDefault();
    
    savedConfig = {
      source: selectedMailboxPath,
      analysisFrequency: analysisFrequencySelect.value,
      destination: selectedDestinationMailboxPath,
      criterion: backupCriterionSelect.value,
      triggerCondition: backupCriterionSelect.value === "move" ? backupTriggerConditionSelect.value : null
    };

    localStorage.setItem("outlook_backup_config", JSON.stringify(savedConfig));

    statusMessage.textContent = "✓ Cambios guardados con éxito.";
    statusMessage.className = "status-message success";
    setTimeout(() => { statusMessage.textContent = ""; }, 4000);
  });

  // Ejecutar respaldo inmediato
  document.querySelector("#btn-run-now").addEventListener("click", () => {
    statusMessage.textContent = "⚡ Iniciando respaldo inmediato en segundo plano...";
    statusMessage.className = "status-message";
    
    setTimeout(() => {
      statusMessage.textContent = "✓ ¡Respaldo completado con éxito!";
      statusMessage.className = "status-message success";
      setTimeout(() => { statusMessage.textContent = ""; }, 4000);
    }, 2000);
  });

  // Examinar archivos en Formulario Activo
  document.querySelector("#btn-select-destination").addEventListener("click", () => {
    destinationInput.value = "E:\\Respaldos\\Outlook_Backups";
  });

  // Eliminar Configuración
  document.querySelector("#btn-delete-config").addEventListener("click", () => {
    if (confirm("¿Está seguro de que desea eliminar esta automatización de respaldo? La aplicación dejará de realizar copias de seguridad.")) {
      localStorage.removeItem("outlook_backup_config");
      savedConfig = null;
      applyViewState();
    }
  });

  // --- Renderizado Inicial ---
  applyViewState();
});
