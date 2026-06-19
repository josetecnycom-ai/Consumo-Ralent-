/* main.js - Controlador principal y ciclo de vida de MyGeotab Add-in */

// Crear espacio de nombres si no existe (para desarrollo local)
if (typeof geotab === 'undefined') {
  window.geotab = { addin: {} };
}

geotab.addin.consumoRalenti = function (api, state) {
  // Instancias y estados internos
  let allDevices = [];
  let selectedDevices = [];
  let calculatedSummary = [];
  let calculatedDetails = [];
  let activeChartInstance = null;
  let selectedChartVehicleId = null;
  let lastRawData = null;

  // Elementos del DOM
  let elDatePreset, elDateFrom, elDateTo, elGroup, elSearchVehicle, elBtnUpdate, elBtnExport;
  let elTableBody, elStatsVehicles, elStatsFuel, elStatsIdle, elStatsKm;
  let elMultiSelectDisplay, elMultiSelectDropdown, elMultiSelectSearch, elMultiSelectList;
  let elSelectAllVehicles, elDeselectAllVehicles;
  let elLoader, elEmptyState, elChartTitle, elChartSelectVehicle;

  // Inicialización de variables del DOM
  function cacheDomElements() {
    elDatePreset = document.getElementById("date-preset");
    elDateFrom = document.getElementById("date-from");
    elDateTo = document.getElementById("date-to");
    elGroup = document.getElementById("group-by");
    elSearchVehicle = document.getElementById("search-vehicle");
    elBtnUpdate = document.getElementById("btn-update");
    elBtnExport = document.getElementById("btn-export");
    
    elTableBody = document.getElementById("table-body");
    elStatsVehicles = document.getElementById("stat-vehicles-count");
    elStatsFuel = document.getElementById("stat-total-fuel");
    elStatsIdle = document.getElementById("stat-total-idle");
    elStatsKm = document.getElementById("stat-total-km");

    elMultiSelectDisplay = document.getElementById("multiselect-display");
    elMultiSelectDropdown = document.getElementById("multiselect-dropdown");
    elMultiSelectSearch = document.getElementById("multiselect-search");
    elMultiSelectList = document.getElementById("multiselect-list");
    elSelectAllVehicles = document.getElementById("select-all-vehicles");
    elDeselectAllVehicles = document.getElementById("deselect-all-vehicles");

    elLoader = document.getElementById("loader-overlay");
    elEmptyState = document.getElementById("empty-state");
    elChartTitle = document.getElementById("chart-title");
    elChartSelectVehicle = document.getElementById("chart-select-vehicle");
  }

  // Cargar presets de fecha
  function handleDatePresetChange() {
    const preset = elDatePreset.value;
    const now = new Date();
    let fromDate = new Date();
    let toDate = new Date();

    // Resetear horas
    fromDate.setHours(0, 0, 0, 0);
    toDate.setHours(23, 59, 59, 999);

    switch (preset) {
      case "today":
        break;
      case "yesterday":
        fromDate.setDate(now.getDate() - 1);
        toDate.setDate(now.getDate() - 1);
        break;
      case "thisWeek":
        // Obtener lunes de esta semana
        const day = now.getDay();
        const diff = now.getDate() - day + (day === 0 ? -6 : 1);
        fromDate.setDate(diff);
        break;
      case "lastWeek":
        const prevMonday = new Date();
        prevMonday.setDate(now.getDate() - now.getDay() - 6);
        const prevSunday = new Date();
        prevSunday.setDate(now.getDate() - now.getDay());
        fromDate = prevMonday;
        toDate = prevSunday;
        break;
      case "thisMonth":
        fromDate = new Date(now.getFullYear(), now.getMonth(), 1);
        break;
      case "lastMonth":
        fromDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        toDate = new Date(now.getFullYear(), now.getMonth(), 0);
        break;
      case "custom":
        return; // No sobreescribir las fechas manuales
    }

    elDateFrom.value = formatDateToISOString(fromDate);
    elDateTo.value = formatDateToISOString(toDate);
  }

  function formatDateToISOString(date) {
    const tzoffset = date.getTimezoneOffset() * 60000; // Offset en milisegundos
    const localISOTime = (new Date(date.getTime() - tzoffset)).toISOString().slice(0, 16);
    return localISOTime;
  }

  // Parsear TimeSpan de Geotab a segundos
  function parseTimeSpanToSeconds(ts) {
    if (!ts) return 0;
    if (typeof ts === 'number') return ts;

    // Quitar milisegundos si existen
    let cleanTs = ts.split('.')[0];
    let days = 0;

    // Verificar si contiene días (ej: "1.12:34:56")
    if (ts.includes('.')) {
      const dotParts = ts.split('.');
      if (dotParts.length > 1 && dotParts[0].match(/^\d+$/) && dotParts[1].includes(':')) {
        days = parseInt(dotParts[0], 10);
        cleanTs = dotParts[1];
      }
    }

    const parts = cleanTs.split(':');
    if (parts.length === 3) {
      const h = parseInt(parts[0], 10) || 0;
      const m = parseInt(parts[1], 10) || 0;
      const s = parseInt(parts[2], 10) || 0;
      return (days * 86400) + (h * 3600) + (m * 60) + s;
    }
    return 0;
  }

  // Buscar valor de diagnóstico interpolado temporalmente
  function getValueAtTimestamp(records, deviceId, diagnosticId, timestamp) {
    const deviceRecords = records.filter(r => r.device.id === deviceId && r.diagnostic.id === diagnosticId);
    if (deviceRecords.length === 0) return null;

    // Asegurar orden por fecha
    deviceRecords.sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));

    const targetTime = new Date(timestamp).getTime();

    let beforeRecord = null;
    let afterRecord = null;

    for (let i = 0; i < deviceRecords.length; i++) {
      const recTime = new Date(deviceRecords[i].dateTime).getTime();
      if (recTime <= targetTime) {
        beforeRecord = deviceRecords[i];
      } else {
        afterRecord = deviceRecords[i];
        break;
      }
    }

    if (beforeRecord && afterRecord) {
      const tA = new Date(beforeRecord.dateTime).getTime();
      const tB = new Date(afterRecord.dateTime).getTime();
      const vA = beforeRecord.data;
      const vB = afterRecord.data;

      if (tB === tA) return vA;
      return vA + (vB - vA) * (targetTime - tA) / (tB - tA);
    } else if (beforeRecord) {
      return beforeRecord.data;
    } else if (afterRecord) {
      return afterRecord.data;
    }
    return null;
  }

  // Generar la lista de intervalos (bins) de tiempo basados en la agrupación
  function generateTimeBins(fromDate, toDate, groupingType) {
    const bins = [];
    const start = new Date(fromDate);
    const end = new Date(toDate);

    if (groupingType === "none") {
      bins.push({
        id: "total",
        label: "Total del Periodo",
        start: start.toISOString(),
        end: end.toISOString()
      });
      return bins;
    }

    const current = new Date(start);
    
    while (current <= end) {
      let binStart, binEnd, binLabel;
      const y = current.getFullYear();
      const m = current.getMonth();
      const d = current.getDate();

      if (groupingType === "day") {
        binStart = new Date(y, m, d, 0, 0, 0, 0);
        binEnd = new Date(y, m, d, 23, 59, 59, 999);
        binLabel = current.toLocaleDateString("es-ES", { day: '2-digit', month: '2-digit', year: 'numeric' });
        
        // Siguiente día
        current.setDate(d + 1);
      } 
      else if (groupingType === "week") {
        // Obtener lunes de la semana actual
        const dayOfWeek = current.getDay();
        const mondayDiff = current.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
        
        binStart = new Date(y, m, mondayDiff, 0, 0, 0, 0);
        binEnd = new Date(y, m, mondayDiff + 6, 23, 59, 59, 999);
        
        // Ajustar al rango de consulta original si se sale de los bordes
        if (binStart < start) binStart = new Date(start);
        if (binEnd > end) binEnd = new Date(end);

        // Formato etiqueta: "Semana DD/MM"
        const formatOptions = { day: '2-digit', month: '2-digit' };
        binLabel = `Sem. ${binStart.toLocaleDateString("es-ES", formatOptions)} al ${binEnd.toLocaleDateString("es-ES", formatOptions)}`;

        // Siguiente semana
        current.setDate(mondayDiff + 7);
      } 
      else if (groupingType === "month") {
        binStart = new Date(y, m, 1, 0, 0, 0, 0);
        binEnd = new Date(y, m + 1, 0, 23, 59, 59, 999);
        
        if (binStart < start) binStart = new Date(start);
        if (binEnd > end) binEnd = new Date(end);

        binLabel = current.toLocaleDateString("es-ES", { month: 'long', year: 'numeric' });
        binLabel = binLabel.charAt(0).toUpperCase() + binLabel.slice(1); // Capitalizar

        // Siguiente mes
        current.setMonth(m + 1);
        current.setDate(1);
      }

      // Evitar bins redundantes que superen el límite final estricto
      if (binStart <= end) {
        bins.push({
          id: binStart.getTime().toString(),
          label: binLabel,
          start: binStart.toISOString(),
          end: binEnd.toISOString()
        });
      }
    }

    return bins;
  }

  // Realizar los cálculos matemáticos agrupados
  function processCalculations(devices, trips, statusData, fromDate, toDate, groupingType) {
    const bins = generateTimeBins(fromDate, toDate, groupingType);
    const summaryList = [];
    const detailsList = [];

    devices.forEach(device => {
      // 1. Cálculos generales acumulados (Resumen de vehículo para todo el periodo)
      const devTrips = trips.filter(t => t.device.id === device.id);
      
      let totalDistance = 0;
      let totalActivity = 0;
      let totalIdle = 0;

      devTrips.forEach(t => {
        totalDistance += t.distance || 0;
        totalActivity += parseTimeSpanToSeconds(t.drivingDuration) + parseTimeSpanToSeconds(t.idlingDuration);
        totalIdle += parseTimeSpanToSeconds(t.idlingDuration);
      });

      // Combustible total del periodo completo
      const fuelStart = getValueAtTimestamp(statusData, device.id, GeotabApiService.DIAGNOSTICS.totalFuel, fromDate);
      const fuelEnd = getValueAtTimestamp(statusData, device.id, GeotabApiService.DIAGNOSTICS.totalFuel, toDate);
      let periodTotalFuel = null;
      if (fuelStart !== null && fuelEnd !== null) {
        periodTotalFuel = Math.max(0, fuelEnd - fuelStart);
      }

      // Combustible ralentí del periodo completo
      const idleFuelStart = getValueAtTimestamp(statusData, device.id, GeotabApiService.DIAGNOSTICS.totalIdleFuel, fromDate);
      const idleFuelEnd = getValueAtTimestamp(statusData, device.id, GeotabApiService.DIAGNOSTICS.totalIdleFuel, toDate);
      let periodIdleFuel = null;
      if (idleFuelStart !== null && idleFuelEnd !== null) {
        periodIdleFuel = Math.max(0, idleFuelEnd - idleFuelStart);
      }

      const idlePctTotal = totalActivity > 0 ? (totalIdle / totalActivity) * 100 : 0;

      summaryList.push({
        deviceId: device.id,
        deviceName: device.name,
        activityDuration: totalActivity,
        idleDuration: totalIdle,
        idlePercentage: idlePctTotal,
        idleFuel: periodIdleFuel,
        totalFuel: periodTotalFuel,
        distance: totalDistance // Ya está en Km
      });

      // 2. Cálculos desglosados por bin de tiempo (Detalles)
      bins.forEach(bin => {
        const binTrips = devTrips.filter(t => {
          const tripStart = new Date(t.start);
          return tripStart >= new Date(bin.start) && tripStart <= new Date(bin.end);
        });

        let binDistance = 0;
        let binActivity = 0;
        let binIdle = 0;

        binTrips.forEach(t => {
          binDistance += t.distance || 0;
          binActivity += parseTimeSpanToSeconds(t.drivingDuration) + parseTimeSpanToSeconds(t.idlingDuration);
          binIdle += parseTimeSpanToSeconds(t.idlingDuration);
        });

        // Combustible del bin
        const binFuelStart = getValueAtTimestamp(statusData, device.id, GeotabApiService.DIAGNOSTICS.totalFuel, bin.start);
        const binFuelEnd = getValueAtTimestamp(statusData, device.id, GeotabApiService.DIAGNOSTICS.totalFuel, bin.end);
        let binTotalFuel = null;
        if (binFuelStart !== null && binFuelEnd !== null) {
          binTotalFuel = Math.max(0, binFuelEnd - binFuelStart);
        }

        // Combustible en ralentí del bin
        const binIdleFuelStart = getValueAtTimestamp(statusData, device.id, GeotabApiService.DIAGNOSTICS.totalIdleFuel, bin.start);
        const binIdleFuelEnd = getValueAtTimestamp(statusData, device.id, GeotabApiService.DIAGNOSTICS.totalIdleFuel, bin.end);
        let binIdleFuel = null;
        if (binIdleFuelStart !== null && binIdleFuelEnd !== null) {
          binIdleFuel = Math.max(0, binIdleFuelEnd - binIdleFuelStart);
        }

        const binIdlePct = binActivity > 0 ? (binIdle / binActivity) * 100 : 0;

        detailsList.push({
          period: bin.label,
          periodStart: bin.start,
          deviceId: device.id,
          deviceName: device.name,
          activityDuration: binActivity,
          idleDuration: binIdle,
          idlePercentage: binIdlePct,
          idleFuel: binIdleFuel,
          totalFuel: binTotalFuel,
          distance: binDistance // Ya está en Km
        });
      });
    });

    return { summary: summaryList, details: detailsList };
  }

  // Formatear segundos a string legible (ej. "15h 23m 10s" o "02:15:30")
  function formatSeconds(totalSeconds) {
    const hrs = Math.floor(totalSeconds / 3600);
    const mins = Math.floor((totalSeconds % 3600) / 60);
    const secs = Math.floor(totalSeconds % 60);
    
    if (hrs > 0) {
      return `${hrs}h ${mins}m`;
    }
    return `${mins}m ${secs}s`;
  }

  // Dibujar tabla de resultados
  function renderTable(data, groupingType) {
    elTableBody.innerHTML = "";

    if (data.length === 0) {
      elTableBody.innerHTML = `<tr><td colspan="7" class="text-center">No se encontraron datos para la selección.</td></tr>`;
      return;
    }

    data.forEach(item => {
      const isHighIdle = item.idlePercentage > 30 ? "high" : item.idlePercentage > 15 ? "medium" : "low";
      const idlePctClass = `percentage-badge ${isHighIdle}`;

      const row = document.createElement("tr");
      row.style.cursor = "pointer";
      row.dataset.deviceId = item.deviceId;
      
      if (selectedChartVehicleId === item.deviceId) {
        row.classList.add("selected-row");
      }

      row.innerHTML = `
        <td><strong>${item.deviceName}</strong></td>
        <td class="text-right">${formatSeconds(item.activityDuration)}</td>
        <td class="text-right">${formatSeconds(item.idleDuration)}</td>
        <td class="text-center"><span class="${idlePctClass}">${item.idlePercentage.toFixed(1)}%</span></td>
        <td class="text-right">${item.idleFuel !== null ? item.idleFuel.toFixed(1) + ' L' : '<span class="na-value">N/A</span>'}</td>
        <td class="text-right">${item.totalFuel !== null ? item.totalFuel.toFixed(1) + ' L' : '<span class="na-value">N/A</span>'}</td>
        <td class="text-right">${item.distance.toFixed(1)} km</td>
      `;

      // Evento de clic en fila para cargar gráfica del vehículo
      row.addEventListener("click", () => {
        document.querySelectorAll("#table-body tr").forEach(r => r.classList.remove("selected-row"));
        row.classList.add("selected-row");
        selectedChartVehicleId = item.deviceId;

        // Si la agrupación es "none", la cambiamos automáticamente a "day" para poder ver la evolución
        if (elGroup.value === "none") {
          elGroup.value = "day";
          
          if (lastRawData) {
            const fromDate = new Date(elDateFrom.value).toISOString();
            const toDate = new Date(elDateTo.value).toISOString();
            const results = processCalculations(
              allDevices.filter(d => selectedDevices.includes(d.id)),
              lastRawData.trips,
              lastRawData.statusData,
              fromDate,
              toDate,
              "day"
            );
            calculatedSummary = results.summary;
            calculatedDetails = results.details;
          }
        }

        elChartSelectVehicle.value = selectedChartVehicleId;
        updateChartSelect();
        renderCharts();
      });

      elTableBody.appendChild(row);
    });
  }

  // Renderizar tarjetas estadísticas de resumen
  function renderStats(summaryData) {
    const validFuelVehicles = summaryData.filter(v => v.idleFuel !== null);
    
    const count = summaryData.length;
    const totalIdleTime = summaryData.reduce((acc, v) => acc + v.idleDuration, 0);
    const totalKm = summaryData.reduce((acc, v) => acc + v.distance, 0);
    const totalIdleFuel = validFuelVehicles.reduce((acc, v) => acc + v.idleFuel, 0);
    const totalFuelUsed = validFuelVehicles.reduce((acc, v) => acc + v.totalFuel, 0);

    elStatsVehicles.textContent = count;
    elStatsKm.textContent = `${totalKm.toLocaleString("es-ES", { maximumFractionDigits: 1 })} km`;
    
    elStatsIdle.textContent = formatSeconds(totalIdleTime);

    if (validFuelVehicles.length > 0) {
      elStatsFuel.textContent = `${totalIdleFuel.toLocaleString("es-ES", { maximumFractionDigits: 1 })} L`;
    } else {
      elStatsFuel.innerHTML = '<span class="na-value">N/A</span>';
    }
  }

  // Dibujar las gráficas de consumo en ralentí
  function renderCharts() {
    if (activeChartInstance) {
      activeChartInstance.destroy();
      activeChartInstance = null;
    }

    const canvas = document.getElementById("chart-canvas");
    const ctx = canvas.getContext("2d");

    const groupingType = elGroup.value;

    // Caso 1: Agrupación "Ninguna" -> Gráfico comparativo de barras de todos los vehículos
    if (groupingType === "none" || !selectedChartVehicleId) {
      elChartSelectVehicle.style.display = "none";
      elChartTitle.textContent = "Comparación de Consumo en Ralentí por Vehículo (Litros)";

      // Filtrar vehículos que tengan datos de combustible
      const dataWithFuel = calculatedSummary.filter(v => v.idleFuel !== null);

      if (dataWithFuel.length === 0) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        // Escribir mensaje en canvas
        ctx.font = "14px Arial";
        ctx.fillStyle = "#666666";
        ctx.textAlign = "center";
        ctx.fillText("No hay datos de consumo en ralentí disponibles para graficar.", canvas.width / 2, canvas.height / 2);
        return;
      }

      const labels = dataWithFuel.map(v => v.deviceName);
      const idleFuelData = dataWithFuel.map(v => v.idleFuel);
      const totalFuelData = dataWithFuel.map(v => v.totalFuel);

      activeChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            {
              label: 'Combustible Ralentí (Litros)',
              data: idleFuelData,
              backgroundColor: '#006ED5',
              borderColor: '#002D62',
              borderWidth: 1
            },
            {
              label: 'Combustible Total (Litros)',
              data: totalFuelData,
              backgroundColor: '#E2E8F0',
              borderColor: '#9CA3AF',
              borderWidth: 1
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: {
              beginAtZero: true,
              title: { display: true, text: 'Litros (L)' }
            }
          },
          plugins: {
            legend: { position: 'bottom' }
          }
        }
      });
    } 
    // Caso 2: Con agrupación -> Gráfico de evolución temporal del vehículo seleccionado
    else {
      elChartSelectVehicle.style.display = "inline-block";
      const vehicle = allDevices.find(d => d.id === selectedChartVehicleId);
      elChartTitle.textContent = `Evolución de Consumo en Ralentí: ${vehicle ? vehicle.name : ''}`;

      // Obtener datos detallados de este vehículo
      const vehicleDetails = calculatedDetails.filter(d => d.deviceId === selectedChartVehicleId);
      
      // Ordenar por fecha
      vehicleDetails.sort((a, b) => new Date(a.periodStart) - new Date(b.periodStart));

      const labels = vehicleDetails.map(d => d.period);
      const idleFuelData = vehicleDetails.map(d => d.idleFuel);
      const reportsFuel = vehicleDetails.some(d => d.idleFuel !== null);

      if (!reportsFuel) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.font = "14px Arial";
        ctx.fillStyle = "#9CA3AF";
        ctx.textAlign = "center";
        ctx.fillText("Este vehículo no reporta datos de consumo de combustible.", canvas.width / 2, canvas.height / 2);
        return;
      }

      activeChartInstance = new Chart(ctx, {
        type: 'line',
        data: {
          labels: labels,
          datasets: [
            {
              label: 'Consumo Ralentí (Litros)',
              data: idleFuelData,
              borderColor: '#006ED5',
              backgroundColor: 'rgba(0, 110, 213, 0.1)',
              borderWidth: 3,
              fill: true,
              tension: 0.3,
              pointBackgroundColor: '#002D62',
              pointRadius: 5
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: {
              beginAtZero: true,
              title: { display: true, text: 'Litros (L)' }
            }
          },
          plugins: {
            legend: { position: 'bottom' }
          }
        }
      });
    }
  }

  // Actualizar el selector de vehículos del gráfico
  function updateChartSelect() {
    elChartSelectVehicle.innerHTML = "";
    
    // Si la agrupación es None, ocultar el selector
    if (elGroup.value === "none") {
      elChartSelectVehicle.style.display = "none";
      return;
    }

    elChartSelectVehicle.style.display = "inline-block";
    
    // Añadir opción por defecto
    const defOpt = document.createElement("option");
    defOpt.value = "";
    defOpt.textContent = "-- Seleccionar Vehículo --";
    elChartSelectVehicle.appendChild(defOpt);

    // Añadir vehículos que tengan datos en la consulta
    calculatedSummary.forEach(item => {
      const opt = document.createElement("option");
      opt.value = item.deviceId;
      opt.textContent = item.deviceName;
      if (item.deviceId === selectedChartVehicleId) {
        opt.selected = true;
      }
      elChartSelectVehicle.appendChild(opt);
    });
  }

  // Filtrado del dropdown de vehículos
  function filterDropdownVehicles() {
    const query = elMultiSelectSearch.value.toLowerCase().trim();
    const items = elMultiSelectList.querySelectorAll(".multiselect-item");

    items.forEach(item => {
      const name = item.textContent.toLowerCase();
      if (name.includes(query)) {
        item.style.display = "flex";
      } else {
        item.style.display = "none";
      }
    });
  }

  // Actualizar la cabecera del dropdown multiselect
  function updateMultiSelectDisplay() {
    const selectedCheckboxes = elMultiSelectList.querySelectorAll("input[type='checkbox']:checked");
    const displaySpan = elMultiSelectDisplay.querySelector("span");
    
    if (selectedCheckboxes.length === 0) {
      displaySpan.textContent = "Ningún vehículo seleccionado";
      selectedDevices = [];
    } else if (selectedCheckboxes.length === allDevices.length) {
      displaySpan.textContent = "Todos los vehículos";
      selectedDevices = allDevices.map(d => d.id);
    } else {
      displaySpan.textContent = `${selectedCheckboxes.length} vehículos seleccionados`;
      selectedDevices = Array.from(selectedCheckboxes).map(cb => cb.value);
    }
  }

  // Llenar la lista del dropdown multiselect
  function populateVehicleDropdown(devices) {
    elMultiSelectList.innerHTML = "";
    selectedDevices = [];

    devices.forEach(device => {
      const div = document.createElement("div");
      div.className = "multiselect-item";
      
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = device.id;
      checkbox.checked = true; // Seleccionado por defecto
      checkbox.id = `cb-veh-${device.id}`;

      const label = document.createElement("label");
      label.htmlFor = `cb-veh-${device.id}`;
      label.textContent = device.name;
      label.style.cursor = "pointer";
      label.style.width = "100%";

      div.appendChild(checkbox);
      div.appendChild(label);

      // Evento al cambiar selección individual
      checkbox.addEventListener("change", () => {
        updateMultiSelectDisplay();
      });

      elMultiSelectList.appendChild(div);
      selectedDevices.push(device.id);
    });

    updateMultiSelectDisplay();
  }

  // Consultar y actualizar datos principales
  function queryData() {
    if (selectedDevices.length === 0) {
      alert("Por favor, seleccione al menos un vehículo.");
      return;
    }

    elLoader.classList.add("active");
    elEmptyState.style.display = "none";
    
    const fromDate = new Date(elDateFrom.value).toISOString();
    const toDate = new Date(elDateTo.value).toISOString();
    const grouping = elGroup.value;

    GeotabApiService.getData(selectedDevices, fromDate, toDate, function (data) {
      elLoader.classList.remove("active");
      lastRawData = data;
      
      const results = processCalculations(
        allDevices.filter(d => selectedDevices.includes(d.id)),
        data.trips,
        data.statusData,
        fromDate,
        toDate,
        grouping
      );

      calculatedSummary = results.summary;
      calculatedDetails = results.details;

      // Si no tenemos un vehículo seleccionado para el gráfico, o el actual no está en la consulta, asignamos el primero
      if (calculatedSummary.length > 0) {
        if (!selectedChartVehicleId || !selectedDevices.includes(selectedChartVehicleId)) {
          // Intentar elegir uno que reporte combustible
          const reporting = calculatedSummary.find(v => v.idleFuel !== null);
          selectedChartVehicleId = reporting ? reporting.deviceId : calculatedSummary[0].deviceId;
        }
      } else {
        selectedChartVehicleId = null;
      }

      // Renderizar elementos de UI
      renderTable(calculatedSummary, grouping);
      renderStats(calculatedSummary);
      updateChartSelect();
      renderCharts();
    });
  }

  // Exportación a Excel
  function exportData() {
    if (calculatedSummary.length === 0) {
      alert("No hay datos cargados para exportar.");
      return;
    }

    const fromDateStr = new Date(elDateFrom.value).toLocaleDateString("es-ES");
    const toDateStr = new Date(elDateTo.value).toLocaleDateString("es-ES");
    const periodText = `${fromDateStr}_a_${toDateStr}`;
    const grouping = elGroup.value;

    ExcelExporter.exportConsumoRalenti(calculatedSummary, calculatedDetails, periodText, grouping);
  }

  return {
    initialize: function (api, state, initializeCallback) {
      // Registrar el servicio
      GeotabApiService.init(api, state);

      cacheDomElements();

      // Configurar fechas iniciales (preset "Este Mes" por defecto)
      elDatePreset.value = "thisMonth";
      handleDatePresetChange();

      // Enlazar eventos de fechas
      elDatePreset.addEventListener("change", handleDatePresetChange);
      
      elDateFrom.addEventListener("change", () => {
        elDatePreset.value = "custom";
      });
      elDateTo.addEventListener("change", () => {
        elDatePreset.value = "custom";
      });

      // Enlazar eventos del multiselect
      elMultiSelectDisplay.addEventListener("click", (e) => {
        e.stopPropagation();
        elMultiSelectDropdown.classList.toggle("open");
      });

      // Cerrar dropdown al hacer clic fuera
      document.addEventListener("click", (e) => {
        if (!elMultiSelectDropdown.contains(e.target)) {
          elMultiSelectDropdown.classList.remove("open");
        }
      });

      elMultiSelectSearch.addEventListener("input", filterDropdownVehicles);

      elSelectAllVehicles.addEventListener("click", (e) => {
        e.preventDefault();
        elMultiSelectList.querySelectorAll("input[type='checkbox']").forEach(cb => {
          if (cb.parentNode.style.display !== "none") cb.checked = true;
        });
        updateMultiSelectDisplay();
      });

      elDeselectAllVehicles.addEventListener("click", (e) => {
        e.preventDefault();
        elMultiSelectList.querySelectorAll("input[type='checkbox']").forEach(cb => {
          if (cb.parentNode.style.display !== "none") cb.checked = false;
        });
        updateMultiSelectDisplay();
      });

      // Selector de vehículo del gráfico
      elChartSelectVehicle.addEventListener("change", () => {
        selectedChartVehicleId = elChartSelectVehicle.value;
        
        // Resaltar en la tabla
        document.querySelectorAll("#table-body tr").forEach(row => {
          if (row.dataset.deviceId === selectedChartVehicleId) {
            row.classList.add("selected-row");
          } else {
            row.classList.remove("selected-row");
          }
        });

        renderCharts();
      });

      // Evento de cambio de agrupación temporal
      elGroup.addEventListener("change", () => {
        if (lastRawData) {
          const fromDate = new Date(elDateFrom.value).toISOString();
          const toDate = new Date(elDateTo.value).toISOString();
          const grouping = elGroup.value;

          const results = processCalculations(
            allDevices.filter(d => selectedDevices.includes(d.id)),
            lastRawData.trips,
            lastRawData.statusData,
            fromDate,
            toDate,
            grouping
          );

          calculatedSummary = results.summary;
          calculatedDetails = results.details;

          updateChartSelect();
          renderCharts();
        }
      });

      // Botones principales
      elBtnUpdate.addEventListener("click", queryData);
      elBtnExport.addEventListener("click", exportData);

      // Señalar que la inicialización ha concluido
      initializeCallback();
    },

    focus: function (api, state) {
      elLoader.classList.add("active");
      
      // Obtener vehículos
      GeotabApiService.getDevices(function (devices) {
        elLoader.classList.remove("active");
        allDevices = devices;
        
        if (devices.length === 0) {
          elEmptyState.style.display = "flex";
          return;
        }

        populateVehicleDropdown(devices);
        
        // Cargar datos por primera vez
        queryData();
      });
    },

    blur: function (api, state) {
      console.log("Add-In blur");
    }
  };
};

// Auto-arranque si se ejecuta en navegador fuera del entorno MyGeotab (desarrollo local)
if (typeof api === 'undefined' || !api) {
  document.addEventListener("DOMContentLoaded", function () {
    // Verificar si no estamos en iframe / MyGeotab
    if (window.location.protocol !== 'https:' || !window.geotab.addin.consumoRalenti.length) {
      const addinInstance = geotab.addin.consumoRalenti();
      GeotabApiService.init(null, null); // Iniciar en modo simulación
      
      const mockState = {
        translate: function (text) { return text; }
      };

      addinInstance.initialize(null, mockState, function () {
        addinInstance.focus(null, mockState);
      });
    }
  });
}
