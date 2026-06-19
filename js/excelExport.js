/* excelExport.js - Exportador a Excel en dos pestañas estructuradas utilizando SheetJS */

const ExcelExporter = (function () {
  
  // Función para auto-ajustar el ancho de las columnas
  function autoFitColumns(ws) {
    if (!ws || !ws['!ref']) return;
    
    const range = XLSX.utils.decode_range(ws['!ref']);
    const colWidths = [];
    
    for (let col = range.s.c; col <= range.e.c; col++) {
      let maxLen = 12; // Ancho mínimo por defecto
      for (let row = range.s.r; row <= range.e.r; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = ws[cellAddress];
        if (cell && cell.v !== undefined && cell.v !== null) {
          const valStr = String(cell.v);
          if (valStr.length > maxLen) {
            maxLen = valStr.length;
          }
        }
      }
      colWidths.push({ wch: maxLen + 3 }); // Añadimos margen extra
    }
    ws['!cols'] = colWidths;
  }

  // Formatear segundos a formato de duración HH:MM:SS
  function formatSecondsToHMS(totalSeconds) {
    if (isNaN(totalSeconds) || totalSeconds === null) return "00:00:00";
    const hrs = Math.floor(totalSeconds / 3600);
    const mins = Math.floor((totalSeconds % 3600) / 60);
    const secs = Math.floor(totalSeconds % 60);
    return `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  }

  return {
    exportConsumoRalenti: function (summaryData, detailData, periodText, groupingType) {
      if (typeof XLSX === 'undefined') {
        alert("La librería de Excel (SheetJS) aún no se ha cargado. Por favor, espere a que se inicialice.");
        return;
      }

      // 1. Crear libro de trabajo
      const wb = XLSX.utils.book_new();

      // 2. Preparar datos de Resumen (Hoja 1)
      const summaryRows = summaryData.map(item => {
        return {
          "Vehículo": item.deviceName,
          "Tiempo Actividad (Motor ON)": formatSecondsToHMS(item.activityDuration),
          "Tiempo en Ralentí": formatSecondsToHMS(item.idleDuration),
          "% Ralentí sobre Actividad": item.idlePercentage !== null ? `${item.idlePercentage.toFixed(1)}%` : "N/A",
          "Consumo en Ralentí (Litros)": item.idleFuel !== null ? Number(item.idleFuel.toFixed(2)) : "N/A",
          "Consumo Total (Litros)": item.totalFuel !== null ? Number(item.totalFuel.toFixed(2)) : "N/A",
          "Distancia Recorrida (Km)": Number(item.distance.toFixed(2))
        };
      });

      // Crear hoja de Resumen vacía para evitar que las cabeceras se generen en la fila 1 y queden duplicadas
      const wsSummary = XLSX.utils.aoa_to_sheet([]);
      
      // Añadir fila de título a la hoja de Resumen para mejorar presentación
      XLSX.utils.sheet_add_aoa(wsSummary, [
        [`INFORME DE CONSUMO EN RALENTÍ - PERIODO: ${periodText.toUpperCase()}`],
        []
      ], { origin: "A1" });

      // Escribir los datos (con su cabecera automática) a partir de la fila 3 (A3)
      XLSX.utils.sheet_add_json(wsSummary, summaryRows, { origin: "A3" });
      
      // Auto-ajustar columnas
      autoFitColumns(wsSummary);
      
      // Añadir la hoja al libro
      XLSX.utils.book_append_sheet(wb, wsSummary, "Resumen");

      // 3. Preparar datos de Detalle (Hoja 2) si hay agrupación
      if (detailData && detailData.length > 0) {
        const detailRows = detailData.map(item => {
          let periodLabel = "Periodo";
          if (groupingType === "day") periodLabel = "Fecha / Día";
          else if (groupingType === "week") periodLabel = "Semana";
          else if (groupingType === "month") periodLabel = "Mes";

          const row = {};
          row[periodLabel] = item.period;
          row["Vehículo"] = item.deviceName;
          row["Tiempo Actividad (Motor ON)"] = formatSecondsToHMS(item.activityDuration);
          row["Tiempo en Ralentí"] = formatSecondsToHMS(item.idleDuration);
          row["% Ralentí sobre Actividad"] = item.idlePercentage !== null ? `${item.idlePercentage.toFixed(1)}%` : "N/A";
          row["Consumo en Ralentí (Litros)"] = item.idleFuel !== null ? Number(item.idleFuel.toFixed(2)) : "N/A";
          row["Consumo Total (Litros)"] = item.totalFuel !== null ? Number(item.totalFuel.toFixed(2)) : "N/A";
          row["Distancia Recorrida (Km)"] = Number(item.distance.toFixed(2));

          return row;
        });

        // Crear hoja de Detalle vacía
        const wsDetail = XLSX.utils.aoa_to_sheet([]);
        
        // Añadir título a la hoja de Detalle
        XLSX.utils.sheet_add_aoa(wsDetail, [
          [`DESGLOSE DETALLADO (${groupingType === "day" ? "DIARIO" : groupingType === "week" ? "SEMANAL" : "MENSUAL"}) - PERIODO: ${periodText.toUpperCase()}`],
          []
        ], { origin: "A1" });

        // Escribir los datos detallados a partir de la fila 3 (A3)
        XLSX.utils.sheet_add_json(wsDetail, detailRows, { origin: "A3" });

        autoFitColumns(wsDetail);
        XLSX.utils.book_append_sheet(wb, wsDetail, "Detalle Temporal");
      }

      // 4. Nombre del archivo
      const cleanPeriodText = periodText.replace(/[^a-zA-Z0-9]/g, "_");
      const filename = `Reporte_Consumo_Ralenti_${cleanPeriodText}.xlsx`;

      // 5. Descargar archivo
      XLSX.writeFile(wb, filename);
    }
  };
})();
