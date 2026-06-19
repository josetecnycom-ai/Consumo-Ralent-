/* apiService.js - Servicio para interactuar con la API de Geotab o simular datos en local */

const GeotabApiService = (function () {
  let apiInstance = null;
  let stateInstance = null;
  let isOffline = false;

  // Diagnósticos conocidos de Geotab
  const DIAGNOSTICS = {
    totalFuel: "DiagnosticDeviceTotalFuelId",       // Consumo total acumulado (L)
    totalIdleFuel: "DiagnosticDeviceTotalIdleFuelId", // Consumo en ralentí acumulado (L)
    odometer: "DiagnosticOdometerId"                // Odómetro acumulado (m)
  };

  // Datos simulados (Mock Data)
  const mockDevices = [
    { id: "b1", name: "Camión Volvo FH16 - M-8812-ZZ", groups: [{ name: "Logística" }] },
    { id: "b2", name: "Furgoneta Reparto - M-2245-AA", groups: [{ name: "Distribución" }] },
    { id: "b3", name: "Scania R450 Longline - M-9031-BB", groups: [{ name: "Ruta Nacional" }] },
    { id: "b4", name: "Pickup Toyota Hilux - M-1188-CC", groups: [{ name: "Supervisión" }] }, // Sin datos de combustible
    { id: "b5", name: "Furgón Eléctrico Maxus - M-5566-EE", groups: [{ name: "Última Milla" }] }
  ];

  // Generador de datos simulados para un periodo de fechas
  function generateMockStatusData(devices, fromDate, toDate) {
    const trips = [];
    const statusData = [];

    const start = new Date(fromDate);
    const end = new Date(toDate);
    const dayCount = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) || 1;

    devices.forEach(device => {
      // Valores acumulados de inicio para este vehículo
      let currentOdometer = 100000000 + Math.random() * 50000000; // metros (100k - 150k km)
      let currentTotalFuel = 25000 + Math.random() * 15000;       // litros
      let currentIdleFuel = 3000 + Math.random() * 2000;          // litros

      // El dispositivo b4 no reporta combustible
      const reportsFuel = device.id !== "b4";

      // Insertar lecturas iniciales en la fecha de inicio
      statusData.push({
        device: { id: device.id },
        dateTime: start.toISOString(),
        diagnostic: { id: DIAGNOSTICS.odometer },
        data: currentOdometer
      });

      if (reportsFuel) {
        statusData.push({
          device: { id: device.id },
          dateTime: start.toISOString(),
          diagnostic: { id: DIAGNOSTICS.totalFuel },
          data: currentTotalFuel
        });
        statusData.push({
          device: { id: device.id },
          dateTime: start.toISOString(),
          diagnostic: { id: DIAGNOSTICS.totalIdleFuel },
          data: currentIdleFuel
        });
      }

      // Generar viajes y lecturas día por día
      for (let d = 0; d < dayCount; d++) {
        const currentDate = new Date(start);
        currentDate.setDate(start.getDate() + d);

        // Probabilidad de viaje en el día (80%)
        if (Math.random() > 0.2) {
          // Generar entre 1 y 3 viajes por día
          const tripCount = Math.floor(Math.random() * 3) + 1;
          for (let t = 0; t < tripCount; t++) {
            const tripStart = new Date(currentDate);
            tripStart.setHours(8 + t * 4, Math.floor(Math.random() * 60), 0);
            
            // Duración del viaje: 30 a 120 minutos
            const durationMin = 30 + Math.floor(Math.random() * 90);
            const tripEnd = new Date(tripStart);
            tripEnd.setMinutes(tripStart.getMinutes() + durationMin);

            // Tiempo en ralentí del viaje:
            // Volvo (b1) y Furgoneta (b2) tienen alto ralentí. Scania (b3) y Maxus (b5) bajo.
            let idlePct = 0.1; // 10% por defecto
            if (device.id === "b1") idlePct = 0.35 + Math.random() * 0.15; // 35-50%
            else if (device.id === "b2") idlePct = 0.20 + Math.random() * 0.15; // 20-35%
            else if (device.id === "b3") idlePct = 0.05 + Math.random() * 0.08; // 5-13%
            else if (device.id === "b5") idlePct = 0.02 + Math.random() * 0.05; // 2-7%
            else if (device.id === "b4") idlePct = 0.15 + Math.random() * 0.10; // 15-25%

            const idleDurationMin = durationMin * idlePct;
            const drivingDurationMin = durationMin - idleDurationMin;

            // Distancia en metros (velocidad promedio 40-75 km/h durante conducción)
            const avgSpeedKmh = 40 + Math.random() * 35;
            const tripDistance = (drivingDurationMin / 60) * avgSpeedKmh; // en Km

            // Actualizar acumuladores
            currentOdometer += tripDistance * 1000; // en metros

            // Consumo de combustible (Volvo y Scania consumen más; Maxus eléctrico no consume o consume kWh simulado)
            let fuelRateLh = 25; // 25 litros por hora de conducción promedio
            let idleRateLh = 2.2; // 2.2 litros por hora en ralentí
            if (device.id === "b1") { fuelRateLh = 32; idleRateLh = 2.8; }
            else if (device.id === "b3") { fuelRateLh = 28; idleRateLh = 2.4; }
            else if (device.id === "b2") { fuelRateLh = 12; idleRateLh = 1.2; }
            else if (device.id === "b5") { fuelRateLh = 8; idleRateLh = 0.5; } // Simulación de consumo bajo / híbrido

            const drivingFuel = (drivingDurationMin / 60) * fuelRateLh;
            const idleFuel = (idleDurationMin / 60) * idleRateLh;

            if (reportsFuel) {
              currentTotalFuel += (drivingFuel + idleFuel);
              currentIdleFuel += idleFuel;
            }

            // Registrar viaje
            trips.push({
              device: { id: device.id },
              start: tripStart.toISOString(),
              stop: tripEnd.toISOString(),
              distance: tripDistance, // en Km
              drivingDuration: formatTimeSpan(drivingDurationMin * 60),
              idlingDuration: formatTimeSpan(idleDurationMin * 60)
            });

            // Registrar lecturas de diagnóstico al final del viaje
            statusData.push({
              device: { id: device.id },
              dateTime: tripEnd.toISOString(),
              diagnostic: { id: DIAGNOSTICS.odometer },
              data: currentOdometer
            });

            if (reportsFuel) {
              statusData.push({
                device: { id: device.id },
                dateTime: tripEnd.toISOString(),
                diagnostic: { id: DIAGNOSTICS.totalFuel },
                data: currentTotalFuel
              });
              statusData.push({
                device: { id: device.id },
                dateTime: tripEnd.toISOString(),
                diagnostic: { id: DIAGNOSTICS.totalIdleFuel },
                data: currentIdleFuel
              });
            }
          }
        }
      }
    });

    return { trips, statusData };
  }

  // Formatear segundos a string "HH:MM:SS" para simular TimeSpan
  function formatTimeSpan(totalSeconds) {
    const hrs = Math.floor(totalSeconds / 3600);
    const mins = Math.floor((totalSeconds % 3600) / 60);
    const secs = Math.floor(totalSeconds % 60);
    return `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  }

  return {
    init: function (api, state) {
      if (api && state) {
        apiInstance = api;
        stateInstance = state;
        isOffline = false;
        console.log("GeotabApiService inicializado en modo ONLINE");
      } else {
        isOffline = true;
        console.log("GeotabApiService inicializado en modo SIMULACIÓN (offline)");
      }
    },

    getIsOffline: function () {
      return isOffline;
    },

    getDevices: function (callback) {
      if (isOffline) {
        setTimeout(() => callback(mockDevices), 100);
        return;
      }

      // Consulta real a Geotab
      apiInstance.call("Get", {
        typeName: "Device",
        search: {
          // Se puede añadir filtro si es necesario
        }
      }, function (result) {
        callback(result);
      }, function (error) {
        console.error("Error al obtener vehículos de Geotab:", error);
        callback([]);
      });
    },

    // Obtener los datos consolidados (viajes y combustible) del periodo
    getData: function (deviceIds, fromDate, toDate, callback) {
      if (isOffline) {
        setTimeout(() => {
          const mockData = generateMockStatusData(mockDevices, fromDate, toDate);
          
          // Filtrar por vehículos seleccionados
          const filteredTrips = mockData.trips.filter(t => deviceIds.includes(t.device.id));
          const filteredStatusData = mockData.statusData.filter(s => deviceIds.includes(s.device.id));
          
          callback({
            trips: filteredTrips,
            statusData: filteredStatusData
          });
        }, 300);
        return;
      }

      // Modo Online - Construir un MultiCall para optimizar red
      const calls = [
        ["Get", {
          typeName: "Trip",
          search: {
            fromDate: fromDate,
            toDate: toDate
          }
        }],
        ["Get", {
          typeName: "StatusData",
          search: {
            diagnosticSearch: { id: DIAGNOSTICS.totalFuel },
            fromDate: fromDate,
            toDate: toDate
          }
        }],
        ["Get", {
          typeName: "StatusData",
          search: {
            diagnosticSearch: { id: DIAGNOSTICS.totalIdleFuel },
            fromDate: fromDate,
            toDate: toDate
          }
        }],
        ["Get", {
          typeName: "StatusData",
          search: {
            diagnosticSearch: { id: DIAGNOSTICS.odometer },
            fromDate: fromDate,
            toDate: toDate
          }
        }]
      ];

      apiInstance.multiCall(calls, function (results) {
        const trips = results[0] || [];
        const fuelData = results[1] || [];
        const idleFuelData = results[2] || [];
        const odometerData = results[3] || [];

        // Combinar todos los StatusData en un solo arreglo
        const statusData = [...fuelData, ...idleFuelData, ...odometerData];

        // Filtrar en memoria por los vehículos seleccionados para no saturar al servidor
        const filteredTrips = trips.filter(t => t.device && deviceIds.includes(t.device.id));
        const filteredStatusData = statusData.filter(s => s.device && deviceIds.includes(s.device.id));

        callback({
          trips: filteredTrips,
          statusData: filteredStatusData
        });
      }, function (error) {
        console.error("Error al obtener datos combinados por MultiCall:", error);
        callback({ trips: [], statusData: [] });
      });
    },

    DIAGNOSTICS: DIAGNOSTICS
  };
})();
