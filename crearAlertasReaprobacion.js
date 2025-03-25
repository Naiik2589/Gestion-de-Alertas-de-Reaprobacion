function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Acciones') // Nombre del menú
    .addItem('Ejecutar Script', 'crearEventos') // Nombre del botón y función a ejecutar
    .addToUi();
}

function crearEventos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0); // Establecer la hora a 00:00:00 para comparar solo la fecha

  Logger.log("Iniciando proceso de creación/eliminación de eventos y envío de correos...");

  try {
    // Obtener el calendario
    const calendar = CalendarApp.getCalendarById("c_96d98b85816d17f8d577f1b8685ebfe217bd91f0cc71525ef5bbae83560685c4@group.calendar.google.com");
    if (!calendar) {
      throw new Error("No se pudo acceder al calendario. Verifica el ID y los permisos.");
    }

    // Empezar desde la fila 2 (omitir la cabecera)
    for (let i = 1; i < data.length; i++) {
      const estado = data[i][3]; // Columna D (Estado)
    // Verificar si la fila está vacía en la columna A
      if (!data[i][0]) { // Si la celda en la columna A está vacía
      Logger.log("Fila vacía encontrada en la fila " + (i + 1) + ". Deteniendo ejecución.");
      break;
      }
      Logger.log(`Procesando fila ${i + 1}: Estado = ${estado}`);

      if (estado === "Activo") {
        Logger.log(`Fila ${i + 1}: Estado es Activo. Creando eventos y verificando alertas...`);

        // Crear eventos para las alertas
        for (let j = 0; j < 5; j++) {
          const fechaAlerta = data[i][7 + j * 4]; // Columnas H, L, P, T, X
          const fechaReaprobacion = data[i][6 + j * 4]; // Columnas G, K, O, S, W
          const enCalendar = data[i][8 + j * 4]; // Columnas I, M, Q, U, Y
          const correoEnviado = data[i][9 + j * 4]; // Columnas J, N, R, V, Z
          const nombreProtocolo = data[i][0]; // Columna A
          const periodo = data[i][5]; // Columna F
          const correoPrincipal = data[i][26]; // Columna AA
          const correoBackup = data[i][27]; // Columna AB
          Logger.log(`Procesando Alerta ${j + 1} para ${nombreProtocolo}`);
          Logger.log(`Fecha Alerta: ${fechaAlerta}`);
          Logger.log(`En Calendar: ${enCalendar}`);
          Logger.log(`Correo Enviado: ${correoEnviado}`);

          // Crear evento si no existe
          if (fechaAlerta && enCalendar !== "Sí") {
            Logger.log(`Creando evento para Alerta ${j + 1}...`);

            const descripcion = `Buen día y cordial saludo\n\nEsta es una notificación de que está próxima la reaprobación del siguiente estudio:\n\n📌 ${nombreProtocolo}\n📅 Fecha Inicial ó Ultima Reaprobación: ${formatDate(data[i][4])}\n🔄 Fecha Reaprobación${j + 1}: ${formatDate(fechaReaprobacion)}\n⏳ Frecuencia de reaprobación: ${periodo}\n\nFoscal\nInspirados por la vida.`;

            // Ajustar la hora del evento (8:00 AM a 9:00 AM)
            const fechaInicio = new Date(fechaAlerta);
            fechaInicio.setHours(8, 0, 0); // 8:00 AM

            const fechaFin = new Date(fechaAlerta);
            fechaFin.setHours(9, 0, 0); // 9:00 AM

            // Crear el evento en Google Calendar
            const evento = calendar.createEvent(
              `Alerta ${j + 1} - ${nombreProtocolo}`, // Título del evento
              fechaInicio, // Fecha de inicio (8:00 AM)
              fechaFin, // Fecha de fin (9:00 AM)
              {
                description: descripcion,
                guests: `${correoPrincipal}, ${correoBackup}`,
                sendInvites: true
              }
            );
            Logger.log(`Evento creado: ${evento.getTitle()} - ${evento.getStartTime()} a ${evento.getEndTime()}`);

            // Actualizar la columna "En Calendar" a "Sí"
            sheet.getRange(i + 1, 9 + j * 4).setValue("Sí");
            Logger.log(`Columna "En Calendar${j + 1}" actualizada a "Sí"`);
          } else {
            Logger.log(`No se crea evento para Alerta ${j + 1} (ya existe o falta fecha)`);
          }

          // Enviar correo si la Fecha Alerta es hoy
          const fechaAlertaObj = new Date(fechaAlerta);
          fechaAlertaObj.setHours(0, 0, 0, 0); // Establecer la hora a 00:00:00 para comparar solo la fecha

          if (fechaAlertaObj.getTime() === hoy.getTime() && correoEnviado !== "Sí") {
            Logger.log(`Enviando correo para Alerta ${j + 1}...`);

            // Crear la descripción del correo
            const descripcionCorreo = `Buen día y cordial saludo\n\nEsta es una notificación de que está próxima la reaprobación del siguiente estudio:\n\n📌 ${nombreProtocolo}\n📅 Fecha Inicial ó Ultima Reaprobación: ${formatDate(data[i][4])}\n🔄 Fecha Reaprobación${j + 1}: ${formatDate(fechaReaprobacion)}\n⏳ Frecuencia de reaprobación: ${periodo}\n\nFoscal\nInspirados por la vida.`;

            // Enviar el correo
            MailApp.sendEmail({
              to: `${correoPrincipal}, ${correoBackup}`,
              subject: `Alerta ${j + 1} - ${nombreProtocolo}`,
              body: descripcionCorreo
            });

            Logger.log(`Correo enviado a ${correoPrincipal} y ${correoBackup}`);

            // Actualizar la columna "Correo Enviado" a "Sí"
            sheet.getRange(i + 1, 10 + j * 4).setValue("Sí");
            Logger.log(`Columna "Correo${j + 1} Enviado" actualizada a "Sí"`);
          } else {
            Logger.log(`No se envía correo para Alerta ${j + 1} (no es hoy o ya se envió)`);
          }
        }
      } else if (estado === "Cerrado") {
        Logger.log(`Fila ${i + 1}: Estado es Cerrado. Eliminando eventos futuros y actualizando columnas...`);

        // Borrar eventos futuros y actualizar columnas
        for (let j = 0; j < 5; j++) {
          const enCalendar = data[i][8 + j * 4]; // Columnas I, M, Q, U, Y
          const correoEnviado = data[i][9 + j * 4]; // Columnas J, N, R, V, Z

          if (enCalendar === "Sí") {
            // Buscar y borrar el evento correspondiente
            const fechaAlerta = data[i][7 + j * 4]; // Columnas H, L, P, T, X
            Logger.log(`Buscando evento para Alerta ${j + 1} en fecha: ${fechaAlerta}`);

            // Buscar eventos en un rango de un día completo
            const fechaInicioBusqueda = new Date(fechaAlerta);
            fechaInicioBusqueda.setHours(0, 0, 0, 0); // Inicio del día
            const fechaFinBusqueda = new Date(fechaAlerta);
            fechaFinBusqueda.setHours(23, 59, 59, 999); // Fin del día
            const eventos = calendar.getEvents(fechaInicioBusqueda, fechaFinBusqueda);
            eventos.forEach(evento => {
              if (evento.getTitle().includes(`Alerta ${j + 1} - ${data[i][0]}`)) {
                Logger.log(`Evento encontrado: ${evento.getTitle()} - ${evento.getStartTime()}`);
                evento.deleteEvent();
                Logger.log(`Evento eliminado: ${evento.getTitle()}`);
              }
            });

            // Actualizar la columna "En Calendar" a "No"
            sheet.getRange(i + 1, 9 + j * 4).setValue("No");
            Logger.log(`Columna "En Calendar${j + 1}" actualizada a "No"`);
          }

          if (correoEnviado !== "No") {
            // Actualizar la columna "Correo Enviado" a "No"
            sheet.getRange(i + 1, 10 + j * 4).setValue("No");
            Logger.log(`Columna "Correo${j + 1} Enviado" actualizada a "No"`);
          }
        }
      }
    }
    Logger.log("Proceso completado.");

    // Mostrar una alerta al usuario solo si se ejecuta manualmente
    if (typeof ScriptApp !== 'undefined') {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Proceso completado", "El script ha finalizado correctamente.", ui.ButtonSet.OK);
    }
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    if (typeof ScriptApp !== 'undefined') {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Error", `Ocurrió un error: ${error.message}`, ui.ButtonSet.OK);
    }
  }
}

// Función para formatear fechas en formato "21 de Enero de 2025"
function formatDate(date) {
  const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const d = new Date(date);
  const dia = d.getDate();
  const mes = meses[d.getMonth()]; // Obtiene el nombre del mes
  const año = d.getFullYear();
  return `${dia} de ${mes} de ${año}`;
}
