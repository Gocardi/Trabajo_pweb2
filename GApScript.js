function leerCorreosYCrearEventos() {
    var startDate = new Date("2024-01-01");
    var endDate = new Date("2024-12-12");
  
    var bandejaEntrada = GmailApp.search('subject:^prueba after:' + formatDate(startDate) + ' before:' + formatDate(endDate));
  
    var correosEncontrados = 0;
  
    Logger.log("Cantidad de correos encontrados: " + bandejaEntrada.length);
  
    for (var i = 0; i < bandejaEntrada.length; i++) {
      var mensajes = bandejaEntrada[i].getMessages();
      for (var j = 0; j < mensajes.length; j++) {
        var correo = mensajes[j];
        var asunto = correo.getSubject();
        var cuerpo = correo.getPlainBody();
  
        Logger.log("Asunto del correo: " + asunto);
  
        if (asunto && asunto.toLowerCase().startsWith("prueba")) {
          var titulo = extraerCampo(cuerpo, "Título del Evento:");
          var fecha = extraerCampo(cuerpo, "Fecha:");
          var hora = extraerCampo(cuerpo, "Hora:");
          var lugar = extraerCampo(cuerpo, "Lugar:");
  
          Logger.log("Título del evento: " + titulo);
          Logger.log("Fecha: " + fecha);
          Logger.log("Hora: " + hora);
          Logger.log("Lugar: " + lugar);
  
          if (titulo && fecha && hora && lugar) {
            crearEventoGoogleCalendar(titulo, fecha, hora, lugar);
            correosEncontrados++;
          }
        }
      }
    }
  
    if (correosEncontrados > 0) {
      Logger.log("Se encontraron " + correosEncontrados + " correo(s) con el asunto que comienza con 'prueba' y se crearon los eventos correspondientes.");
    } else {
      Logger.log("No se encontraron correos con el asunto que comience con 'prueba'.");
    }
  }
  
  function extraerCampo(texto, etiqueta) {
    var inicio = texto.indexOf(etiqueta);
    if (inicio !== -1) {
      inicio += etiqueta.length;
      var fin = texto.indexOf('\n', inicio);
      if (fin !== -1) {
        return texto.substring(inicio, fin).trim();
      } else {
        return texto.substring(inicio).trim();
      }
    }
    return null;
  }
  
  function crearEventoGoogleCalendar(titulo, fecha, hora, lugar) {
    try {
      var calendar = CalendarApp.getDefaultCalendar();
  
      var partesFecha = fecha.split(" ");
      var dia = partesFecha[1];
      var mes = obtenerNumeroMes(partesFecha[3]);
      var año = partesFecha[5];
      var fechaEvento = new Date(año, mes - 1, dia);
  
      var horaInicio = hora.split(" - ")[0];
      var horaFin = hora.split(" - ")[1];
  
      var horaInicioValida = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9] (a\.m\.|p\.m\.)$/i.test(horaInicio);
      var horaFinValida = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9] (a\.m\.|p\.m\.)$/i.test(horaFin);
  
      if (horaInicioValida && horaFinValida) {
        var evento = calendar.createEvent(titulo, fechaEvento, fechaEvento, {
          description: "Descripción del evento",
          location: lugar,
          startTime: fechaEvento.toISOString().slice(0, 10) + "T" + horaInicio,
          endTime: fechaEvento.toISOString().slice(0, 10) + "T" + horaFin
        });
  
        Logger.log("Evento creado: " + evento.getTitle());
      } else {
        Logger.log("Error: Formato de hora no válido.");
      }
    } catch (error) {
      Logger.log("Error al crear el evento: " + error);
    }
  }
  
  function obtenerNumeroMes(nombreMes) {
    var meses = {
      "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
      "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
    };
    return meses[nombreMes.toLowerCase()];
  }
  
  function formatDate(date) {
    var day = date.getDate();
    var month = date.getMonth() + 1;
    var year = date.getFullYear();
  
    if (day < 10) {
      day = "0" + day;
    }
    if (month < 10) {
      month = "0" + month;
    }
  
    return year + "-" + month + "-" + day;
  }
  
  leerCorreosYCrearEventos();