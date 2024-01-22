// Script para enviar invitaciones de Google Calendar a personas inscriptas en el taller de [MODIFICAR: nombre del taller] e invitarlas a un evento en Google Calendar


// Documentación: https://developers.google.com/apps-script/reference/spreadsheet/sheet,
// https://developers.google.com/apps-script/reference/mail,
// https://developers.google.com/apps-script/reference/utilities,
// https://developers.google.com/apps-script/reference/calendar




// Función general para ejecutar todo el script
function enviaMailYAgregaCalendario() {




// ENVIAR MAIL A INSCRIPTES EN EL TALLER


// Trae correos de la spreadsheet de inscripciones
  var url_validate = "[MODIFICAR: url de planilla de inscripción]";


  var ActiveSpreadsheet = SpreadsheetApp.openByUrl(url_validate);


  var ActiveSheet = ActiveSpreadsheet.getSheets()[0];


  var data = ActiveSheet.getDataRange().getValues();




  // Crea envío de mail y entrada en la planilla para confirmarlo
  for(var i = 1; i < data.length; i++) {


  // Indica columna para verificar si ya fue enviado un mail - o sea, no envía mails de confirmación duplicados
    if (ActiveSheet.getRange(i+1, 32).isBlank()) {


      // Encuentra nombre y mail
      var nombre = data[i][1];
      var mail = data[i][4];




      // Mensaje que va en el mail a la persona inscripta
      // Para usar emojis en html, ver: https://www.w3schools.com/html/html_emojis.asp, https://www.w3schools.com/charsets/ref_emoji.asp, https://emojipedia.org/


      var message =
        "<p style='text-align:center'><img style='max-width: 250px; height: auto;' src='https://www.metadocencia.org/img/MD_original.png' alt='Logo de MetaDocencia'></p>" +
        "<p>Gracias por tu interés en participar en los cursos de MetaDocencia &#129303;.</p>" +
        "<p>El <b>taller [MODIFICAR: nombre del taller. Ejemplo: piloto de MetaEvaluaciones: enseñar evaluando]</b> se llevará a cabo el <b>próximo [fecha. Ejemplo: miércoles 7 de diciembre de 17 a 19 (Hora Argentina)].</b></p>" +
        "<p>Te agregamos a un evento en Google Calendar en la fecha y hora correspondientes al curso. El enlace a Zoom se enviará en un mail de recordatorio un día antes del taller.</p>" +
        "<p>Todos nuestros espacios siguen <a href='https://www.metadocencia.org/cdc/'>estas pautas de convivencia</a>. Tu participación en este evento asume que aceptás seguir estas pautas. Cualquier consulta, escribinos a cursos@metadocencia.org</p>" +
        "<p>¡Te esperamos!</p>" +
        "<p><br></p>" +
        "<p style='text-align:center; color:white; background-color:#0b5394;'><br><a href='https://www.metadocencia.org/' style='color:white'>www.metadocencia.org</a><br><br><a href='https://www.linkedin.com/company/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/linkedin-brands-white.png' target='_blank' alt='LinkedIn'></a>&nbsp;<a href='https://www.facebook.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/facebook-square-brands-white.png' target='_blank' alt='Facebook'></a>&nbsp;<a href='https://floss.social/@MetaDocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/mastodon-logo.PNG' target='_blank' alt='Mastodon'></a>&nbsp;<a href='https://www.twitter.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/twitter-square-brands-white.png' target='_blank'  alt='Twitter'></a>&nbsp;<a href='https://www.youtube.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/youtube-brands-white.png' target='_blank'  alt='YouTube'></a><br><br></p>";




      // Crea asunto del mail
      var Subject = "Te inscribiste al taller [MODIFICAR: nombre del taller. Ejemplo: piloto de MetaEvaluaciones: enseñar evaluando]";


      var SendTo = mail;

      // Arma el mail a ser enviado
      MailApp.sendEmail({
        to: SendTo,
        from: alias[0],
        name: "MetaDocencia",
        subject: Subject,
        htmlBody: message,
      });


      // Escribe en la planilla que el mail fue enviado
      ActiveSheet.getRange(i+1, 32).setValue("Sí");


    }


  }




// AGREGAR INVITADES AL EVENTO DEL TALLER




  // Encuentra mail


  for(var i = 1; i < data.length; i++) {


    // Indica columna para verificar si la persona ya está invitada
    if (ActiveSheet.getRange(i+1, 33).isBlank()) {


      var mail = data[i][4];


      let attendeeEmail = mail; // Mail de la persona que se quiere invitar




      // OBSERVACIONES SOBRE COMO ENCONTRAR LOS IDS DEL CALENDARIO Y DEL EVENTO


      // A continuación necesitamos los IDs del calendario usado y del evento del taller. Para eso, debes ir al calendario Cursos MetaDocencia > clicar en el evento correspondiente > editar (ícono de lápiz) > Más acciones > Publicar evento > Copiar el enlace de abajo, sin HTML. 
      // El calendarId es lo que viene después de "&tmsrc=", hasta el final del enlace.
      // El eventoIdOriginal es lo que está entre "&tmeid=" y "&tmsrc=".


      let calendarId = "[MODIFICAR: ID del calendario]"; // ID del calendario que contiene el evento (Cursos MetaDocencia)


      let eventoIdOriginal = "[MODIFICAR: ID del evento]";


      // El ID arriba está encoded en base 64. Abajo, lo decodamos:


      var eventodecoded = Utilities.base64Decode(eventoIdOriginal, Utilities.Charset.UTF_8);


      var eventId = Utilities.newBlob(eventodecoded).getDataAsString().split(" ")[0];




      // Verifica si el calendario y el evento fueron encontrados por el script


      let calendar = CalendarApp.getCalendarById(calendarId);
      if (calendar === null) {
        console.log("Calendario no encontrado", calendarId);
        return;
      }


      let event = calendar.getEventById(eventId);
      if (event === null) {
        console.log("Evento no encontrado", eventId);
        return;
      }


      // Agrega invitade al evento en Google Calendar


      event.addGuest(attendeeEmail);




      // Escribe en la planilla que el la persona fue invitada al evento en Calendar


      ActiveSheet.getRange(i+1, 33).setValue("Sí");


    }


  }




// ENVIAR MAIL SI HAY REQUERIMIENTO DE ACCESIBILIDAD


// Crea envío de mail y entrada en la planilla para confirmarlo


  for(var i = 1; i < data.length; i++) {


    // Indica columnas para verificar si hay respuesta de accesibilidad
    if (!ActiveSheet.getRange(i+1, 8).isBlank() && ActiveSheet.getRange(i+1, 34).isBlank()) {


      // Crea lista de preguntas y respuestas desde la planilla al mail que se envía a Accesibilidad
      var datos_string = "<ul>";
      for (col = 1; col < 10; col++) {
        var pregunta = data[0][col];
        var respuesta = data[i][col];
        if (col == 7) {
          datos_string += "<li>" + pregunta + ": <b>" + respuesta + "</b></li>";
        } else {
          datos_string += "<li>" + pregunta + ": " + respuesta + "</li>";
        }
      }
      datos_string += "</ul>";


      // Mensaje que va en el mail a Accesibilidad
      var message =
        "<p>Un nuevo registro en el Taller [MODIFICAR: nombre del taller. Ejemplo: MetaEvaluaciones] respondió cosas de accesibilidad.</p>" +
        "<p>" + data[i][1] + " " + data[i][2] + " &lt;" + data[i][4] + "&gt; respondió:</p>" +
        datos_string;


      // Crea asunto del mail
      var Subject = "[MODIFICAR: Taller [nombre del taller. Ejemplo: MetaEvaluaciones]] Requerimiento de accesibilidad de " + data[i][1] + " " + data[i][2];


      // Arma el mail a ser enviado
      MailApp.sendEmail({
          to: "accesibilidad@metadocencia.org",
          cc: "info@metadocencia.org",
          subject: Subject,
          htmlBody: message,
      });


      // Escribe en la planilla que el mail fue enviado al equipo accesibilidad
      ActiveSheet.getRange(i+1, 34).setValue("Sí");
    }
  }


}
