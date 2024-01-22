// Script para enviar recordatorios de Google Calendar a personas inscriptas en el taller de [MODIFICAR: nombre del taller] e invitarlas a un evento en Google Calendar


// Documentación: https://developers.google.com/apps-script/reference/spreadsheet/sheet,
// https://developers.google.com/apps-script/reference/mail,
// https://developers.google.com/apps-script/reference/utilities,




// Función general para ejecutar todo el script
function enviaMailRecordatorio() {




// ENVIAR MAIL A INSCRIPTES EN EL TALLER


// Trae correos de la spreadsheet de inscripciones
  var url_validate = "MODIFICAR: url de planilla de inscripción"; //                  **MODIFICAR**


  var ActiveSpreadsheet = SpreadsheetApp.openByUrl(url_validate);


  var ActiveSheet = ActiveSpreadsheet.getSheets()[0];


  var data = ActiveSheet.getDataRange().getValues();




  // Crea envío de mail y entrada en la planilla para confirmarlo
  for(var i = 1; i < data.length; i++) {


  // Indica columna para verificar si ya fue enviado un mail - o sea, no envía mails duplicados
    if (ActiveSheet.getRange(i+1, 35).isBlank()) {


      // Encuentra nombre y mail
      var nombre = data[i][1];
      var mail = data[i][4];




      // Mensaje que va en el mail a la persona inscripta
      // Para usar emojis en html, ver: https://www.w3schools.com/html/html_emojis.asp, https://www.w3schools.com/charsets/ref_emoji.asp, https://emojipedia.org/


      var message =
        "<p style='text-align:center'><img style='max-width: 250px; height: auto;' src='https://www.metadocencia.org/img/MD_original.png' alt='Logo de MetaDocencia'></p>" +
        "<p>Te recordamos que el taller <b>[MODIFICAR: nombre del taller. Ejemplo: piloto de MetaEvaluaciones: enseñar evaluando]</b> se desarrollará mañana, <b>[MODIFICAR: fecha. Ejemplo: miércoles 7 de diciembre de 17 a 19 (Hora Argentina)].</b></p>" +
        "<p><b>&#128279; <a href='[MODIFICAR: Link a evento de Zoom]'>Podrás sumarte a través de este link</a></b></p>" +
        "<p>Te invitamos además a sumarte al canal [MODIFICAR: nombre del canal. Ejemplo: #taller-metaevaluaciones] en <a href='https://w3id.org/metadocencia/slack'>nuestro Slack</a>, para conectar con otras personas y docentes y compartir materiales, ideas y recursos sobre la propuesta &#129299;&#128171;.</p>" +
        "<p>Todos nuestros espacios siguen <a href='https://www.metadocencia.org/pdc/'>estas pautas de convivencia</a>. Tu participación en este evento asume que aceptás seguir estas pautas. Cualquier consulta, escribinos a cursos@metadocencia.org</p>" +
        "<p>¡Te esperamos!</p>" +
        "<p><br></p>" +
        "<p style='text-align:center; color:white; background-color:#0b5394;'><br><a href='https://www.metadocencia.org/' style='color:white'>www.metadocencia.org</a><br><br><a href='https://www.linkedin.com/company/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/linkedin-brands-white.png' target='_blank' alt='LinkedIn'></a>&nbsp;<a href='https://www.facebook.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/facebook-square-brands-white.png' target='_blank' alt='Facebook'></a>&nbsp;<a href='https://www.twitter.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/twitter-square-brands-white.png' target='_blank'  alt='Twitter'></a>&nbsp;<a href='https://www.youtube.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/youtube-brands-white.png' target='_blank'  alt='YouTube'></a><br><br></p>";




      // Crea asunto del mail
      var Subject = "MODIFICAR: Recordatorio: [fecha y nombre del curso. Ejemplo: 7 de diciembre 17h (UTC-3): Taller MetaEvaluaciones: enseñar evaluando]"; //                  **MODIFICAR**


      var SendTo = mail;




      // Arma el mail a ser enviado
      MailApp.sendEmail({
        to: SendTo,
        from: alias[0],
        name: "MetaDocencia",
        subject: Subject,
        htmlBody: message,
      });


      // Escribe en la planilla que el recordatorio fue enviado
      ActiveSheet.getRange(i+1, 35).setValue("Sí");


    }


  }
}
