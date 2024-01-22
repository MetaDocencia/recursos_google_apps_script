// Script para enviar certificados de participación en talleres de MetaDocencia (diciembre 2022)


// DOCUMENTACION INTERNA:
// 


// Documentación Google: 
// https://developers.google.com/apps-script/reference/spreadsheet/sheet,
// https://developers.google.com/apps-script/reference/mail,
// https://developers.google.com/apps-script/reference/utilities,


// Referencia original:
// https://developers.google.com/apps-script/samples/automations/employee-certificate


/*
Copyright 2022 Google LLC


Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at


    https://www.apache.org/licenses/LICENSE-2.0


Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/


// Función para ENVIAR certificado
function enviaCertis() {


  // Trae planilla de inscripcioness.
  var url_validate = "https://docs.google.com/spreadsheets/d/[]/"; //                **MODIFICAR**


  var ActiveSpreadsheet = SpreadsheetApp.openByUrl(url_validate);


  var ActiveSheet = ActiveSpreadsheet.getSheets()[0];


  var data = ActiveSheet.getDataRange().getValues();


  var slideId = DriveApp.getFileById;
  
  for(var i = 1; i < data.length; i++) {


   // Indica columna para verificar si ya fue enviado un certificado
    if (ActiveSheet.getRange(i+1, 38).isBlank()) {


           // Encuentra mail
      //var nombre = data[i][1];
      //var apellido = data[i][2];
      var mail = data[i][4];
      var idasistente = data[i][38]
  


      let carpetaCertificadosId = "[]"; // Carpeta para guardar los certificados IV

  
      // Make a copy of the Slide template and rename it with employee name
      let carpetaCertis = DriveApp.getFolderById(carpetaCertificadosId);
      //let asistenteId = templateCertis.makeCopy(carpetaCertis).setName(apellido +" "+nombre).getId();        
      let certiAsistente = SlidesApp.openById(idasistente).getSlides()[0];




      // Carga archivo del certificado personalizado de cada asistente
      let archivoAdjunto = DriveApp.getFileById(idasistente);


     
      var message =
      "<p style='text-align:center'><img style='max-width: 250px; height: auto;' src='https://www.metadocencia.org/img/MD_original.png' alt='Logo de MetaDocencia'></p>" +
      "<p>Enviamos adjunto tu certificado de participación en el taller <b>[]</b>.</p>" +
"<p>Podés encontrar el Documento compartido utilizado durante el taller en el siguiente enlace: <a href='[]'>Documento compartido</a></p>" +
      "<p>Te invitamos además a sumarte al canal #taller-zoom-accesible en <a href='https://w3id.org/metadocencia/slack'>nuestro Slack</a>, para conectar con otras personas y docentes y compartir materiales, ideas y recursos sobre la propuesta &#129299;&#128171;.</p>" +
      "<p>¡Muchas gracias!</p>" +
      "<p><br></p>" +
      "<p style='text-align:center; color:white; background-color:#0b5394;'><br><a href='https://www.metadocencia.org/' style='color:white'>www.metadocencia.org</a><br><br><a href='https://www.linkedin.com/company/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/linkedin-brands-white.png' target='_blank' alt='LinkedIn'></a>&nbsp;<a href='https://www.facebook.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/facebook-square-brands-white.png' target='_blank' alt='Facebook'></a>&nbsp;<a href='https://floss.social/@MetaDocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/mastodon-logo.PNG' target='_blank' alt='Mastodon'></a>&nbsp;<a href='https://www.twitter.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/twitter-square-brands-white.png' target='_blank'  alt='Twitter'></a>&nbsp;<a href='https://www.youtube.com/metadocencia'><img style='max-width: 2.0em; height: auto;' src='https://www.metadocencia.org/img/youtube-brands-white.png' target='_blank'  alt='YouTube'></a><br><br></p>";


      // Crea asunto del mail
      var Subject = "Certificado: Taller Zoom accesible con lector de pantalla de MetaDocencia";


      var SendTo = mail;




      // Arma el mail a ser enviado
      MailApp.sendEmail({
        to: SendTo,
        from: alias[0],
        name: "MetaDocencia",
        subject: Subject,
        htmlBody: message,
        attachments: [archivoAdjunto.getAs(MimeType.PDF)],
      });


      // Escribe en la planilla que el certificado fue enviado
      ActiveSheet.getRange(i+1, 38).setValue("Sí");
    }
  }
}

