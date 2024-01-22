// Script para crear y personalizar certificados de participación en talleres de MetaDocencia (diciembre 2022)


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


// Función para CREAR certificado
function creaCertificado() {


  // Trae planilla de inscripciones
  var url_validate = "https://docs.google.com/spreadsheets/d/[]/"; //                  **MODIFICAR**


  var ActiveSpreadsheet = SpreadsheetApp.openByUrl(url_validate);


  var ActiveSheet = ActiveSpreadsheet.getSheets()[0];


  var data = ActiveSheet.getDataRange().getValues();




  // Crea archivos PDF de los certificados y entradas en la planilla para confirmarlos
  for(var i = 1; i < data.length; i++) {


   // Indica columna para verificar si asistió y si ya fue creado un certificado (o sea, no crea archivos duplicados)
    if ((!ActiveSheet.getRange(i+1, 36).isBlank()) && (ActiveSheet.getRange(i+1, 37).isBlank())) {


      // Encuentra nombre
      var nombre = data[i][1]; // (count-1)
      var apellido = data[i][2]; // (count-1)
      // Encuentra fecha
      //var fecha = data[i][40]; // (count-1)




      // ID de plantilla de inscripción
      let templateCertificadoId = "[]"; //                     **MODIFICAR**
      // ID de carpeta para guardar los certificados personalizados
      let carpetaCertificadosId = "[]"; //                     **MODIFICAR**


      // Crea ID de certificado persionalizado
      let templateCertis = DriveApp.getFileById(templateCertificadoId);
  
      // Crea copia con nombre de participante en la carpeta correspondiente dentro de "certificados personalizados"
      let carpetaCertis = DriveApp.getFolderById(carpetaCertificadosId);


      //Se generan las copias, se renombran los archivos 
      let asistenteId = templateCertis.makeCopy(carpetaCertis).setName( apellido +" "+nombre).getId();  
      let certiAsistente = SlidesApp.openById(asistenteId).getSlides()[0];


      // Sustituye los placeholders con nombre, apellido y fecha
      certiAsistente.replaceAllText("<<First Name>>", nombre);
      certiAsistente.replaceAllText("<<Last Name>>", apellido);
      certiAsistente.replaceAllText("<<Fecha>>", fecha);


      // Escribe en la planilla que el certificado fue creado
      ActiveSheet.getRange(i+1, 37).setValue("Sí");


      // Escribe en la planilla el ID del archivo de cada certificado
      ActiveSheet.getRange(i+1, 39).setValue(asistenteId);
  }}
}     
