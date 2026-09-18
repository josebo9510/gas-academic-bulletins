function testAcceso() {

  // 1. CONEXIÓN Y CONFIGURACIÓN

  var idPlantilla =  'AQUI_DEBES_PONER_EL_ID_DE_TU_PLANTILLA';
  var nombreHoja = "AQUI_DEBES_PONER_EL_NOMBRE_DE_TU_HOJA";

   var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja) {

    throw new Error("Fallo crítico: No existe ninguna pestaña que se llame exactamente '" + nombreHoja + "'. Revisa el nombre abajo en tu Excel.");

  }

  var datos = hoja.getDataRange().getValues();
  var archivoPlantilla = DriveApp.getFileById(idPlantilla);
  var carpetaDestino = archivoPlantilla.getParents().next();

  // 2. CREACIÓN DEL DOCUMENTO MAESTRO (CORRECCIÓN APLICADA)

  // Copiamos la plantilla original para heredar el tamaño Oficio y el Encabezado.

  var archivoFinalDrive = archivoPlantilla.makeCopy("Boletines Unificados Bimestre 1", carpetaDestino);
  var docFinal = DocumentApp.openById(archivoFinalDrive.getId());
  var cuerpoFinal = docFinal.getBody();

 
  // Vaciamos el cuerpo del documento. Esto borra las etiquetas, pero MANTIENE intacto el tamaño de la hoja y el encabezado.

  cuerpoFinal.clear();

  // 3. PROCESAMIENTO DE FILAS

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];

    var nombreOriginal = fila[1];

    if (!nombreOriginal) continue;

       var nombreMayusculas = String(nombreOriginal).toUpperCase();

       // Crear documento temporal para este estudiante

    var docTemp = archivoPlantilla.makeCopy("temp_" + nombreMayusculas, carpetaDestino);

    var bodyTemp = DocumentApp.openById(docTemp.getId()).getBody();
    // 4. REEMPLAZO DE ETIQUETAS

    bodyTemp.replaceText("{{NOMBRE_ESTUDIANTE}}", nombreMayusculas);
    bodyTemp.replaceText("{{NOTA_CIENCIAS}}", formatearNota(fila[2]));
    bodyTemp.replaceText("{{DES_CIENCIAS}}", formatearDesempeno(fila[3]));

    bodyTemp.replaceText("{{NOTA_SOCIALES}}", formatearNota(fila[4]));

    bodyTemp.replaceText("{{DES_SOCIALE}}", formatearDesempeno(fila[5]));
    bodyTemp.replaceText("{{NOTA_ESPAÑOL}}", formatearNota(fila[6]));

    bodyTemp.replaceText("{{DES_ESPAÑ}}", formatearDesempeno(fila[7]));
    bodyTemp.replaceText("{{NOTA_INGLES}}", formatearNota(fila[8]));

    bodyTemp.replaceText("{{DES_INGLES}}", formatearDesempeno(fila[9]));
    bodyTemp.replaceText("{{NOTA_ORTOGRA}}", formatearNota(fila[10]));

    bodyTemp.replaceText("{{DES_ORTOGR}}", formatearDesempeno(fila[11]));
    bodyTemp.replaceText("{{NOTA_MATE}}", formatearNota(fila[12]));

    bodyTemp.replaceText("{{DES_MATE}}", formatearDesempeno(fila[13]));
    bodyTemp.replaceText("{{NOTA_GEOMET}}", formatearNota(fila[14]));

    bodyTemp.replaceText("{{DES_GEOM}}", formatearDesempeno(fila[15]));
    bodyTemp.replaceText("{{NOTA_SISTEMAS}}", formatearNota(fila[16]));

    bodyTemp.replaceText("{{DES_SISTEMA}}", formatearDesempeno(fila[17]));
    bodyTemp.replaceText("{{NOTA_FINANCIE}}", formatearNota(fila[18]));

    bodyTemp.replaceText("{{DES_FINANCI}}", formatearDesempeno(fila[19]));
    bodyTemp.replaceText("{{NOTA_PROYEC}}", formatearNota(fila[20]));

    bodyTemp.replaceText("{{DES_PROYEC}}", formatearDesempeno(fila[21]));

   bodyTemp.replaceText("{{NOTA_EDU_FISI}}", formatearNota(fila[22]));

    bodyTemp.replaceText("{{DES_EDU_FIS}}", formatearDesempeno(fila[23]));

   var desempenoFinal = String(fila[25]).toUpperCase();

    bodyTemp.replaceText("{{PROM_PF}}", desempenoFinal);

    bodyTemp.replaceText("{{NOT_PF}}", formatearNota(fila[24]));

    bodyTemp.replaceText("{{PT}}", formatearNota(fila[27]));

    // 5. COPIAR CONTENIDO AL DOCUMENTO MAESTRO

    for (var j = 0; j < bodyTemp.getNumChildren(); j++) {

      var elemento = bodyTemp.getChild(j).copy();
      var tipo = elemento.getType();

           if (tipo == DocumentApp.ElementType.PARAGRAPH) {

        cuerpoFinal.appendParagraph(elemento);

      } else if (tipo == DocumentApp.ElementType.TABLE) {

        cuerpoFinal.appendTable(elemento);

      } else if (tipo == DocumentApp.ElementType.LIST_ITEM) {

        cuerpoFinal.appendListItem(elemento);

      }

    }

    // Agregar salto de página excepto para el último estudiante

    if (i < datos.length - 1) {

      cuerpoFinal.appendPageBreak();

    }
    // Borrar documento temporal

    docTemp.setTrashed(true);

  }

  // Eliminar el primer párrafo vacío que queda tras el clear()

  cuerpoFinal.removeChild(cuerpoFinal.getChild(0));

  docFinal.saveAndClose();

  console.log("¡Proceso terminado con éxito!");}
function formatearNota(nota) {

  if (isNaN(nota) || nota === "") return nota;

  return Number(nota).toFixed(1);

}

function formatearDesempeno(letra) {

  letra = String(letra).trim().toUpperCase();

  switch (letra) {
    case "B": return "BAJO";
    case "BS": return "BÁSICO";
    case "A": return "ALTO";
    case "S": return "SUPERIOR";
    default: return letra;

  }

} 
