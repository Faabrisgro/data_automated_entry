function actualizarFechaVentas() {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("movimientos");
    var data = hoja.getDataRange().getValues();
  
    for (var i = 1; i < data.length; i++) {
      if (data[i][4] === "Ventas" && data[i][19] == "") { // Columna E (4) y Columna T (19). Si es igual a ventas y la columna fecha está vacía
        hoja.getRange(i + 1, 20).setValue(new Date()); // Columna T (índice 20)
      }
    }
  }
  
  
  var ss= SpreadsheetApp.getActiveSpreadsheet()
  
  
  // Función para ejecuciones de administradores
  
  function esAdministrador() {
    var userEmail = Session.getEffectiveUser().getEmail();
    var administrators = ["francomoreno6991@gmail.com"]; // Reemplaza con los correos electrónicos de los administradores
    return administrators.includes(userEmail);
  
  }
  
  function onEdit(e) {
    var range = e.range;
    var sheet = range.getSheet();
    var cell = sheet.getRange("B66");
    var cell2 = sheet.getRange("C23");
  
    if (range.getA1Notation() == "B66" && cell.getValue() == "") {
      cell.setValue("Celda vacía, agregar Estado del equipo");
    }
      if (range.getA1Notation() == "C23" && cell2.getValue() == "") {
      cell2.setValue("Recuerde agregar Estado del Equipo antes de realizar una búsqueda");
    }
  }
  
  function actualizarFecha(sheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var dataRange;
    var dateRange;
    var dateValues;
  
    if (sheetName === "movimientos") {
      var dataRange = sheet.getRange("P2:P" + sheet.getLastRow());
      var dateRange = sheet.getRange("S2:S" + sheet.getLastRow());
    } else if (sheetName === "accesorios") {
      var dataRange = sheet.getRange("K2:K" + sheet.getLastRow());
      var dateRange = sheet.getRange("N2:N" + sheet.getLastRow());
    } else {
      var dataRange = sheet.getRange("M2:M" + sheet.getLastRow());
      var dateRange = sheet.getRange("S2:S" + sheet.getLastRow());
    }
  
    var dateValues = dateRange.getValues();
    var dataValues = dataRange.getValues();
  
    for (var i = 0; i < dataValues.length; i++) {
      if (dataValues[i][0] == "Llegó" && dateValues[i][0] == "") {
        dateValues[i][0] = new Date();
      }
    }
  
    dateRange.setValues(dateValues);
  }
  
  function actualizarFechasLlego() {
    actualizarFecha("equipos usados");
    actualizarFecha("equipos nuevos");
    actualizarFecha("movimientos");
    actualizarFecha("equipos entrega inmediata");
    actualizarFecha("equipos reservados");
    actualizarFecha("accesorios");
  }
  
  // Actualización monedas cada 1 hora
  
  function dolarBlue() {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var hoja_viz = ss.getSheetByName("capital");
    var celda = hoja_viz.getRange("L3");
    celda.setFormula('=SUBSTITUTE(INDEX(IMPORTXML("https://dolarhoy.com/cotizaciondolarblue";"//div[@class=\'value\']");2;1);".";",")');
  }                  
  
  function clpBlue() {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var hoja_viz = ss.getSheetByName("capital");
    var celda = hoja_viz.getRange("L5");
    celda.setFormula('=SUBSTITUTE(INDEX(IMPORTXML("https://dolarhoy.com/cotizacion-peso-chileno";"//div[@class=\'value\']");2;1);".";",")');
  }         
  
  function euroBlue() {
    var url = "https://www.precioeuroblue.com.ar/";
    var html = UrlFetchApp.fetch(url).getContentText();
    var startIndex = html.indexOf('<span class="label reference">', html.indexOf('<span class="label reference">') + 1) + '<span class="label reference">'.length;
    var endIndex = html.indexOf('</span>', startIndex);
    var valorEuroBlue = html.substring(startIndex, endIndex);
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var hoja_viz = ss.getSheetByName("capital");
    var celda = hoja_viz.getRange("L4");
    celda.setValue(parseFloat(valorEuroBlue.replace(",", ".")));  
  }
  
  //Menú compras 
  
  //Función para encontrar la hoja correspondiente para cada valor de C6:C14
  function obtenerNombreHoja(estado) {
    var nombreHoja = "";
    
    if (estado === "Nuevo") {
      nombreHoja = "equipos nuevos";
    } else if (estado === "Usado") {
      nombreHoja = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      nombreHoja = "equipos entrega inmediata";
    } else {
      nombreHoja = "equipos reservados";
    }
    
    return nombreHoja;
  }
  
  // Función que carga los valores correspondientes en cada hoja a pesar de ser hoja diferentes.
  function cargarCompras() {
    var ss= SpreadsheetApp.getActiveSpreadsheet()
    var menu_principal = ss.getSheetByName('menu principal');
    var movimientos = ss.getSheetByName('movimientos');
    var menu_accesorios = ss.getSheetByName('menu accesorios');
    var equipos_nuevos = ss.getSheetByName('equipos nuevos');
    var accesorios = ss.getSheetByName('accesorios');
    var balance = ss.getSheetByName('balance');
    var v_balance = ss.getSheetByName('visualizado de balance');
  
    // Obtener los datos de las celdas en la hoja "Menu completo"
    var fecha = menu_principal.getRange("B2:I2").getValues();
    var persona = menu_principal.getRange("B3:I3").getValues();
    var anotaciones = menu_principal.getRange("B4:J4").getValues();
    var imei = menu_principal.getRange("A6:A14").getValues();
    var modelo = menu_principal.getRange("B6:B14").getValues();
    var estado = menu_principal.getRange("C6:C14").getValues();
    var cuenta = menu_principal.getRange("D6:D14").getValues();
    var pago_parcial = menu_principal.getRange("E6:E14").getValues();
    var precio_compra = menu_principal.getRange("F6:F14").getValues();
    var precio_venta = menu_principal.getRange("G6:G14").getValues();
    var ganancias = [];
      for (var i = 0; i < precio_compra.length; i++) {
        var compra = precio_compra[i][0];
        var venta = precio_venta[i][0];
        var ganancia = venta - compra;
        ganancias.push([ganancia]);
      }
    var moneda = menu_principal.getRange("H6:H14").getValues();
    var estado_pago = menu_principal.getRange("I6:I14").getValues();
    var compras = 'Compra'
    var responsable = Session.getActiveUser().getEmail();
    
    var resultadoResta = []
  
    for (var i = 0; i < precio_compra.length; i++) {
      if(precio_compra[i][0] > 0){
        var compra = precio_compra[i][0];
        var pp = pago_parcial[i][0];
        var result =  (compra - pp)*-1;
        resultadoResta.push([result]);
      } else{ 
        var pv = precio_venta[i][0]
        var precio_venta_negativo = pv * -1
        var resultado = (pv)*-1
        resultadoResta.push([resultado])
      }
    }
  
    var pago_parcial_negativo = []
    for (var i = 0; i< pago_parcial.length; i++) {
      pp = pago_parcial[i][0]
      result = pp * -1
      pago_parcial_negativo.push([result])
    }
    
    var precio_compra_negativo = []
    for (var i = 0; i< precio_compra.length; i++) {
      pp = precio_compra[i][0]
      result = pp * -1
      precio_compra_negativo.push([result])
    }
  
    // Obtener el último ID de la hoja "Movimientos" y sumar 1 para generar un nuevo ID
    var ultimoId = parseInt(movimientos.getRange("B2").getDisplayValue()) || 0;
    var nuevoId = ultimoId + 1;
  
    
    // Iterar sobre los datos de las variables para agregar cada fila a la hoja "Movimientos"
  
    for (var i = modelo.length - 1; i >= 0; i--) {
      for (var j = 0; j < modelo[i].length; j++) {
        if (modelo[i][j] !== "") { // Verificar si hay un valor en la variable "modelo"
          if (estado_pago[i][j] === "Pendiente"|| estado_pago[i][j] === "Envian") { // Verificar si el estado es "Pendiente"
            if (pago_parcial[i][j] === "") { // Si pago parcial está vacío, hacer una carga simple
              var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], cuenta[i][j], 
              resultadoResta[i][j],"",resultadoResta[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j], fecha, responsable];
              movimientos.insertRowAfter(1);
              movimientos.getRange("A2:R2").setValues([nuevaFila]);
              nuevoId++;
              
              } else { // Si hay un pago parcial, hacer dos registros
              var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], cuenta[i][j], resultadoResta[i][j], 
              resultadoResta[i][j], resultadoResta[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j], fecha, responsable];
              movimientos.insertRowAfter(1);
              movimientos.getRange("A2:R2").setValues([nuevaFila]);
              nuevoId++;
              var nueva_fila_adicional = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], "Caja", 
              pago_parcial_negativo[i][j],pago_parcial_negativo[i][j] ,pago_parcial_negativo[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j], fecha, responsable];
              movimientos.insertRowAfter(1);
              movimientos.getRange("A2:R2").setValues([nueva_fila_adicional]);
              nuevoId++;
            }
          } else if (pago_parcial[i][j] !== "") { // Si el estado no es "Pendiente" pero hay un pago parcial, hacer dos registros
            var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], cuenta[i][j], 
            resultadoResta[i][j], resultadoResta[i][j],resultadoResta[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j], "", responsable];
            movimientos.insertRowAfter(1);
            movimientos.getRange("A2:R2").setValues([nuevaFila]);
            nuevoId++;
            
            var nueva_fila_adicional = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], "Caja", pago_parcial_negativo[i][j], 
            pago_parcial_negativo[i][j], pago_parcial_negativo[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j],"",responsable];
  
            movimientos.insertRowAfter(1);
  
            movimientos.getRange("A2:R2").setValues([nueva_fila_adicional]);
            
            nuevoId++;  
            } else { // Si el estado no es "Pendiente" y no hay pago parcial, hacer una carga simple
                var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], cuenta[i][j], 
                resultadoResta[i][j], "",resultadoResta[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j], "", responsable];
                movimientos.insertRowAfter(1);
                movimientos.getRange("A2:R2").setValues([nuevaFila]);
                nuevoId++;
              }
        }
      }
    }
  
    // Actualizar el valor del último ID en la hoja "Movimientos"
    var id_correcto = movimientos.getRange("B2").setValue(nuevoId - 1);
    if (id_correcto == 0) {
      movimientos.getRange("B2").setValue(1);
    } else {
      movimientos.getRange("B2").setValue(nuevoId - 1);
    }
    
    
    // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
    var numFilas = movimientos.getLastRow() - 1;
    
    // Establecer el formato para las filas agregadas automáticamente
    var formatoBasico = movimientos.getRange(2,1,numFilas,20).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true,true,true,true,true,true);
  
  
    // AGREGAR VALORES A EQUIPOS 
  
    var valores = estado.map(function(row) {
      return row[0];
    }).filter(function(value) {
      return value !== "";
    });
  
    for (var i = 0; i < valores.length; i++) {
      var hojaNombre = obtenerNombreHoja(valores[i]);
      var hoja = ss.getSheetByName(hojaNombre);
  
      var nuevoId = hoja.getRange("B2").getValue() + 1;
      if (nuevoId == ""){ 
        nuevoId = 1; 
      }
  
      for (var j = 0; j < modelo[i].length; j++) {
        if (modelo[i][j] !== "") { // Verificar si hay un valor en la variable "modelo"
          if (estado_pago[i][j] === "Pendiente"|| estado_pago[i][j] === "Envian") { // Verificar si el estado es "Pendiente"
            if (pago_parcial[i][j] === "") { // Si pago parcial está vacío, hacer una carga simple
              var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], precio_compra[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j], fecha, responsable];
  
              hoja.insertRowAfter(1);
              hoja.getRange("A2:O2").setValues([nuevaFila]);
              nuevoId++;
              // Actualizar el valor del último ID en la hoja "Movimientos"
              hoja.getRange("B2").setValue(nuevoId - 1);
              // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
              var numFilas = hoja.getLastRow() - 1;
              // Establecer el formato para las filas agregadas automáticamente
              var formatoBasico = hoja.getRange(2,1,numFilas,19).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true,true,true,true,true,true);
  
            } else if (pago_parcial[i][j] !== "") { // Si el estado no es "Pendiente" pero hay un pago parcial, hacer un solo registro que tenga el valor de precio de compra.
              var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], precio_compra[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j], "", responsable];
  
              hoja.insertRowAfter(1)
              hoja.getRange("A2:O2").setValues([nuevaFila]);
              nuevoId++;
              // Actualizar el valor del último ID en la hoja "Movimientos"
              hoja.getRange("B2").setValue(nuevoId - 1);
              // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
              var numFilas = hoja.getLastRow() - 1;
              // Establecer el formato para las filas agregadas automáticamente
              var formatoBasico = hoja.getRange(2,1,numFilas,19).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true,true,true,true,true,true);
            } 
          } else { // Si el estado no es "Pendiente" y no hay pago parcial, hacer una carga simple
            var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, imei[i][0], modelo[i][j], estado[i][j], precio_compra[i][j], precio_venta[i][j], ganancias[i][j], moneda[i][j], estado_pago[i][j], "", responsable];
  
            hoja.insertRowAfter(1)
            hoja.getRange("A2:O2").setValues([nuevaFila]);
            nuevoId++;  
            // Actualizar el valor del último ID en la hoja "Movimientos"
            hoja.getRange("B2").setValue(nuevoId - 1);
            // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
            var numFilas = hoja.getLastRow() - 1;
            // Establecer el formato para las filas agregadas automáticamente
            var formatoBasico = hoja.getRange(2,1,numFilas,19).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true,true,true,true,true,true);
          }
        }
      }            
    }
  
  
  
      var rango= menu_principal.getRange("B3:J4");
      var rango2= menu_principal.getRange("A6:J14");
      rango.clearContent();
      rango2.clearContent();
  }
  
  
  // Menú Ventas 
    //Búsqueda por ID
  function BúsquedaEquipos() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu principal');
    var id = sheet.getRange("B21").getValue();
    var estado = sheet.getRange("C23").getValue();
    var equipoSheetName;
  
    if (estado == "" || estado == null){
      var mensaje = "Por favor, especifique una hoja en la cual buscar el equipo";
      Logger.log(mensaje);
    }
  
    if (estado === "Nuevo") {
      equipoSheetName = "equipos nuevos";
    } else if (estado === "Usado") {
      equipoSheetName = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      equipoSheetName = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      equipoSheetName = "equipos reservados";
    } else {
      // Estado no válido, no se realiza la búsqueda y se envía mensaje
      var mensaje = "Por favor, especifique una hoja en la cual buscar el equipo";
      Logger.log(mensaje);
    }
  
    var equipoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(equipoSheetName);
    var rangeValues = equipoSheet.getRange("B2:B").getValues();
    var result = [];
    var id_result = [];
    var date_result = [];
    var persona_result = [];
    var anotations_result = [];
  
    for (var i = 0; i < rangeValues.length; i++) {
      if (rangeValues[i][0] === id) {
        var row = equipoSheet.getRange(i + 2, 1, 1, 15).getValues()[0];
        var Fecha = row[0];
        var ID = row[1];
        var Persona = row[2];
        var Anotaciones = row[3];
        var IMEI = row[5];
        var Modelo = row[6];
        var Estado = row[7];
        var PrecioVenta = row[9];
        var Moneda = row[11];
        var Estado_pago = row[12];
        result = [IMEI, Modelo, Estado, "", "", PrecioVenta, Moneda, Estado_pago];
        date_result = [Fecha];
        id_result = [ID];
        persona_result = [Persona];
        anotations_result = [Anotaciones];
        break;
      }
    }
  
    var outputRange = sheet.getRange("A23:H23");
    var outputRange2 = sheet.getRange("B18");
    var outputRange3 = sheet.getRange("B19");
    var outputRange4 = sheet.getRange("B20");
    var outputRange5 = sheet.getRange("B21");
  
    outputRange.clearContent();
    outputRange2.clearContent();
    outputRange3.clearContent();
    outputRange4.clearContent();
    outputRange5.clearContent();
  
    outputRange.setValues([result]);
    outputRange2.setValues([date_result]);
    outputRange3.setValues([persona_result]);
    outputRange4.setValues([anotations_result]);
    outputRange5.setValues([id_result]);
  }
    //Búsqueda por IMEI 
  function BúsquedaEquipos2() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu principal');
    var id = sheet.getRange("A23").getValue();
    var estado = sheet.getRange("C23").getValue();
    var equipoSheetName;
  
    if (estado === "Nuevo") {
      equipoSheetName = "equipos nuevos";
    } else if (estado === "Usado") {
      equipoSheetName = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      equipoSheetName = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      equipoSheetName = "equipos reservados";
    } else {
      // Estado no válido, no se realiza la búsqueda y se envía mensaje
      var mensaje = "Por favor, especifique una hoja en la cual buscar el equipo";
      Logger.log(mensaje);
    }
  
    var equipoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(equipoSheetName);
    var rangeValues = equipoSheet.getRange("F2:F").getValues();
    var result = [];
    var id_result = [];
    var date_result = [];
    var persona_result = [];
    var anotations_result = [];
  
    for (var i = 0; i < rangeValues.length; i++) {
      if (rangeValues[i][0] === id) {
        var row = equipoSheet.getRange(i + 2, 1, 1, 15).getValues()[0];
        var Fecha = row[0];
        var ID = row[1];
        var Persona = row[2];
        var Anotaciones = row[3];
        var IMEI = row[5];
        var Modelo = row[6];
        var Estado = row[7];
        var PrecioVenta = row[9];
        var Moneda = row[11];
        var Estado_pago = row[12];
        result = [IMEI, Modelo, Estado, "", "", PrecioVenta, Moneda, Estado_pago];
        date_result = [Fecha];
        id_result = [ID];
        persona_result = [Persona];
        anotations_result = [Anotaciones];
        break;
      }
    }
  
    var outputRange = sheet.getRange("A23:H23");
    var outputRange2 = sheet.getRange("B18");
    var outputRange3 = sheet.getRange("B19");
    var outputRange4 = sheet.getRange("B20");
    var outputRange5 = sheet.getRange("B21");
  
    outputRange.clearContent();
    outputRange2.clearContent();
    outputRange3.clearContent();
    outputRange4.clearContent();
    outputRange5.clearContent();
  
    outputRange.setValues([result]);
    outputRange2.setValues([date_result]);
    outputRange3.setValues([persona_result]);
    outputRange4.setValues([anotations_result]);
    outputRange5.setValues([id_result]);
  }
  
    //Cargar la venta
  
  function cargarVentas() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menu_principal = ss.getSheetByName('menu principal');
    var movimientos = ss.getSheetByName('movimientos');
    var equipos_nuevos = ss.getSheetByName('equipos nuevos');
    var equipos_usados = ss.getSheetByName('equipos usados');
    var equipos_entrega_inmediata = ss.getSheetByName('equipos entrega inmediata');
    var equipos_reservados = ss.getSheetByName('equipos reservados');
  
    // Obtener los datos de las celdas en la hoja "Menu completo"
    var fecha = menu_principal.getRange("B18:I18").getValues();
    var persona = menu_principal.getRange("B19:I19").getValues();
    var anotaciones = menu_principal.getRange("B20:I20").getValues();
    var imei = menu_principal.getRange("A23:A32").getValues();
    var modelo = menu_principal.getRange("B23:B32").getValues();
    var estado_del_equipo = menu_principal.getRange("C23:C32").getValues();
    var cuenta = menu_principal.getRange("D23:D32").getValues();
    var pago_parcial = menu_principal.getRange("E23:E32").getValues();
    var precio_venta = menu_principal.getRange("F23:F32").getValues();
    var moneda = menu_principal.getRange("G23:G32").getValues();
    var estado_pago = menu_principal.getRange("H23:H32").getValues();
    var ventas = 'Ventas'
    var responsable = Session.getActiveUser().getEmail();
    var resultadoResta = [];
    for (var i = 0; i < precio_venta.length; i++) {
      var venta = precio_venta[i][0];
      var pp = pago_parcial[i][0];
      var result = venta - pp;
      resultadoResta.push([result]);
    }
  
    // Obtener el último ID de la hoja "Movimientos" y sumar 1 para generar un nuevo ID
    var ultimoId = parseInt(movimientos.getRange("B2").getDisplayValue()) || 0;
    var nuevoId = ultimoId + 1;
  
    // Iterar sobre los datos de las variables para agregar cada fila a la hoja "Movimientos"
    for (var i = modelo.length - 1; i >= 0; i--) {
      for (var j = 0; j < modelo[i].length; j++) {
        if (modelo[i][j] !== "") { // Verificar si hay un valor en la variable "modelo"
          if (pago_parcial[i][j] === "") { // Si pago parcial está vacío, hacer una carga simple
            var nuevaFila = [fecha, nuevoId, persona, anotaciones, ventas, imei[i][0], modelo[i][j], estado_del_equipo[i][j], cuenta[i][j],precio_venta[i][j] ,"", "", precio_venta[i][j], "", moneda[i][j], estado_pago[i][j], "", responsable];
            movimientos.insertRowAfter(1);
            movimientos.getRange("A2:R2").setValues([nuevaFila]);
            nuevoId++;
          } else { // Si hay un pago parcial, hacer dos registros
            var nuevaFila = [fecha, nuevoId, persona, anotaciones, ventas, imei[i][0], modelo[i][j], estado_del_equipo[i][j], cuenta[i][j], resultadoResta[i][j], resultadoResta[i][j],"", resultadoResta[i][j], "", moneda[i][j], estado_pago[i][j], "", responsable];
            movimientos.insertRowAfter(1);
            movimientos.getRange("A2:R2").setValues([nuevaFila]);
            nuevoId++;
            var nueva_fila_adicional = [fecha, nuevoId, persona, anotaciones, ventas, imei[i][0], modelo[i][j], estado_del_equipo[i][j], "Caja", pago_parcial[i][j],pago_parcial[i][j] ,"", pago_parcial[i][j], "", moneda[i][j], estado_pago[i][j], "", responsable];
            movimientos.insertRowAfter(1);
            movimientos.getRange("A2:R2").setValues([nueva_fila_adicional]);
            nuevoId++;
          }
        }
      }
    }
  
  
    // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
    var numFilas = movimientos.getLastRow() - 1;
  
    // Establecer el formato para las filas agregadas automáticamente
    var formatoBasico = movimientos.getRange(2, 1, numFilas, 20).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true, true, true, true, true, true);
  
    // ELIMINAR VALOR COINCIDENTE EN LA RESPECTIVA HOJA
    var id = menu_principal.getRange("B21:I21").getValue();
    var modelos = menu_principal.getRange("B23:B32").getValues();
    var modelosArray = modelos.map(function (row) {
      return row[0];
    });
  
    for (var i = 0; i < modelosArray.length; i++) {
      if (modelosArray[i] !== "") {
        var estado = menu_principal.getRange("C" + (i + 23)).getValue();
        if (estado === "Entrega inmediata") {
          var equiposEntregaInmediata = ss.getSheetByName("equipos entrega inmediata");
          var equiposEntregaInmediataDatos = equiposEntregaInmediata.getDataRange().getValues();
          for (var j = equiposEntregaInmediataDatos.length - 1; j >= 0; j--) {
            if (equiposEntregaInmediataDatos[j][1] === id) {
              equiposEntregaInmediata.deleteRow(j + 1);
            }
          }
        } else if (estado === "Reservado") {
          var equiposReservados = ss.getSheetByName("equipos reservados");
          var equiposReservadosDatos = equiposReservados.getDataRange().getValues();
          for (var j = equiposReservadosDatos.length - 1; j >= 0; j--) {
            if (equiposReservadosDatos[j][1] === id) {
              equiposReservados.deleteRow(j + 1);
            }
          }
        } else if (estado === "Usado") {
          var equiposUsados = ss.getSheetByName("equipos usados");
          var equiposUsadosDatos = equiposUsados.getDataRange().getValues();
          for (var j = equiposUsadosDatos.length - 1; j >= 0; j--) {
            if (equiposUsadosDatos[j][1] === id) {
              equiposUsados.deleteRow(j + 1);
            }
          }
        } else {
          var equiposNuevos = ss.getSheetByName("equipos nuevos");
          var equiposNuevosDatos = equiposNuevos.getDataRange().getValues();
          for (var j = equiposNuevosDatos.length - 1; j >= 0; j--) {
            if (equiposNuevosDatos[j][1] === id) {
              equiposNuevos.deleteRow(j + 1);
            }
          }
        }
      }
    }
  
    var rango1 = menu_principal.getRange("B18:I21");
    var rango2 = menu_principal.getRange("A23:I32");
    rango1.clearContent();
    rango2.clearContent();
  }
   
  
  // Menú compra de accesorios
  function cargarComprasAccesorios() {
    var movimientos = ss.getSheetByName('movimientos');
    var menu_accesorios = ss.getSheetByName('menu accesorios');
    var accesorios = ss.getSheetByName('accesorios');
  
    // Obtener los datos de las celdas en la hoja "Menu completo"
    var fecha = menu_accesorios.getRange("B2:F2").getValues();
    var persona = menu_accesorios.getRange("B3:F3").getValues();
    var anotaciones = menu_accesorios.getRange("B4:F4").getValues();
    var modelo = menu_accesorios.getRange("A6:A14").getValues();
    var cantidad = menu_accesorios.getRange("B6:B14").getValues();
    var cuenta = menu_accesorios.getRange("C6:C14").getValues();
    var precio_original = menu_accesorios.getRange("D6:D14").getValues();
    var moneda = menu_accesorios.getRange("E6:E14").getValues();
    var estado_pago = menu_accesorios.getRange("F6:F14").getValues();
    var compras = 'Compra'
    var responsable = Session.getActiveUser().getEmail();
    var precio_compra = []
  
     for (var i = 0; i < precio_original.length; i++) {
      var  original= precio_original[i][0];
      var compra = original * -1
      precio_compra.push([compra]);
    }
  
  
    // Obtener el último ID de la hoja "Movimientos" y sumar 1 para generar un nuevo ID
    var ultimoId = parseInt(movimientos.getRange("B2").getDisplayValue()) || 0;
    var nuevoId = ultimoId + 1;
    
    // Iterar sobre los datos de las variables para agregar cada fila a la hoja "Movimientos"
  
    for (var i = modelo.length - 1; i >= 0; i--) {
      for (var j = 0; j < modelo[i].length; j++) {
        if (modelo[i][j] !== "") { // Verificar si hay un valor en la variable "modelo"
          if (estado_pago[i][j] === "Pendiente"|| estado_pago[i][j] === "Envian") { // Verificar si el estado es "Pendiente"
              var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, "Accesorio", modelo[i][j], "", cuenta[i][j], precio_compra[i][j],
              "",precio_compra[i][j], "", "", moneda[i][j], estado_pago[i][j], fecha, responsable];
              movimientos.insertRowAfter(1);
              movimientos.getRange("A2:R2").setValues([nuevaFila]);
              nuevoId++;
              } else { // Si el estado no es "Pendiente" hacer una carga simple
                var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, "Accesorio", modelo[i][j], "", cuenta[i][j],precio_compra[i][j], "",
                precio_compra[i][j], "", "", moneda[i][j], estado_pago[i][j],"", responsable ];
                movimientos.insertRowAfter(1);
                movimientos.getRange("A2:R2").setValues([nuevaFila]);
                nuevoId++;
              }    
            }
          }
        }
  
    
    // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
    var numFilas = movimientos.getLastRow() - 1;
    
    // Establecer el formato para las filas agregadas automáticamente
    var formatoBasico = movimientos.getRange(2,1,numFilas,20).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true,true,true,true,true,true);
  
  
    // AGREGAR VALORES A ACCESORIOS 
    
     // Obtener el último ID de la hoja "Acceosrios" y sumar 1 para generar un nuevo ID
    var fecha = menu_accesorios.getRange("B2:F2").getValues();
    var persona = menu_accesorios.getRange("B3:F3").getValues();
    var anotaciones = menu_accesorios.getRange("B4:F4").getValues();
    var modelo = menu_accesorios.getRange("A6:A14").getValues();
    var cantidad = menu_accesorios.getRange("B6:B14").getValues();
    var cuenta = menu_accesorios.getRange("C6:C14").getValues();
    var precio_original = menu_accesorios.getRange("D6:D14").getValues();
    var precio_compra = []
  
    for (var i = 0; i < precio_original.length; i++) {
          var original = precio_original[i][0];
          var result =  (original)*-1;
          precio_compra.push([result]);
        }
   
    var moneda = menu_accesorios.getRange("E6:E14").getValues();
    var estado_pago = menu_accesorios.getRange("F6:F14").getValues();
    var compras = 'Compra'
    var responsable = Session.getActiveUser().getEmail();
  
    var accesorios = ss.getSheetByName('accesorios');
  
  
    // Obtener el último ID de la hoja "Accesorios" y sumar 1 para generar un nuevo ID
    var ultimoId = parseInt(accesorios.getRange("B2").getDisplayValue()) || 0;
    var nuevoId = ultimoId + 1;
    
    // Iterar sobre los datos de las variables para agregar cada fila a la hoja "Movimientos"
  
    for (var i = modelo.length - 1; i >= 0; i--) {
      for (var j = 0; j < modelo[i].length; j++) {
        if (modelo[i][j] !== "") { // Verificar si hay un valor en la variable "modelo"
          if (estado_pago[i][j] === "Pendiente"|| estado_pago[i][j] === "Envian") { // Verificar si el estado es "Pendiente"
              var fila = [fecha, nuevoId, persona, anotaciones, modelo[i][j], cantidad[i][j], cuenta[i][j], precio_compra[i][j],"", moneda[i][j], estado_pago[i][j], fecha, responsable];
              accesorios.insertRowAfter(1);
              accesorios.getRange("A2:M2").setValues([fila]);
              nuevoId++;
              }else { // Si el estado no es "Pendiente" hacer una carga simple
                var fila = [fecha, nuevoId, persona, anotaciones, modelo[i][j],cantidad[i][j], cuenta[i][j], precio_compra[i][j],"",moneda[i][j], estado_pago[i][j], "", responsable];
                accesorios.insertRowAfter(1);
                accesorios.getRange("A2:M2").setValues([fila]);
                nuevoId++;
              }    
        } 
      }
    }
    
    // Obtener el número de filas en la hoja "Accesorios" después de agregar las nuevas filas
    var numFilas = accesorios.getLastRow() - 1;
    
    // Establecer el formato para las filas agregadas automáticamente
    var formatoBasico = accesorios.getRange(2,1,numFilas,14).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true,true,true,true,true,true);
  
    var rango= menu_accesorios.getRange("B3:F4");
    var rango2= menu_accesorios.getRange("A6:F14");
    rango.clearContent();
    rango2.clearContent();
  }
  
    // Búsqueda de accesorios para vender
  function BuscarAccesorios() {
  
    var ss= SpreadsheetApp.getActiveSpreadsheet()
    var menu_accesorios = ss.getSheetByName('menu accesorios');
    var accesorios = ss.getSheetByName('accesorios');
    var modelo = menu_accesorios.getRange("A23").getValue();
    var rangeModelos = accesorios.getRange("E2:E").getValues();
    var result = [];
    var id_result = [];
    var date_result = [];
    var id_result = [];
    var provider_result= [];
    var anotations_result = []; 
  
    for (var i = 0; i < rangeModelos.length; i++) {
    if ((rangeModelos[i][0]).toLowerCase() === modelo.toLowerCase()) {
        var row = accesorios.getRange("A" + (i+2) + ":M" + (i+2)).getValues()[0];
        var Fecha = row[0];
        var ID = row[1];
        var Anotaciones = row[3];
        var Modelo = row[4];
        var Cantidad = row [5];
        var Cuenta = row[6];
        var PrecioCompra = row[7];
        var Moneda = row[9];
        var Estado = row[10];
        result = [Modelo, Cantidad, Cuenta, Math.abs(PrecioCompra), Moneda, Estado];
        date_result = [Fecha];
        id_result = [ID];
        anotations_result = [Anotaciones]; 
        break;
      }
    }
    var rango1 = menu_accesorios.getRange("B18:F21");
    var rango2 = menu_accesorios.getRange("A23:F32");
    rango1.clearContent();
    rango2.clearContent();
  
    var outputRange = menu_accesorios.getRange("A23:F23");
    var outputRange2 = menu_accesorios.getRange("B18");
    var outputRange3 = menu_accesorios.getRange("B19");
    var outputRange4 = menu_accesorios.getRange("B21");
    outputRange.setValues([result]);
    outputRange2.setValues([date_result]);
    outputRange3.setValues([id_result]);
    outputRange4.setValues([anotations_result]);
  }
  
  // Menú ventas de accesorios
  function cargarVentasAccesorios() {
    var ss= SpreadsheetApp.getActiveSpreadsheet()
    var menu_accesorios = ss.getSheetByName('menu principal');
    var movimientos = ss.getSheetByName('movimientos');
    var menu_accesorios = ss.getSheetByName('menu accesorios');
    var accesorios = ss.getSheetByName('accesorios');
  
    // Obtener los datos de las celdas en la hoja "Menu accesorios"
    var fecha = menu_accesorios.getRange("B18:I18").getValues();
    var persona = menu_accesorios.getRange("B20:I20").getValues();
    var anotaciones = menu_accesorios.getRange("B21:I21").getValues();
    var modelo = menu_accesorios.getRange("A23:A31").getValues();
    var cantidad = menu_accesorios.getRange("B23:B31").getValues();
    var cuenta = menu_accesorios.getRange("C23:C31").getValues();
    var precio_venta = menu_accesorios.getRange("D23:D32").getValues();
    var moneda = menu_accesorios.getRange("E23:E32").getValues();
    var estado_pago = menu_accesorios.getRange("F23:F32").getValues();
    var ventas = 'Ventas'
    var responsable = Session.getActiveUser().getEmail();
  
    // Obtener el último ID de la hoja "Movimientos" y sumar 1 para generar un nuevo ID
    var ultimoId = parseInt(movimientos.getRange("B2").getDisplayValue()) || 0;
    var nuevoId = ultimoId + 1;
    
    // Iterar sobre los datos de las variables para agregar cada fila a la hoja "Movimientos"
  
    for (var i = modelo.length - 1; i >= 0; i--) {
      for (var j = 0; j < modelo[i].length; j++) {
        if (modelo[i][j] !== "") { // Verificar si hay un valor en la variable "modelo"
          if (estado_pago[i][j] === "Pendiente"|| estado_pago[i][j] === "Envian") { 
              var nuevaFila = [fecha, nuevoId, persona, anotaciones, ventas, "Accesorio", modelo[i][j], "", cuenta[i][j],precio_venta[i][j], 
                "","", precio_venta[i][j],"", moneda[i][j], estado_pago[i][j], fecha, responsable];
                movimientos.insertRowAfter(1);
                movimientos.getRange("A2:R2").setValues([nuevaFila]);
                nuevoId++;
              }else{ 
                var nuevaFila = [fecha, nuevoId, persona, anotaciones, ventas, "Accesorio", modelo[i][j], "", cuenta[i][j],precio_venta[i][j], 
                "","", precio_venta[i][j],"", moneda[i][j], estado_pago[i][j], "", responsable];
                movimientos.insertRowAfter(1);
                movimientos.getRange("A2:R2").setValues([nuevaFila]);
                nuevoId++;
              }
        }
      }
    }
      
    
    // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
    var numFilas = movimientos.getLastRow() - 1;
    
    // Establecer el formato para las filas agregadas automáticamente
    var formatoBasico = movimientos.getRange(2,1,numFilas,20).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true,true,true,true,true,true);
  
  
    // Disminuir cantidades, si es 0 eliminar la coincidencia
  
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menu_accesorios = ss.getSheetByName("menu accesorios");
    var accesorios = ss.getSheetByName("accesorios");
  
    var idBuscado = menu_accesorios.getRange("B19:F19").getValue();
    var valoresResta = menu_accesorios.getRange("B23").getValue();
  
    var ultimaFila = accesorios.getLastRow();
    var filasAEliminar = [];
  
    for (var i = 2; i <= ultimaFila; i++) {
      if (accesorios.getRange(i, 2).getValue() === idBuscado) {
        var valorFila = accesorios.getRange(i, 6).getValue();
        var resultadoResta = valorFila - valoresResta;
        var precio_compra = accesorios.getRange(i, 8).getValue();
        var precio_unitario = ( precio_compra / valorFila);
        var precio_final = precio_compra - precio_unitario;
  
        if (resultadoResta <= 0) {
          filasAEliminar.push(i);
        } else {
          accesorios.getRange(i, 6).setValue(resultadoResta);
          accesorios.getRange(i, 8).setValue(precio_final);
  
        }
      }
    }
  
      // Eliminar filas marcadas para eliminación
      for (var j = filasAEliminar.length - 1; j >= 0; j--) {
        accesorios.deleteRow(filasAEliminar[j]);
      }
    
    
    
    var rango1 = menu_accesorios.getRange("B18:F21");
    var rango2 = menu_accesorios.getRange("A23:F31");
    rango1.clearContent();
    rango2.clearContent();
  } 
  
  // Menú Equipos con Problemas
  
   // Búsqueda de equipos rotos
  
      //Búsqueda por ID
  function BúsquedaEquiposDefectuosos() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu principal');
    var id = sheet.getRange("B56").getValue();
    var estado = sheet.getRange("C58").getValue();
    var equipoSheetName;
  
    if (estado === "Nuevo") {
      equipoSheetName = "equipos nuevos";
    } else if (estado === "Usado") {
      equipoSheetName = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      equipoSheetName = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      equipoSheetName = "equipos reservados";
    }  else if (estado === "Con problemas") {
      equipoSheetName = "equipos con problemas";
    } else {
      // Estado no válido, no se realiza la búsqueda
      return;
    }
  
    var equipoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(equipoSheetName);
    var rangeValues = equipoSheet.getRange("B2:B").getValues();
  
      var result = [];
      var id_result = [];
      var date_result = [];
      var id_result = [];
      var provider_result= [];
      var anotations_result = []; 
  
      //nuevo for loop
      for (var i = 0; i < rangeValues.length; i++) {
      if (rangeValues[i][0] === id) {
        var row = equipoSheet.getRange(i + 2, 1, 1, 15).getValues()[0];
        var Fecha = row[0];
        var ID = row[1];
        var IMEI = row[5];
        var Modelo = row[6];
        var Estado = row[7];
        var Cliente = row[5];
        var PrecioCompra = row[8];
        var Moneda = row[11];
        date_result = [Fecha];
        id_result = [ID];
        result = [IMEI, Modelo, Estado, "","","",PrecioCompra,Moneda,""];
        break;
      }
    }
      var outputRange = sheet.getRange("A58:I58");
      var outputRange2 = sheet.getRange("B55");
      var outputRange4 = sheet.getRange("B56");
  
      outputRange.clearContent();
      outputRange2.clearContent();
      outputRange4.clearContent();
  
      outputRange.setValues([result]);
      outputRange2.setValues([date_result]);
      outputRange4.setValues([id_result]);
  }
  
      //Búsqueda por IMEI 
  function BúsquedaEquiposDefectuosos2() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu principal');
    var imei = sheet.getRange("A58").getValue();
    var estado = sheet.getRange("C58").getValue();
    var equipoSheetName;
  
    if (estado === "Nuevo") {
      equipoSheetName = "equipos nuevos";
    } else if (estado === "Usado") {
      equipoSheetName = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      equipoSheetName = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      equipoSheetName = "equipos reservados";
    }  else if (estado === "Con problemas") {
      equipoSheetName = "equipos con problemas";
    } else {
      // Estado no válido, no se realiza la búsqueda
      return;
    }
  
    var equipoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(equipoSheetName);
    var rangeValues = equipoSheet.getRange("F2:F").getValues();
  
      var result = [];
      var id_result = [];
      var date_result = [];
      var id_result = [];
      var provider_result= [];
      var anotations_result = []; 
  
      //nuevo for loop
      for (var i = 0; i < rangeValues.length; i++) {
      if (rangeValues[i][0] === imei) {
        var row = equipoSheet.getRange(i + 2, 1, 1, 15).getValues()[0];
        var Fecha = row[0];
        var ID = row[1];
        var IMEI = row[5];
        var Modelo = row[6];
        var Estado = row[7];
        var Cliente = row[5];
        var PrecioCompra = row[8];
        var Moneda = row[11];
        date_result = [Fecha];
        id_result = [ID];
        result = [IMEI, Modelo, Estado, "","","",PrecioCompra,Moneda,""];
        break;
      }
    }
      var outputRange = sheet.getRange("A58:I58");
      var outputRange2 = sheet.getRange("B55");
      var outputRange4 = sheet.getRange("B56");
  
      outputRange.clearContent();
      outputRange2.clearContent();
      outputRange4.clearContent();
  
      outputRange.setValues([result]);
      outputRange2.setValues([date_result]);
      outputRange4.setValues([id_result]);
  }
  
      //Búsqueda por Modelo
  function BúsquedaEquiposDefectuosos3() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu principal');
    var modelo = sheet.getRange("B58").getValue();
    var estado = sheet.getRange("C58").getValue();
    var equipoSheetName;
  
    if (estado === "Nuevo") {
      equipoSheetName = "equipos nuevos";
    } else if (estado === "Usado") {
      equipoSheetName = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      equipoSheetName = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      equipoSheetName = "equipos reservados";
    }  else if (estado === "Con problemas") {
      equipoSheetName = "equipos con problemas";
    } else {
      // Estado no válido, no se realiza la búsqueda
      return;
    }
  
    var equipoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(equipoSheetName);
    var rangeValues = equipoSheet.getRange("G2:G").getValues();
  
      var result = [];
      var id_result = [];
      var date_result = [];
      var id_result = [];
      var provider_result= [];
      var anotations_result = []; 
  
      //nuevo for loop
      for (var i = 0; i < rangeValues.length; i++) {
      if (rangeValues[i][0] === modelo) {
        var row = equipoSheet.getRange(i + 2, 1, 1, 15).getValues()[0];
        var Fecha = row[0];
        var ID = row[1];
        var IMEI = row[5];
        var Modelo = row[6];
        var Estado = row[7];
        var Cliente = row[5];
        var PrecioCompra = row[8];
        var Moneda = row[11];
        date_result = [Fecha];
        id_result = [ID];
        result = [IMEI, Modelo, Estado, "","","",PrecioCompra,Moneda,""];
        break;
      }
    }
      var outputRange = sheet.getRange("A58:I58");
      var outputRange2 = sheet.getRange("B55");
      var outputRange4 = sheet.getRange("B56");
  
      outputRange.clearContent();
      outputRange2.clearContent();
      outputRange4.clearContent();
  
      outputRange.setValues([result]);
      outputRange2.setValues([date_result]);
      outputRange4.setValues([id_result]);
  }
  
   // Carga de equipos rotos
  function cargaDeEquiposRotos() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('menu principal');
  
    // Paso 1: Obtener los valores de las celdas B56, C58 y D58
    var id = sheet.getRange("B56").getValue();
    var estado = sheet.getRange("C58").getValue();
    var tipoPago = sheet.getRange("D58").getValue();
  
    // Paso 2: Determinar la hoja correspondiente según el estado del equipo
    var equipoSheetName;
  
    if (estado === "Nuevo") {
      equipoSheetName = "equipos nuevos";
    } else if (estado === "Usado") {
      equipoSheetName = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      equipoSheetName = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      equipoSheetName = "equipos reservados";
    } else {
      // Estado no válido, no se realiza la búsqueda
      return;
    }
  
    var equipoSheet = ss.getSheetByName(equipoSheetName);
  
    // Paso 3: Buscar filas coincidentes con el valor de la celda B56 en la columna B de la hoja correspondiente
    var dataRange = equipoSheet.getDataRange();
    var values = dataRange.getValues();
  
    // Paso 4: Modificar los valores de las filas coincidentes
    var problema = sheet.getRange("E58").getValue();
    var precioArreglo = sheet.getRange("F58").getValue(); 
    var servicioTecnico = sheet.getRange("I58").getValue();
  
    var precioTotal = 0;
    var filaCoincidente = null;
  
    for (var i = 1; i < values.length; i++) {
      var rowId = values[i][1]; // Columna B en base 0
  
      if (rowId === id) {
        filaCoincidente = values[i].slice(); // Copiar la fila coincidente con las modificaciones ya aplicadas
        
        // Calcular el precio total solo en las filas que coinciden con el valor de D58
        var precioCompraFila = equipoSheet.getRange("I" + (i + 1)).getValue();
        precioTotal = precioCompraFila + precioArreglo;
  
        // Actualizar el precio total y más datos en la fila coincidente
        filaCoincidente[7] = "Con problemas";
        filaCoincidente[8] = precioTotal;
        filaCoincidente[12] = "Servicio técnico";
        filaCoincidente[15] = precioArreglo;
        filaCoincidente[16] = problema; 
        filaCoincidente[17] = servicioTecnico
  
        // Salir del bucle una vez que se ha encontrado la fila coincidente
        break;
      }
    }
  
    if (filaCoincidente !== null) {
      // Obtener la hoja "equipos con problemas" y agregar la fila coincidente modificada
      var equiposConProblemasSheet = ss.getSheetByName("equipos con problemas");
      var ultimaFila = equiposConProblemasSheet.getLastRow();
      var nuevoId = ultimaFila !== 1 ? equiposConProblemasSheet.getRange(ultimaFila, 2).getValue() + 1 : 1;
  
      // Actualizar el ID en la fila coincidente
      filaCoincidente[1] = parseInt(nuevoId);
  
      // Insertar la fila modificada en la hoja "equipos con problemas"
      equiposConProblemasSheet.insertRowAfter(ultimaFila);
      equiposConProblemasSheet.getRange(ultimaFila + 1, 1, 1, filaCoincidente.length).setValues([filaCoincidente]);
      
      var numFilas = equiposConProblemasSheet.getLastRow() - 1;
    
      // Establecer el formato para las filas agregadas automáticamente
      equiposConProblemasSheet.getRange(2, 1, numFilas, 19).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true, true, true, true, true, true);
  
      // Eliminar la fila coincidente de la hoja correspondiente
      equipoSheet.deleteRow(i + 1);
    }
    //Cargar en movimientos del proceso.
    var menu_principal = ss.getSheetByName('menu principal');
    var movimientos = ss.getSheetByName('movimientos');
    var fecha = menu_principal.getRange("B55:I55").getValue();
    var imei = menu_principal.getRange("A58").getValue();
    var modelo = menu_principal.getRange("B58").getValue();
    var estado_del_equipo = "Con problemas";
    var cuenta = menu_principal.getRange("D58").getValue();
    var problema = menu_principal.getRange("E58").getValue();
    var precio_del_arreglo = menu_principal.getRange("F58").getValue();
    var moneda = menu_principal.getRange("H58").getValue();
    var servicio_tecnico = menu_principal.getRange("I58").getValue();
    var arreglos = "Servicio Técnico";
    var responsable = Session.getActiveUser().getEmail();
  
  
    // Obtener el último ID de la hoja "Movimientos" y sumar 1 para generar un nuevo ID
    var ultimoId = parseInt(movimientos.getRange("B2").getDisplayValue()) || 0;
    var nuevoId = ultimoId + 1;
    
    // Iterar sobre los datos de las variables para agregar cada fila a la hoja "Movimientos"
    var nuevaFila = [fecha, nuevoId, "", "", arreglos, imei, modelo, estado_del_equipo, cuenta, -precioArreglo,"",
                     -precioArreglo, "", "", moneda, "", "", responsable];
    movimientos.insertRowAfter(1);
    movimientos.getRange("A2:R2").setValues([nuevaFila]);
    
    // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
    var numFilasMovimientos = movimientos.getLastRow() - 1;
    
    // Establecer el formato para las filas agregadas automáticamente en la hoja "Movimientos"
    var formatoBasicoMovimientos = movimientos.getRange(2, 1, numFilasMovimientos, 20).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true, true, true, true, true, true);
  
    var rango1 = menu_principal.getRange("B55:I56");
    var rango2 = menu_principal.getRange("A58:I58");
    rango1.clearContent();
    rango2.clearContent();
  }
  
   // Salida de equipos rotos
  function salidaEquiposRotos() {
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu principal');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  
    // Paso 1: Obtener los valores de las celdas B56, C58 y D58
    var id = mainSheet.getRange("B56").getValue();
    var estadoEquipo = mainSheet.getRange("C58").getValue();
    var tipoPago = mainSheet.getRange("D58").getValue();
  
    // Paso 2: Determinar la hoja correspondiente según el estado del equipo y obtener el estado correspondiente
    var equipoSheetName;
    var estado;
  
    if (estadoEquipo === "Con problemas") {
      equipoSheetName = "equipos con problemas";
      estado = "Con problemas";
    } else {
      var mensaje = "Estado no válido. Solo para salida de equipos con problemas";
      Logger.log(mensaje);
      return;
    }
  
    var equipoSheet = ss.getSheetByName(equipoSheetName);
    var dataRange = equipoSheet.getDataRange();
    var data = dataRange.getValues();
  
    var filaCoincidente = null;
  
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var idEquipo = row[1]; // Columna B (ID)
  
      if (idEquipo === id) {
        // Copiar la fila coincidente con las modificaciones ya aplicadas
        filaCoincidente = row.slice();
        break; // Salir del bucle una vez que se ha encontrado la fila coincidente
      }
    }
  
    if (!filaCoincidente) {
      Logger.log("No se encontró ninguna coincidencia para el ID especificado");
      return;
    }
  
    // Pedir al usuario que especifique la hoja de destino
    var hojaDestino = Browser.inputBox("Especifique la hoja de destino: (equipos nuevos, equipos usados, equipos entrega inmediata, equipos reservados)", Browser.Buttons.OK_CANCEL);
  
    var hojaDestinoName;
    
    if (hojaDestino === "equipos nuevos") {
      hojaDestinoName = "equipos nuevos";
      estado = "Nuevo";
    } else if (hojaDestino === "equipos usados") {
      hojaDestinoName = "equipos usados";
      estado = "Usado";
    } else if (hojaDestino === "equipos entrega inmediata") {
      hojaDestinoName = "equipos entrega inmediata";
      estado = "Entrega inmediata";
    } else if (hojaDestino === "equipos reservados") {
      hojaDestinoName = "equipos reservados";
      estado = "Reservado";
    } else {
      Logger.log("Hoja de destino no válida");
      return;
    }
  
    var hojaDestinoSheet = ss.getSheetByName(hojaDestinoName);
    var ultimaFila = hojaDestinoSheet.getRange("B:B").getValues().filter(String).length;
    var ultimoId = hojaDestinoSheet.getRange(ultimaFila, 2).getValue();
    var nuevoId = isNaN(ultimoId) ? 1 : parseInt(ultimoId.replace("ID", "")) + 1;
  
    // Insertar la fila modificada en la hoja de destino con el nuevo ID y el estado correspondiente
    hojaDestinoSheet.getRange(ultimaFila + 1, 1, 1, filaCoincidente.length).setValues([filaCoincidente]);
    hojaDestinoSheet.getRange(ultimaFila + 1, 2).setValue(nuevoId);
    hojaDestinoSheet.getRange(ultimaFila + 1, 8).setValue(estado);
    hojaDestinoSheet.getRange(ultimaFila + 1, 16).setValue("");
    hojaDestinoSheet.getRange(ultimaFila + 1, 17).setValue("");
    hojaDestinoSheet.getRange(ultimaFila + 1, 18).setValue("");
    hojaDestinoSheet.getRange(ultimaFila + 1, 13).setValue("Llegó");
  
    // Eliminar la fila original en la hoja de equipos con problemas
    equipoSheet.deleteRow(i + 1);
  
    // Limpiar contenido en las celdas del menú principal
    var rango1 = mainSheet.getRange("B55:I56");
    var rango2 = mainSheet.getRange("A58:I58");
    rango1.clearContent();
    rango2.clearContent();
  }
  // Menú otros movimientos
  function cargarMovimientos() {
    
    var ss= SpreadsheetApp.getActiveSpreadsheet()
    var menu_principal = ss.getSheetByName('menu principal');
    var movimientos = ss.getSheetByName('movimientos');
  
    // Obtener los datos de las celdas en la hoja "Menu completo"
    var fecha = menu_principal.getRange("B37").getValues();
    var persona = menu_principal.getRange("B38").getValues();
    var anotaciones = menu_principal.getRange("B39").getValues();
    var cuenta = menu_principal.getRange("A41:A49").getValues();
    var monto = menu_principal.getRange("B41:B49").getValues();
    var moneda = menu_principal.getRange("C41:C49").getValues();
    var compras = 'Otros movimientos'
    var responsable = Session.getActiveUser().getEmail();
    
    
  
    // Obtener el último ID de la hoja "Movimientos" y sumar 1 para generar un nuevo ID
    var ultimoId = parseInt(movimientos.getRange("B2").getDisplayValue()) || 0;
    var nuevoId = ultimoId + 1;
    
    // Iterar sobre los datos de las variables para agregar cada fila a la hoja "Movimientos"
    
    for (var i = 0; i < monto.length; i++) {
      for (var j = 0; j < monto[i].length; j++) {
        if (monto[i][j] !== "") { // Verificar si hay un valor en la variable "monto"
              var nuevaFila = [fecha, nuevoId, persona, anotaciones, compras, "", "", "", cuenta[i][j],- monto[i][j], 
              "", - monto[i][j], "", "", moneda[i][j], "", "", responsable];
              movimientos.insertRowAfter(1);
              movimientos.getRange("A2:R2").setValues([nuevaFila]);
              nuevoId++;
            }
          }
        }
    
    // Obtener el número de filas en la hoja "Movimientos" después de agregar las nuevas filas
    var numFilas = movimientos.getLastRow() - 1;
    
    // Establecer el formato para las filas agregadas automáticamente
    var formatoBasico = movimientos.getRange(2,1,numFilas,20).setBackground("#ffffff").setFontFamily("DM Sans").setFontSize(10).setFontColor('#000000').setBorder(true,true,true,true,true,true);
  
    var rango1 = menu_principal.getRange("B37:C39"); 
    var rango2 = menu_principal.getRange("A41:C49");
    rango1.clearContent();
    rango2.clearContent();
  }
  
  
  // Menú Editar
  
  function buscarFilas() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hojaActual = spreadsheet.getActiveSheet();
  
    var id = hojaActual.getRange("B63").getValue();
    var imei = hojaActual.getRange("B64").getValue();
    var modelo = hojaActual.getRange("B65").getValue();
    var estado = hojaActual.getRange("B66").getValue();
  
    var equipoSheetName;
  
    if (estado === "Nuevo") {
      equipoSheetName = "equipos nuevos";
    } else if (estado === "Usado") {
      equipoSheetName = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      equipoSheetName = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      equipoSheetName = "equipos reservados";
    } else if (estado === "Con problemas") {
      equipoSheetName = "equipos con problemas";
    } else {
      // Estado no válido, no se realiza la búsqueda
      return;
    }
  
    var hojaDestino = spreadsheet.getSheetByName(equipoSheetName);
    var datos = hojaDestino.getDataRange().getValues();
    var filasCoincidentes = [];
  
    for (var i = 0; i < datos.length; i++) {
      var fila = datos[i];
      if ((id && fila[1] === id) || (imei && fila[5] === imei) || (modelo && fila[6] === modelo)) {
        filasCoincidentes.push(fila);
      }
    }
  
    if (filasCoincidentes.length > 0) {
      var rangoDestino = hojaActual.getRange("A69:S" + (68 + filasCoincidentes.length));
      rangoDestino.setValues(filasCoincidentes);
    } else {
      // No se encontraron filas coincidentes, mostrar mensaje de error
      Browser.msgBox("No hay equipos para los valores de búsqueda ingresados.");
    }
  }
  
  function actualizarFilasModificadas() {
    var menu_principal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu principal');
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hojaActual = spreadsheet.getActiveSheet();
  
    var estado = hojaActual.getRange("B66").getValue();
  
    var equipoSheetName;
  
    if (estado === "Nuevo") {
      equipoSheetName = "equipos nuevos";
    } else if (estado === "Usado") {
      equipoSheetName = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      equipoSheetName = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      equipoSheetName = "equipos reservados";
    } else if (estado === "Con problemas") {
      equipoSheetName = "equipos con problemas";
    } else {
      // Estado no válido, no se realiza la búsqueda
      return;
    }
  
    var hojaDestino = spreadsheet.getSheetByName(equipoSheetName);
  
    var rangoFilasModificadas = hojaActual.getRange("A69:U").getValues();
    var filasModificadas = rangoFilasModificadas.filter(function (fila) {
      return fila[1] !== ""; // Filtrar filas con valor en la columna B
    });
  
    var idColumna = hojaActual.getRange("B69:B" + (68 + filasModificadas.length)).getValues();
  
    var filasCoincidentes = [];
  
    for (var i = 0; i < filasModificadas.length; i++) {
      var filaModificada = filasModificadas[i];
      var id = idColumna[i][0];
  
      var datosDestino = hojaDestino.getDataRange().getValues();
  
      for (var j = 0; j < datosDestino.length; j++) {
        var filaDestino = datosDestino[j];
  
        if (filaDestino[1] === id) {
          var filaCoincidente = [];
  
          if (esAdministrador()) {
            // Tomar todos los valores de la fila modificada
            filaCoincidente = filaModificada;
          } else {
            // Tomar todos los valores de la fila modificada excepto las columnas I y J
            for (var k = 0; k < filaModificada.length; k++) {
              if (k !== 8 && k !== 9) {
                filaCoincidente.push(filaModificada[k]);
              } else {
                // Mantener los valores originales en las columnas I y J
                filaCoincidente.push(filaDestino[k]);
              }
            }
          }
  
          filasCoincidentes.push({ fila: j, datos: filaCoincidente });
        }
      }
    }
  
    for (var k = 0; k < filasCoincidentes.length; k++) {
      var filaCoincidente = filasCoincidentes[k];
      hojaDestino.getRange(filaCoincidente.fila + 1, 1, 1, filaCoincidente.datos.length).setValues([filaCoincidente.datos]);
    }
  
    var rango1 = menu_principal.getRange("A69:U");
    rango1.clearContent();
  }
  
  function cambioInventario() {
    var menu_principal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu principal');
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var estado = menu_principal.getRange("B66").getValue();
  
    var hojaOriginal;
  
    if (estado === "Nuevo") {
      hojaOriginal = "equipos nuevos";
    } else if (estado === "Usado") {
      hojaOriginal = "equipos usados";
    } else if (estado === "Entrega inmediata") {
      hojaOriginal = "equipos entrega inmediata";
    } else if (estado === "Reservado") {
      hojaOriginal = "equipos reservados";
    } else {
      // Estado no válido, no se realiza la búsqueda
      return;
    }
  
    var hoja_original = spreadsheet.getSheetByName(hojaOriginal);
  
    var rangoFilasModificadas = menu_principal.getRange("A69:U").getValues();
    var filasModificadas = rangoFilasModificadas.filter(function (fila) {
      return fila[1] !== ""; // Filtrar filas con valor en la columna B
    });
  
    Logger.log(filasModificadas) 
    
    var idColumna = menu_principal.getRange("B69:B" + (68 + filasModificadas.length)).getValues();
  
    var filasCoincidentes = [];
  
    var hojaDestino = Browser.inputBox("Especifique la hoja de destino: (nuevos,usados,entrega inmediata,reservados)", Browser.Buttons.OK_CANCEL);
  
    var hojaDestinoName;
    
    if (hojaDestino === "nuevos") {
      hojaDestinoName = "equipos nuevos";
      estado = "Nuevo";
    } else if (hojaDestino === "usados") {
      hojaDestinoName = "equipos usados";
      estado = "Usado";
    } else if (hojaDestino === "entrega inmediata") {
      hojaDestinoName = "equipos entrega inmediata";
      estado = "Entrega inmediata";
    } else if (hojaDestino === "reservados") {
      hojaDestinoName = "equipos reservados";
      estado = "Reservado";
    } else {
      Logger.log("Hoja de destino no válida");
      return;
    }
  
    var hojaDestino = spreadsheet.getSheetByName(hojaDestinoName);
    var ultimaFilaDestino = hojaDestino.getLastRow();
    
    var idUltimaFila = hojaDestino.getRange("B2").getValue();
    var datosOriginales = hoja_original.getDataRange().getValues();
  
      
      // Resto del código...
      
      // Actualiza el valor de idUltimaFila con el nuevo ID calculado
  
  
  
    if (datosOriginales.length > 1) {
      for (var i = 0; i < filasModificadas.length; i++) {
        var filaModificada = filasModificadas[i];
        var id = idColumna[i][0];
        var eliminarFila = false;
        var filaOriginalIndex = -1;
        var nuevoID = idUltimaFila !== "" ? idUltimaFila + 1 : 1;
  
        for (var j = 0; j < datosOriginales.length; j++) {
          var filaOriginal = datosOriginales[j];
          
          if (filaOriginal[1] === id) {
            var filaCoincidente = [];
            var filaOriginalID = filaOriginal[1];
            eliminarFila = true;
            filaOriginalIndex = j;
  
            if (esAdministrador()) {
              // Tomar todos los valores de la fila modificada
              filaCoincidente = filaModificada;
            } else {
              // Tomar todos los valores de la fila modificada excepto las columnas I y J
              for (var k = 0; k < filaModificada.length; k++) {
                if (k !== 8 && k !== 9) {
                  filaCoincidente.push(filaModificada[k]);
                } else {
                  // Mantener los valores originales en las columnas I y J
                  filaCoincidente.push(filaOriginal[k]);
                }
              }
            }
            filaCoincidente[1] = nuevoID
            filasCoincidentes.push({ fila: j, datos: filaCoincidente });
            break; // Salir del bucle al encontrar la primera coincidencia
          }
        }
  
        if (eliminarFila && filaOriginalIndex !== -1) {
          hoja_original.deleteRow(filaOriginalIndex + 1); // Eliminar la fila en la hoja original
        }
      }
    }
  
    for (var k = 0; k < filasCoincidentes.length; k++) {
      var filaCoincidente = filasCoincidentes[k];
      hojaDestino.insertRowBefore(2);
      hojaDestino.getRange(2, 1, 1, filaCoincidente.datos.length).setValues([filaCoincidente.datos]);
    }
  
    var rango1 = menu_principal.getRange("A69:U");
    rango1.clearContent();
  }
  
  
  
  