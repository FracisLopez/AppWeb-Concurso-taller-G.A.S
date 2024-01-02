// VARIABLES GLOBALES

//Url de googleSheet
var SS = SpreadsheetApp.openById('ID_DE_TU_HOJA_CALCULO');

//Nombre de la Hoja 
var sheetConfiguracion = SS.getSheetByName('Configuracion');

//Nombre de la Hoja 
var sheetBdConfiguracion = SS.getSheetByName('BdConfiguracion');

//Nombre de la Hoja 
var sheetUsuarios = SS.getSheetByName('Usuarios');

//Nombre de la Hoja 
var sheetBdRegistro = SS.getSheetByName('BdRegistro');

//Nombre de la Hoja 
var sheetAccesos = SS.getSheetByName('Accesos');

function doGet() {

  const permitirAcceso = searchUser();

  if (permitirAcceso === true) {

    // Aqui esta parte es crear los campos del formulario de manera dinamica
    var data = sheetUsuarios.getDataRange().getDisplayValues();

    console.log(data);
    var template = HtmlService.createTemplateFromFile('Index');
    //Permite acceder al arreglo data desde la pagina web
    template.data = data;

    var output = template.evaluate();
    //Etiquetas meta para hacer tu web responsive
    output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return output;

  } else {

    const salida = HtmlService.createHtmlOutput("<h2>Acceso no permitido</h2><br><h3>Asegurate loggearte con tu cuenta de correo.</h3>");
    salida.addMetaTag('viewport', 'width=device-width, initial-scale=1');

    return salida;
  }// fin else


} // fin de la funcion doGet()


// Funcion que me va permitir incluir los archvis ccs y js separados
function include(fileName) {

  return HtmlService.createHtmlOutputFromFile(fileName)
    .getContent();
}


function verificarPasswordLogin(txt_usuario, txt_password, rol) {

  var dataUsuarios = sheetUsuarios.getDataRange().getValues();
  //var inicio_sesion;
  var resultadoLogin = 'LoginIncorrecto';
  var resultadotipoPerfil;
  var resultadoNombreUsuario;

  for (var i in dataUsuarios) {

    if (dataUsuarios[i][1] == txt_usuario && dataUsuarios[i][2] == txt_password && dataUsuarios[i][5] == rol) {

      if (rol == 'Administrador') { // Si es Administrador y tiene las credenciales correctas

        resultadoLogin = 'LoginOk'
        resultadotipoPerfil = 'Administrador';
        resultadoNombreUsuario = dataUsuarios[i][0];
      } else {// Si es usuario y tiene las credenciales correctas

        resultadoLogin = 'LoginOk'
        resultadotipoPerfil = 'Usuario';
        resultadoNombreUsuario = dataUsuarios[i][0];

      }// Fin else si es Usuario 

    }// fin if   
  }// fin for

  var resultado = [resultadoLogin, resultadotipoPerfil, resultadoNombreUsuario];
  return resultado;

}// verificarPasswordLogin

function gsagregarAbrircerrarConcurso(valor) {

  sheetConfiguracion.getRange('B1').setValue(valor);
  return valor;

}// fin

function gsleerAbrircerrarConcurso() {

  var valor = sheetConfiguracion.getRange('B1').getValue();
  //console.log(valor);
  return valor;
}// Fin

function gsGuardarFechaHora(txtConFechaini, timeHoraini, txtConFechafin, timeHorafin, txtobjetivo) {

  sheetConfiguracion.getRange('B2').setValue(txtConFechaini);
  sheetConfiguracion.getRange('B3').setValue(timeHoraini);
  sheetConfiguracion.getRange('B4').setValue(txtConFechafin);
  sheetConfiguracion.getRange('B5').setValue(timeHorafin);
  sheetConfiguracion.getRange('B6').setValue(txtobjetivo);

  var valor = txtConFechaini + " " + timeHoraini + " " + txtConFechafin + " " + timeHorafin;

  return valor;

}// fin


function gsleerFechaHora() {

  if (sheetConfiguracion.getRange('B2').getValue() == "") {
    var fechaini = sheetConfiguracion.getRange('B2').getValue();
  } else {
    var fechaini = Utilities.formatDate(sheetConfiguracion.getRange('B2').getValue(), "GMT", "yyyy-MM-dd");
  }

  var horaini = sheetConfiguracion.getRange('B3').getDisplayValue();

  if (sheetConfiguracion.getRange('B4').getValue() == "") {
    var fechafin = sheetConfiguracion.getRange('B4').getValue();
  } else {
    var fechafin = Utilities.formatDate(sheetConfiguracion.getRange('B4').getValue(), "GMT", "yyyy-MM-dd");
  }

  var horafin = sheetConfiguracion.getRange('B5').getDisplayValue();
  var objetivoVentas = sheetConfiguracion.getRange('B6').getValue();

  let resultado

  resultado = [fechaini, agregarceroAhora(horaini), fechafin, agregarceroAhora(horafin), objetivoVentas];

  //console.log(resultado);
  return resultado;

}// Fin


function agregarceroAhora(horaAnadir) {

  // funcion que agrega un cero adelante si la hora esta entre 0 y el 9 
  var hora = horaAnadir; // Tu hora actual

  // Separar la hora y los minutos
  var partes = hora.split(":");
  var horas = partes[0];
  var minutos = partes[1];

  // Verificar si las horas están en el rango de 0 a 9
  if (horas >= 0 && horas <= 9) {
    horas = horas.padStart(2, "0"); // Añadir cero adelante de las horas
  }

  // Crear la hora formateada con cero adelante si es necesario
  var horaFormateada = horas + ":" + minutos;

  //console.log(horaFormateada); // Resultado: "01:07" (si la hora original era "1:07")
  return horaFormateada

}// fin agregarceroAhora

function gsconfigAgregar(datos) {

  sheetBdConfiguracion.appendRow(datos);

}


function gsConfigobtenerFilas(filtro) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("BdConfiguracion");
  var datos = hoja.getRange("A:B").getValues();

  var filasFiltradas = [];

  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] == filtro) {
      filasFiltradas.push(datos[i]);
    }
  }

  // Ordenar filas filtradas de manera ascendente
  filasFiltradas.sort(function (a, b) {
    return a[1].toString().localeCompare(b[1].toString());
  });

  return filasFiltradas;
}// FIN gsobtenerFilasConfig


function gsConfigEliminarFila(filaSeleccionada) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("BdConfiguracion");
  var datos = hoja.getRange("A:B").getValues();

  for (var i = datos.length - 1; i >= 0; i--) {
    if (datos[i][1] == filaSeleccionada) {
      hoja.deleteRow(i + 1);
    }
  }
}// fin gsConfigEliminarFila



function gsCargarSelect() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BdConfiguracion");
  var data = sheet.getDataRange().getValues();

  //for (var i = 0; i < data.length; i++) {
  //console.log(data[i]);
  //}
  //console.log(data);
  return data;
}

function gsAgregarVentar(datos) {

  
  sheetBdRegistro.appendRow(datos);

}

function gslistaVentas(rol, txt_usuario) {
  var sheetBdRegistro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BdRegistro");
  var ultimaFila = sheetBdRegistro.getLastRow();
  var registros = [];
  var rol_bd = rol; // Variable ROL con valor "Administrador"
  var usuario = txt_usuario; // Variable txtUsuario con valor "1"

  for (var fila = 2; fila <= ultimaFila; fila++) {
    var filaDatos = [];
    var columA = sheetBdRegistro.getRange(fila, 1).getValue();

    if (sheetBdRegistro.getRange(fila, 2).getValue() == "") {
      var columB = sheetBdRegistro.getRange(fila, 2).getValue();
    } else {
      var columB = Utilities.formatDate(sheetBdRegistro.getRange(fila, 2).getValue(), "GMT", "dd/MM/yyyy")
    }

    if (sheetBdRegistro.getRange(fila, 3).getValue() == "") {
      var columC = sheetBdRegistro.getRange(fila, 3).getValue();
    } else {
      var columC = sheetBdRegistro.getRange(fila, 3).getDisplayValue();
    }

    var columD = sheetBdRegistro.getRange(fila, 4).getValue();
    var columE = sheetBdRegistro.getRange(fila, 5).getValue();
    var columF = sheetBdRegistro.getRange(fila, 6).getValue();
    var columG = sheetBdRegistro.getRange(fila, 7).getValue();
    var columH = sheetBdRegistro.getRange(fila, 8).getValue();
    var columI = sheetBdRegistro.getRange(fila, 9).getValue();
    var columJ = sheetBdRegistro.getRange(fila, 10).getValue();
    var columK = sheetBdRegistro.getRange(fila, 11).getValue();
    var columL = sheetBdRegistro.getRange(fila, 12).getValue();

    if (rol_bd == "Administrador") {
      filaDatos.push(columA, columB, columC, columD, columE, columF, columG, columH, columI, columJ, columK, columL);
      registros.push(filaDatos);
    } else if (rol_bd == "Usuario" && columD == usuario) { // Filtrar por usuario (comparar con columna D)
      filaDatos.push(columA, columB, columC, columD, columE, columF, columG, columH, columI, columJ, columK, columL);
      registros.push(filaDatos);
    }
  }


  // Ordenar los registros por fecha y hora (columnas B y C) en orden descendente
  registros.sort(function (a, b) {
    var fechaHoraA = new Date(a[1].split('/').reverse().join('/') + ' ' + a[2]);
    var fechaHoraB = new Date(b[1].split('/').reverse().join('/') + ' ' + b[2]);
    return fechaHoraB - fechaHoraA;
  });

  //console.log(registros);
  return registros
}

function gs_antesdeEliminarVenta(variable) {

  // Valida si esta abierto o cerrado el concursos dependiendo de eso se podra registrar
  var valorCelda = sheetConfiguracion.getRange("B1").getValue();
  //console.log(valorCelda);
  if (valorCelda == 0) {
    return valorCelda;
  } else {
    return variable;
  }

}//gs_antesdeEliminarVenta





function gs_EliminarVenta(variable) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("BdRegistro");
  var datos = hoja.getRange("A:A").getValues(); // Obtener los valores de la columna A
  var filaAEliminar = -1; // Inicializar la variable de la fila a eliminar
  var resultado

  // Recorrer los valores de la columna A y buscar la fila que coincida con la variable
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] == variable) {
      filaAEliminar = i + 1; // Guardar la fila a eliminar (se suma 1 porque los índices de las filas comienzan en 1)
      break; // Salir del bucle al encontrar la fila
    }
  }


  // Verificar si se encontró la fila a eliminar
  if (filaAEliminar > -1) {
    hoja.deleteRow(filaAEliminar); // Eliminar la fila encontrada
    resultado = "Registro Eliminado: " + variable
  } else {
    resultado = "No se pudo eliminar registro";
  }

  //return resultado;
  console.log(resultado);

}// fin eliminar



function gsbuscarRegistro(codigo) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("BdRegistro");
  var ultimaFila = hoja.getLastRow();
  var codigos = hoja.getRange("A2:A" + ultimaFila).getValues();
  var registros = [];

  for (var i = 0; i < codigos.length; i++) {
    if (codigos[i][0] === codigo) {
      var rangoFila = hoja.getRange(i + 2, 1, 1, hoja.getLastColumn());
      registros = rangoFila.getValues()[0];
      break;
    }
  }

  var registroventa = [registros[0], registros[5], registros[6], registros[7], registros[8], registros[9], registros[10], registros[11]]
  //console.log(registroventa);
  return registroventa;
} // Fin buscar registro
                             

function gsEditarVentar(key, nuevoValorE, nuevoValorF, nuevoValorG, nuevoValorH, nuevoValorI, nuevoValorJ, nuevoValorK) {

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BdRegistro");
  var data = hoja.getDataRange().getValues();

  // Buscar la fila que coincide con la variable key en la columna A
  var filaEncontrada = -1;
  for (var fila = 0; fila < data.length; fila++) {
    if (data[fila][0] === key) {
      filaEncontrada = fila;
      break;
    }
  }

  if (filaEncontrada !== -1) {
    // Actualizar las columnas E, F, G, H. I, J y K de la fila encontrada
    hoja.getRange(filaEncontrada + 1, 6).setValue(nuevoValorE); // Columna E
    hoja.getRange(filaEncontrada + 1, 7).setValue(nuevoValorF); // Columna F
    hoja.getRange(filaEncontrada + 1, 8).setValue(nuevoValorG); // Columna G
    hoja.getRange(filaEncontrada + 1, 9).setValue(nuevoValorH); // Columna H
    hoja.getRange(filaEncontrada + 1, 10).setValue(nuevoValorI); // Columna I
    hoja.getRange(filaEncontrada + 1, 11).setValue(nuevoValorJ); // Columna J
    hoja.getRange(filaEncontrada + 1, 12).setValue(nuevoValorK); // Columna K
    

    //Logger.log("Fila actualizada exitosamente.");
  } else {
    //Logger.log("No se encontró la fila con la clave '" + key + "'.");
  }
}// fin editar


function gsDashleerFechaHoravigencia() {

  if (sheetConfiguracion.getRange('B2').getValue() == "") {
    var fechaini = sheetConfiguracion.getRange('B2').getValue();
  } else {
    var fechaini = Utilities.formatDate(sheetConfiguracion.getRange('B2').getValue(), "GMT", "yyyy-MM-dd");
  }

  var horaini = sheetConfiguracion.getRange('B3').getDisplayValue();

  if (sheetConfiguracion.getRange('B4').getValue() == "") {
    var fechafin = sheetConfiguracion.getRange('B4').getValue();
  } else {
    var fechafin = Utilities.formatDate(sheetConfiguracion.getRange('B4').getValue(), "GMT", "yyyy-MM-dd");
  }

  var horafin = sheetConfiguracion.getRange('B5').getDisplayValue();

  var resultado

  //resultado=[fechaini,agregarceroAhora(horaini),fechafin,agregarceroAhora(horafin)];
  resultado = 'Concurso vigente desde el ' + fechaini + ' ' + agregarceroAhora(horaini) + ' hasta el ' + fechafin + ' ' + agregarceroAhora(horafin);

  //console.log(resultado);
  return resultado;

}// Fin gsDashleerFechaHoravigencia


function gsDashrankingVentas() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("BdRegistro");
  var data = hoja.getDataRange().getValues();
  var usuarios = {};

  for (var i = 1; i < data.length; i++) {
    var usuario = data[i][4];
    var piezas = data[i][11];

    if (!usuarios[usuario]) {
      usuarios[usuario] = piezas;
    } else {
      usuarios[usuario] += piezas;
    }
  }

  var resultado = [];

  for (var usuario in usuarios) {
    resultado.push([usuario, usuarios[usuario]]);
  }

  // Ordenar por cantidad de manera descendente
  resultado.sort(function (a, b) {
    return b[1] - a[1];
  });

  // Añadir el contador
  for (var i = 0; i < resultado.length; i++) {
    resultado[i].unshift(i + 1);
  }

  //Cantidad de item tiene el arreglo 
  var piezasElementos = resultado.length;
  //console.log(cantidadElementos);

  // funcion valida la cantidad de elementos esto evita el error cuando aun no han hecho ninguna venta o solo hay un vendedor  o 2 de tal forma en el ranking de los 3 primeros puestos se coloca un - para aquellas posiciones que aun no existen 
  if (piezasElementos === 0) {

    resultado.push(["-", "-"]);
    resultado.push(["-", "-"]);
    resultado.push(["-", "-"]);
    //console.log(resultado[0][1],resultado[1][1],resultado[2][1]);
    return resultado;
  }
  if (piezasElementos === 1) {

    resultado.push(["-", "-"]);
    resultado.push(["-", "-"]);
    //console.log(resultado[0][1],resultado[1][1],resultado[2][1]);
    return resultado;
  }
  if (piezasElementos === 2) {

    resultado.push(["-", "-"]);
    console.log(resultado[0][1], resultado[1][1], resultado[2][1]);
    return resultado;
  } else {// si es mas de 2 es decir 3

    //console.log(resultado[0][1],resultado[1][1],resultado[2][1]);  
    return resultado;
  }

  //console.log(resultado[0][1],resultado[1][1],resultado[2][1]);
  //console.log(resultado[0][1]);
  //return resultado;
  //return resultado;

}//fin funcion gsDashrankingVentas


function gsDashtotales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("BdRegistro");
  var ultimaFila = hoja.getLastRow();
  var piezasTotal = 0;

  for (var fila = 2; fila <= ultimaFila; fila++) {
    var piezas = hoja.getRange(fila, 12).getValue();
    piezasTotal += piezas;
  }

  // Imprimir resultado en la consola
  //console.log("Suma de Cantidades: " + cantidadTotal);
  if (sheetConfiguracion.getRange('B6').getValue() == '') {
    var objetivopiezas = 'Sin objetivo';
    var gap = 'Falta objetivo';
    var avance = 0;
  } else {
    var objetivoPiezas = sheetConfiguracion.getRange('B6').getValue();
    var avance = (piezasTotal / objetivoPiezas) * 100;
    var gap = piezasTotal - objetivoPiezas
  }



  let resultados;
  resultados = [objetivoPiezas, piezasTotal, gap, avance];
  console.log(resultados);


  return resultados;

}// fin sumarCantidadVentas



function gsantesdeAgregarAbiertoCerrado() {

  // Valida si esta abierto o cerrado el concursos dependiendo de eso se podra registrar
  var valorCelda = sheetConfiguracion.getRange("B1").getValue();
  //console.log(valorCelda);
  return valorCelda;

}// fin gsantesdeAgregarAbiertoCerrad;



//validacion de acceso
function searchUser() {
  const activeUser = Session.getActiveUser().getEmail();
  //const SS = SpreadsheetApp.getActiveSpreadsheet();
  //const sheetUsers = SS.getSheetByName('Accesos');

  // Valida si la fila A tiene algun valor o un correo ingresado
  var valoresColumnaA = sheetAccesos.getRange("A2:A").getValues();
  var filasConValor = 0;

  for (var i = 0; i < valoresColumnaA.length; i++) {
    if (valoresColumnaA[i][0] !== "") {
      filasConValor++;
    }
  }

  if (filasConValor == 0) {//  si no hay ningun valor da true para que todos tengan acceso

    return true;

  } else { // en caso tenga un valor debe validar las cuentas de correo ingresadas

    const activeUsersList = sheetAccesos.getRange(2, 1, sheetAccesos.getLastRow() - 1, 1).getValues().map(user => user[0]);
    // console.log(activeUsersList);

    if (activeUsersList.indexOf(activeUser) !== -1) {
      //console.log('Dar acceso');
      return true;
    } else {
      //console.log('No dar acceso');
      return false;
    }

  }// fin else

}// Fin funciion searchUser

