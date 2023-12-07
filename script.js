function iniciarBusqueda() {
    // Llama a la función para leer el archivo Excel y construir la tabla
    leerArchivoExcel();
  }
  
  // Función para leer el archivo Excel y construir la tabla
  function leerArchivoExcel() {
    // Ruta al archivo Excel
    var archivoExcel = 'informacion.xlsm';
  
    // Lee el archivo Excel
    var reader = new FileReader();
    reader.onload = function (e) {
      var data = new Uint8Array(e.target.result);
      var workbook = XLSX.read(data, { type: 'array' });
  
      // Selecciona la primera hoja del libro de trabajo
      var sheetName = workbook.SheetNames[0];
      var sheet = workbook.Sheets[sheetName];
  
      // Obtiene los datos de la hoja
      var datos = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
      // Llama a la función para construir la tabla con los datos filtrados
      construirTabla(datos);
    };
  
    // Lee el contenido del archivo
    fetch(archivoExcel)
      .then((response) => response.arrayBuffer())
      .then((data) => reader.readAsArrayBuffer(new Blob([data])));
  }
  
  // Función para construir la tabla con los datos filtrados
  function construirTabla(datos) {
    // Selecciona el elemento donde se insertará la tabla
    var tablaContainer = document.getElementById('tabla-container');
  
    // Limpia el contenido existente en el contenedor
    tablaContainer.innerHTML = '';
  
    // Obtiene el número ingresado por el usuario
    var numeroABuscar = document.getElementById('numero-a-buscar').value;
  
    // Filtra los datos para encontrar la fila correspondiente al número ingresado
    var filaEncontrada = datos.find(function (filaDatos) {
      // Ahora supongo que el número está en la columna D (index 3)
      return filaDatos[3] == numeroABuscar; // Cambiado a == para comparación no estricta
    });
  
    // Si se encuentra la fila, construye la tabla con esa fila
    if (filaEncontrada) {
      // Crea la tabla
      var tabla = document.createElement('table');
  
      // Crea la fila de encabezado
      var encabezado = tabla.createTHead();
      var encabezadoFila = encabezado.insertRow();
  
      // Agrega las columnas al encabezado
      ['NO. DE PREPARACION', 'FECHA', 'OFICINA'].forEach(function (encabezadoTexto) {
        var th = document.createElement('th');
        th.textContent = encabezadoTexto;
        encabezadoFila.appendChild(th);
      });
  
      // Crea el cuerpo de la tabla
      var cuerpo = tabla.createTBody();
  
      // Agrega la fila encontrada y celdas con datos
      var fila = cuerpo.insertRow();
      // Mostrar solo las tres primeras columnas
      for (var i = 0; i < 3; i++) {
        var celda = fila.insertCell();
        celda.textContent = filaEncontrada[i];
      }
  
      // Agrega estilos a la tabla
      tabla.classList.add('styled-table');
  
      // Inserta la tabla en el contenedor
      tablaContainer.appendChild(tabla);
  
      // Agrega el mensaje al final de la tabla
      var mensajeInferior = document.createElement('p');
      mensajeInferior.textContent =
        'Su trámite se encuentra en verificación de información, validación y elaboración, etapas necesarias para la expedición de su documento de identidad, proceso normal de producción. Por favor vuelva a realizar la consulta en los próximos días para validar el avance del documento.';
      tablaContainer.appendChild(mensajeInferior);
  
      // Agrega el mensaje sobre la tabla con el número en negrita
      var mensajeSuperior = document.createElement('p');
      mensajeSuperior.innerHTML = 'El documento <strong>' + numeroABuscar + '</strong> presenta las siguientes solicitudes:';
      tablaContainer.insertBefore(mensajeSuperior, tabla);
    } else {
      // Si no se encuentra la fila, muestra un mensaje de error
      var mensajeError = document.createElement('p');
      mensajeError.innerHTML = 'Su solicitud del Documento <strong>' + numeroABuscar + '</strong> no se encuentra en Trámite. Si solicitó el Trámite de su documento y el mismo no se refleja en esta consulta, por favor vuelva a consultar en los próximos días para verificar el estado de su documento.';
      tablaContainer.appendChild(mensajeError);
    }
  }
  
  // Código original
  var linkCobrandingNegropng = document.getElementById("linkCobrandingNegropng");
  if (linkCobrandingNegropng) {
    linkCobrandingNegropng.addEventListener("click", function() {
      window.open("https://www.registraduria.gov.co/");
    });
  }
  
  var linkContainer = document.getElementById("linkContainer");
  if (linkContainer) {
    linkContainer.addEventListener("click", function() {
      window.open("https://www.facebook.com/RegistraduriaNacional/");
    });  
  }
  
  var linkContainer1 = document.getElementById("linkContainer1");
  if (linkContainer1) {
    linkContainer1.addEventListener("click", function() {
      window.open("https://twitter.com/Registraduria");
    });
  }
  
  var linkContainer2 = document.getElementById("linkContainer2");
  if (linkContainer2) {
    linkContainer2.addEventListener("click", function() {
      window.open("https://www.youtube.com/user/RegistraduriaNal");
    });
  }
  
  var linkContainer18 = document.getElementById("linkContainer18");
  if (linkContainer18) {
    linkContainer18.addEventListener("click", function() {
      window.open("https://www.google.com/chrome/browser/desktop/index.html");
    }); 
  }