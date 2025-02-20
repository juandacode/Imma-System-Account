// ACCOUNT SYSTEM IMMA COMPANY

// START CUSTOMER MODULE

function registrarCliente(cedula, nombre, celular, correo, esProveedor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientes = ss.getSheetByName("USUARIOS"); // Asegúrate de tener una hoja llamada 'Clientes'

  if (!hojaClientes) {
    SpreadsheetApp.getUi().alert("La hoja 'USUARIOS' no existe.");
    return;
  }

  const ultimaFila = hojaClientes.getLastRow() + 1;

  hojaClientes.getRange(ultimaFila, 1).setValue(cedula); // Cédula
  hojaClientes.getRange(ultimaFila, 2).setValue(nombre); // Nombre
  hojaClientes.getRange(ultimaFila, 3).setValue(celular); // Celular
  hojaClientes.getRange(ultimaFila, 4).setValue(correo); // Correo
  hojaClientes.getRange(ultimaFila, 5).setValue(esProveedor ? "PROVEEDOR" : "CLIENTE"); // Tipo
}

function mostrarFormularioCliente() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 20px;
        }
        h2 {
          text-align: center;
          color: #333;
        }
        form {
          display: flex;
          flex-direction: column;
          align-items: center;
        }
        label {
          font-weight: bold;
          margin: 10px 0 5px;
        }
        input[type="text"],
        input[type="email"],
        input[type="number"] {
          width: 90%;
          padding: 10px;
          margin-bottom: 20px;
          border: 1px solid #ccc;
          border-radius: 5px;
          font-size: 16px;
        }
        input[type="checkbox"] {
          transform: scale(1.5);
          margin: 10px;
        }
        button {
          width: 100px;
          padding: 10px;
          border: none;
          border-radius: 5px;
          font-size: 16px;
          color: white;
          cursor: pointer;
          margin: 10px;
        }
        button#registrar {
          background-color: #4CAF50; /* Verde */
        }
        button#cancelar {
          background-color: #f44336; /* Rojo */
        }
        button:hover {
          opacity: 0.9;
        }
        .notification {
          position: fixed;
          bottom: 20px;
          left: 50%;
          transform: translateX(-50%);
          background-color: #4CAF50;
          color: white;
          padding: 15px;
          border-radius: 5px;
          display: none;
        }
      </style>
    </head>
    <body>
      <h2>Registrar Cliente</h2>
      <form id="clienteForm">
        <label for="cedula">Cédula:</label>
        <input type="text" id="cedula" name="cedula" required>

        <label for="nombre">Nombre:</label>
        <input type="text" id="nombre" name="nombre" required>

        <label for="celular">Celular:</label>
        <input type="number" id="celular" name="celular" required>

        <label for="correo">Correo:</label>
        <input type="email" id="correo" name="correo" required>

        <label>
          <input type="checkbox" id="esProveedor">PROVEEDOR
          <input type="checkbox" id="esProveedor">CLIENTE
        </label>

        <div>
          <button type="button" id="registrar" onclick="registrarCliente()">Registrar</button>
          <button type="button" id="cancelar" onclick="google.script.host.close()">Cancelar</button>
        </div>
      </form>

      <div class="notification" id="notification">Cliente registrado correctamente.</div>

      <script>
        function mostrarNotificacion() {
          const notification = document.getElementById("notification");
          notification.style.display = "block";
          setTimeout(() => {
            notification.style.display = "none";
            google.script.host.close();
          }, 2000);
        }

        function registrarCliente() {
          const cedula = document.getElementById("cedula").value;
          const nombre = document.getElementById("nombre").value;
          const celular = document.getElementById("celular").value;
          const correo = document.getElementById("correo").value;
          const esProveedor = document.getElementById("esProveedor").checked;

          if (!cedula || !nombre || !celular || !correo) {
            alert("Por favor, complete todos los campos.");
            return;
          }

          google.script.run
            .withSuccessHandler(() => {
              mostrarNotificacion();
            })
            .registrarCliente(cedula, nombre, celular, correo, esProveedor);
        }
      </script>
    </body>
  </html>`;
  
  const ui = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(ui, "Registro de Cliente");
}

/*function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sistema')
    .addItem('Registrar Cliente', 'mostrarFormularioClientes')
    .addToUi();
}*/
/*
function mostrarFormularioClientes() {
  const htmlFormulario = `
  <style>
    body {
      font-family: Arial, sans-serif;
      text-align: center;
    }
    .form-container {
      max-width: 400px;
      margin: 0 auto;
      padding: 20px;
      border: 2px solid #ccc;
      border-radius: 10px;
      box-shadow: 0 4px 8px rgba(0,0,0,0.2);
      background-color: #f9f9f9;
    }
    h2 {
      color: #333;
    }
    input, select {
      width: 100%;
      padding: 10px;
      margin: 10px 0;
      border: 1px solid #ccc;
      border-radius: 5px;
      box-sizing: border-box;
    }
    .btn-container {
      display: flex;
      justify-content: space-between;
    }
    button {
      padding: 10px 20px;
      border: none;
      border-radius: 5px;
      cursor: pointer;
      font-weight: bold;
    }
    .btn-guardar {
      background-color: #4CAF50;
      color: white;
    }
    .btn-guardar:hover {
      background-color: #45a049;
    }
    .btn-cancelar {
      background-color: #f44336;
      color: white;
    }
    .btn-cancelar:hover {
      background-color: #d32f2f;
    }
  </style>
  <div class="form-container">
    <h2>Registrar Cliente</h2>
    <form id="formCliente">
      <input type="text" id="nombre" placeholder="Nombre" required />
      <input type="text" id="direccion" placeholder="Dirección" />
      <input type="text" id="telefono" placeholder="Teléfono" />
      <input type="email" id="correo" placeholder="Correo Electrónico" />
      <div class="btn-container">
        <button type="button" class="btn-guardar" onclick="guardarCliente()">Guardar</button>
        <button type="button" class="btn-cancelar" onclick="google.script.host.close()">Cancelar</button>
      </div>
    </form>
  </div>
  <script>
    function guardarCliente() {
      const nombre = document.getElementById("nombre").value.trim();
      const direccion = document.getElementById("direccion").value.trim();
      const telefono = document.getElementById("telefono").value.trim();
      const correo = document.getElementById("correo").value.trim();

      if (!nombre) {
        alert("Por favor, completa el nombre del cliente.");
        return;
      }

      google.script.run.registrarCliente(nombre, direccion, telefono, correo);
      alert("Cliente registrado correctamente.");
      google.script.host.close();
    }
  </script>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlFormulario)
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Formulario de Registro de Clientes");
}


/*function mostrarFormularioClientes() {
  const html = HtmlService.createHtmlOutputFromFile('FormularioClientes')
    .setWidth(600)
    .setHeight(500)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Cliente');
}


function registrarCliente(nombre, direccion, telefono, correo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientes = ss.getSheetByName("Clientes"); // Asegúrate de tener una hoja llamada 'Clientes'

  if (!hojaClientes) {
    SpreadsheetApp.getUi().alert("La hoja 'Clientes' no existe.");
    return;
  }

  const ultimaFila = hojaClientes.getLastRow() + 1;

  //hojaClientes.getRange(ultimaFila, 1).setValue(new Date()); // Fecha de registro en la primera columna
  hojaClientes.getRange(ultimaFila, 1).setValue(nombre); // Nombre
  hojaClientes.getRange(ultimaFila, 2).setValue(direccion); // Dirección
  hojaClientes.getRange(ultimaFila, 3).setValue(telefono); // Teléfono
  hojaClientes.getRange(ultimaFila, 4).setValue(correo); // Correo
}
*/

// END CUSTOMER MODULE

 

// START INVENTORY MODULE

function registrarFactura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = ss.getSheetByName("Factura");
  const hojaInventario = ss.getSheetByName("Inventario");

  const filaInicioFactura = 15; // Inicio de productos en Factura
  const filaInicioInventario = 2; // Inicio de productos en Inventario

  const ultimaFilaFactura = hojaFactura.getLastRow();
  const rangoFactura = hojaFactura.getRange(filaInicioFactura, 1, ultimaFilaFactura - filaInicioFactura + 1, 3);
  const datosFactura = rangoFactura.getValues();

  const datosInventario = hojaInventario.getRange(filaInicioInventario, 1, hojaInventario.getLastRow() - filaInicioInventario + 1, 7).getValues();

  let errorStock = false;
  let mensajeErrores = "";

  // Validar el stock antes de registrar la factura
  datosFactura.forEach((filaFactura, index) => {
    const codigo = String(filaFactura[0]).trim();
    const unidades = filaFactura[2]; // Unidades vendidas

    if (codigo && unidades > 0) {
      const indiceProductoInventario = datosInventario.findIndex(filaInventario => String(filaInventario[0]).trim() === codigo);

      if (indiceProductoInventario !== -1) {
        const filaInventario = datosInventario[indiceProductoInventario];
        const stockActual = filaInventario[5]; // Stock actual (Columna F en Inventario)

        if (stockActual - unidades < 0) {
          errorStock = true;
          mensajeErrores += `Producto con código "${codigo}" no tiene suficiente stock para realizar la venta (Stock actual: ${stockActual}, Unidades requeridas: ${unidades}).\n`;
        }
      } else {
        mensajeErrores += `Producto con código "${codigo}" no existe en el inventario.\n`;
        errorStock = true;
      }
    }
  });

  if (errorStock) {
    SpreadsheetApp.getUi().alert("Factura no registrada:\n" + mensajeErrores);
    return;
  }

  // Registrar la factura y actualizar el inventario
  datosFactura.forEach(filaFactura => {
    const codigo = String(filaFactura[0]).trim();
    const unidades = filaFactura[2]; // Unidades vendidas

    if (codigo && unidades > 0) {
      const indiceProductoInventario = datosInventario.findIndex(filaInventario => String(filaInventario[0]).trim() === codigo);

      if (indiceProductoInventario !== -1) {
        const filaInventario = datosInventario[indiceProductoInventario];

        // Actualizar las salidas
        filaInventario[4] += unidades; // Incrementar salidas (Columna E en Inventario)

        // Recalcular el stock actual
        filaInventario[5] = filaInventario[3] - filaInventario[4]; // Stock actual = Entradas - Salidas
      }
    }
  });

  // Guardar los cambios en el inventario
  hojaInventario.getRange(filaInicioInventario, 1, datosInventario.length, 7).setValues(datosInventario);

  SpreadsheetApp.getUi().alert("Factura registrada correctamente.");
}

function registrarCompra() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCompra = ss.getSheetByName("Compra");
  const hojaInventario = ss.getSheetByName("Inventario");

  const filaInicioCompra = 15; // Inicio de productos en Compra
  const filaInicioInventario = 2; // Inicio de productos en Inventario

  const ultimaFilaCompra = hojaCompra.getLastRow();
  const rangoCompra = hojaCompra.getRange(filaInicioCompra, 1, ultimaFilaCompra - filaInicioCompra + 1, 5);
  const datosCompra = rangoCompra.getValues();

  const datosInventario = hojaInventario.getRange(filaInicioInventario, 1, hojaInventario.getLastRow() - filaInicioInventario + 1, 7).getValues();

  datosCompra.forEach((filaCompra, index) => {
    const codigo = String(filaCompra[0]).trim(); // Código del producto (Columna A en Compra)
    const entradas = filaCompra[2]; // Entradas compradas (Columna C en Compra)
    const costoUnitario = filaCompra[3]; // Costo unitario (Columna D en Compra)

    if (codigo && entradas > 0 && costoUnitario > 0) {
      const indiceProductoInventario = datosInventario.findIndex(filaInventario => String(filaInventario[0]).trim() === codigo);

      if (indiceProductoInventario !== -1) {
        const filaInventario = datosInventario[indiceProductoInventario];
        const stockActual = filaInventario[5]; // Stock actual (Columna F en Inventario)
        const costoPromedioActual = filaInventario[6]; // Costo promedio actual (Columna G en Inventario)

        let nuevoCostoPromedio = costoPromedioActual;
        if (costoPromedioActual !== costoUnitario) {
          const unidadesTotales = stockActual + entradas; 
          nuevoCostoPromedio = ((stockActual * costoPromedioActual) + (entradas * costoUnitario)) / unidadesTotales;
        }

        filaInventario[3] += entradas; // Actualizar entradas (Columna D en Inventario)
        filaInventario[5] = filaInventario[3] - filaInventario[4]; // Actualizar stock actual
        filaInventario[6] = nuevoCostoPromedio; // Actualizar costo promedio
      } else {
        SpreadsheetApp.getUi().alert(`Fila ${index + filaInicioCompra}: El producto con código "${codigo}" no existe en el inventario.`);
      }
    }
  });

  hojaInventario.getRange(filaInicioInventario, 1, datosInventario.length, 7).setValues(datosInventario);

  SpreadsheetApp.getUi().alert("Compra registrada correctamente.");
}

// END INVENTORY MODULE

/*
// Incrementa automáticamente la celda en fila 3, columna 3
  if (hojasValidas.includes(hojaActiva.getName())) {
    var contadorCelda = hojaActiva.getRange(3, 3);
    var contadorActual = contadorCelda.getValue();

    if (!contadorActual || isNaN(contadorActual)) {
      contadorActual = 0; // Inicializa si la celda está vacía o no es un número
    }

    contadorCelda.setValue(contadorActual + 1); // Incrementa el valor
  }
  */

// START CODE ID AUTOCOMPLETE MODULE
  function onEdit(e) {
  var hojaActiva = e.source.getActiveSheet();
  var hojaInventarioNombre = "Inventario";
  var hojasValidas = ["Factura", "Compra"];
  var rangoEditado = e.range;
  var columnaEditada = rangoEditado.getColumn();
  var filaEditada = rangoEditado.getRow();

  // Verifica si estamos en una hoja válida y si editamos la columna del Código
  if (hojasValidas.includes(hojaActiva.getName()) && columnaEditada === 1 && filaEditada > 1) {
    var hojaInventario = e.source.getSheetByName(hojaInventarioNombre);
    var inventarioData = hojaInventario.getRange(2, 1, hojaInventario.getLastRow() - 1, 3).getValues();

    var codigoIngresado = rangoEditado.getValue().toString().trim(); // Limpia espacios y convierte a texto
    var nombreCelda = rangoEditado.offset(0, 1);
    var precioCelda = rangoEditado.offset(0, 3);

    // Busca el código en el inventario
    var encontrado = false;
    for (var i = 0; i < inventarioData.length; i++) {
      var codigoInventario = inventarioData[i][0].toString().trim(); // Limpia código en inventario
      if (codigoInventario === codigoIngresado) {
        nombreCelda.setValue(inventarioData[i][1]); // Asigna el Nombre
        if (hojaActiva.getName() === "Factura") {
          precioCelda.setValue(inventarioData[i][2]); // Asigna el Precio Venta
        }
        encontrado = true;
        break;
      }
    }

    // Si no se encuentra el código, limpia las celdas relacionadas
    if (!encontrado) {
      nombreCelda.setValue(""); // Limpia el Nombre
      if (hojaActiva.getName() === "Factura") {
        precioCelda.setValue(""); // Limpia el Precio en Factura
      }
      SpreadsheetApp.getUi().alert("El código ingresado no existe en el inventario. Por favor, verifica.");
    }
  }
}

// END CODE ID AUTOCOMPLETE MODULE


// PROTECT INVENTORY SHEET
function protegerInventario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInventario = ss.getSheetByName("Inventario");

  // Proteger toda la hoja
  const proteccion = hojaInventario.protect();
  proteccion.setDescription("Proteger hoja Inventario");

  // Solo el propietario podrá editar
  const propietario = Session.getEffectiveUser();
  proteccion.addEditor(propietario);
  proteccion.removeEditors(proteccion.getEditors());
  proteccion.setWarningOnly(false); // No permitir ediciones no autorizadas
}
  
  
  
  
/*function confirmarCompraVenta(hoja) {
  var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja);
  
  if (!hojaActiva) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja especificada.");
    return;
  }

  // Obtén la celda del contador (fila 3, columna 3)
  var contadorCelda = hojaActiva.getRange(3, 3);
  var contadorActual = contadorCelda.getValue();

  if (!contadorActual || isNaN(contadorActual)) {
    contadorActual = 0; // Inicializa si la celda está vacía o no es un número
  }

  // Incrementa el contador
  contadorCelda.setValue(contadorActual + 1);

  // Muestra un mensaje de confirmación
  SpreadsheetApp.getUi().alert("¡Factura realizada exitosamente en la hoja " + hoja + "!");
}

confirmarCompraVenta(hojaCompra)*/


// CAJA GENERAL VENTAS

function generarIDEnHoja(nombreHoja) {
  // Abrimos la hoja de cálculo activa
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Accedemos a la hoja especificada por el parámetro nombreHoja
  const hoja = ss.getSheetByName(nombreHoja);
  
  // Verificamos si la hoja existe
  if (!hoja) {
    SpreadsheetApp.getUi().alert(`No se encontró la hoja llamada '${nombreHoja}'.`);
    return;
  }

  // Leemos el valor actual en la celda C3
  const celdaID = hoja.getRange("C3");
  const idActual = celdaID.getValue();

  // Verificamos si la celda tiene un número; si no, empezamos desde 1
  let nuevoID;
  if (typeof idActual === "number" && idActual > 0) {
    nuevoID = idActual + 1; // Si ya hay un ID, le sumamos 1
  } else {
    nuevoID = 1; // Si no hay ID, empezamos desde 1
  }

  // Guardamos el nuevo ID en la celda C3
  celdaID.setValue(nuevoID);

  // Mostramos un mensaje de éxito
  SpreadsheetApp.getUi().alert(`Nuevo ID generado en la hoja '${nombreHoja}': ${nuevoID}`);
}


/*
function registrarVenta() {
  const ssOrigen = SpreadsheetApp.getActiveSpreadsheet(); // Hoja actual
  const hojaFactura = ssOrigen.getSheetByName("Factura"); // Hoja donde se registran las facturas

  // Obtener los datos de la factura
  const fecha = hojaFactura.getRange("C5").getValue();
  const concepto = hojaFactura.getRange("B3").getValue() + " " + hojaFactura.getRange("C3").getValue();
  const cliente = hojaFactura.getRange("B9").getValue(); // Cliente
  const totalVenta = hojaFactura.getRange("E27").getValue();
  const formaPago = hojaFactura.getRange("C8").getValue();

  // Validar que todos los campos necesarios tengan datos
  if (!fecha || !concepto || !cliente || !totalVenta || !formaPago) {
    SpreadsheetApp.getUi().alert("Por favor, asegúrate de llenar todos los campos requeridos.");
    return;
  }

  // Conectar a la hoja de cálculo destino
  const idHojaDestino = "1J2pVzi7D30Df2OMgtHkj1-BiUdmvaCE_V4lVevwilh0"; // Reemplaza con el ID de la hoja de cálculo destino
  const ssDestino = SpreadsheetApp.openById(idHojaDestino);

  // Determinar la hoja del mes correspondiente
  const mes = obtenerMes(new Date(fecha));
  const hojaMes = ssDestino.getSheetByName(mes);

  if (!hojaMes) {
    SpreadsheetApp.getUi().alert(`La hoja para el mes ${mes} no existe en el archivo destino.`);
    return;
  }

  // Registrar los datos en la hoja mensual del archivo destino
  hojaMes.appendRow([fecha, concepto, cliente, totalVenta, formaPago]);
  generarIDEnHoja("Factura");
  // Mensaje de confirmación
  //SpreadsheetApp.getUi().alert("La venta ha sido registrada correctamente en la hoja correspondiente del archivo destino.");
}
*/

/*function registrarVenta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = ss.getSheetByName("Factura");
  const hojaInventario = ss.getSheetByName("Inventario");

  // Abrir el archivo donde están las hojas mensuales
  const archivoVentasId = "1J2pVzi7D30Df2OMgtHkj1-BiUdmvaCE_V4lVevwilh0"; // Reemplaza con el ID del archivo
  const archivoVentas = SpreadsheetApp.openById(archivoVentasId);

  if (!hojaInventario) {
    SpreadsheetApp.getUi().alert("La hoja 'Inventario' no existe.");
    return;
  }

  // Datos generales de la venta
  const fecha = hojaFactura.getRange("C5").getValue(); // Fecha de la factura
  const concepto = hojaFactura.getRange("B3").getValue() + " " + hojaFactura.getRange("C3").getValue();
  const cliente = hojaFactura.getRange("B9").getValue();
  const formaPago = hojaFactura.getRange("C8").getValue();

  if (!fecha || !concepto || !cliente || !formaPago) {
    SpreadsheetApp.getUi().alert("Por favor, completa todos los campos generales de la factura.");
    return;
  }

  // Determinar la hoja mensual basada en la fecha
  const mes = fecha.toLocaleString("es-ES", { month: "long" }); // Nombre del mes en español
  const hojaVentas = archivoVentas.getSheetByName(mes.charAt(0).toUpperCase() + mes.slice(1)); // Capitalizar el nombre

  if (!hojaVentas) {
    SpreadsheetApp.getUi().alert(`La hoja para el mes de ${mes} no existe en el archivo de ventas.`);
    return;
  }

  // Procesar productos de la factura
  const rangoProductos = hojaFactura.getRange("A15:C24"); // Ajusta el rango según tu diseño
  const datosFactura = rangoProductos.getValues();
  const datosInventario = hojaInventario.getDataRange().getValues();

  let filasVenta = []; // Filas de productos vendidos
  let totalVenta = 0;
  let totalCosto = 0;
  let totalGanancia = 0;

  datosFactura.forEach(producto => {
    const codigoProducto = producto[0]; // Código del producto
    const cantidadVendida = producto[2]; // Cantidad vendida

    if (codigoProducto && cantidadVendida > 0) {
      const productoInventario = datosInventario.find(row => row[0] === codigoProducto);

      if (productoInventario) {
        const costoUnitario = productoInventario[6]; // Columna G en inventario
        const precioUnitario = producto[3];
        const costoTotal = costoUnitario * cantidadVendida;
        const subtotalVenta = precioUnitario * cantidadVendida;
        const ganancia = subtotalVenta - costoTotal;

        totalVenta += subtotalVenta;
        totalCosto += costoTotal;
        totalGanancia += ganancia;

        filasVenta.push([
          fecha,
          concepto,
          cliente,
          codigoProducto,
          cantidadVendida,
          precioUnitario,
          costoUnitario, // Agregamos el costo unitario
          costoTotal,
          ganancia,
          formaPago
        ]);
      } else {
        SpreadsheetApp.getUi().alert(`El producto con código ${codigoProducto} no se encontró en el inventario.`);
      }
    }
  });

  // Agregar datos a la hoja mensual
  if (filasVenta.length > 0) {
    const rangoInicio = hojaVentas.getLastRow() + 1;

    // Insertar filas de productos vendidos
    hojaVentas.getRange(rangoInicio, 1, filasVenta.length, filasVenta[0].length).setValues(filasVenta);

    // Agregar la fila de totales
    hojaVentas.getRange(rangoInicio + filasVenta.length, 1, 1, 10).setValues([
      [
        "TOTAL", // Texto de total
        "", "", "", // Concepto, Cliente, Código vacíos
        "", "", // Cantidad y Precio Unitario vacíos
        "", // Costo Unitario vacío
        totalCosto, // Costo total
        totalGanancia, // Ganancia total
        "" // Forma de pago vacía
      ]
    ]);

    SpreadsheetApp.getUi().alert(`Venta registrada correctamente en la hoja de ${mes}.`);
  } else {
    SpreadsheetApp.getUi().alert("No se encontraron productos válidos en la factura.");
  }
}*/

/*function registrarVenta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = ss.getSheetByName("Factura");
  const hojaInventario = ss.getSheetByName("Inventario");

  // Abrir el archivo donde están las hojas mensuales
  const archivoVentasId = "1J2pVzi7D30Df2OMgtHkj1-BiUdmvaCE_V4lVevwilh0"; // Reemplaza con el ID del archivo
  const archivoVentas = SpreadsheetApp.openById(archivoVentasId);

  if (!hojaInventario) {
    SpreadsheetApp.getUi().alert("La hoja 'Inventario' no existe.");
    return;
  }

  // Datos generales de la venta
  const fecha = hojaFactura.getRange("C5").getValue(); // Fecha de la factura
  const concepto = hojaFactura.getRange("B3").getValue() + " " + hojaFactura.getRange("C3").getValue();
  const cliente = hojaFactura.getRange("B9").getValue();
  const formaPago = hojaFactura.getRange("C8").getValue();

  if (!fecha || !concepto || !cliente || !formaPago) {
    SpreadsheetApp.getUi().alert("Por favor, completa todos los campos generales de la factura.");
    return;
  }

  // Determinar la hoja mensual basada en la fecha
  const mes = fecha.toLocaleString("es-ES", { month: "long" }); // Nombre del mes en español
  const hojaVentas = archivoVentas.getSheetByName(mes.charAt(0).toUpperCase() + mes.slice(1)); // Capitalizar el nombre

  if (!hojaVentas) {
    SpreadsheetApp.getUi().alert(`La hoja para el mes de ${mes} no existe en el archivo de ventas.`);
    return;
  }

  // Procesar productos de la factura
  const rangoProductos = hojaFactura.getRange("A15:D24"); // Ajusta el rango para incluir el precio unitario
  const datosFactura = rangoProductos.getValues();
  const datosInventario = hojaInventario.getDataRange().getValues();

  let filasVenta = []; // Filas de productos vendidos
  let totalVenta = 0;
  let totalCosto = 0;
  let totalGanancia = 0;

  datosFactura.forEach(producto => {
    const codigoProducto = producto[0]; // Código del producto
    const cantidadVendida = producto[2]; // Cantidad vendida
    const precioUnitario = producto[3]; // Precio unitario desde la columna D

    if (codigoProducto && cantidadVendida > 0 && precioUnitario) {
      const productoInventario = datosInventario.find(row => row[0] === codigoProducto);

      if (productoInventario) {
        const costoUnitario = productoInventario[6]; // Columna G en inventario
        const costoTotal = costoUnitario * cantidadVendida;
        const subtotalVenta = precioUnitario * cantidadVendida;
        const ganancia = subtotalVenta - costoTotal;

        totalVenta += subtotalVenta;
        totalCosto += costoTotal;
        totalGanancia += ganancia;

        filasVenta.push([
          fecha,
          concepto,
          cliente,
          codigoProducto,
          cantidadVendida,
          precioUnitario,
          costoUnitario, // Costo unitario
          costoTotal,
          ganancia,
          formaPago
        ]);
      } else {
        SpreadsheetApp.getUi().alert(`El producto con código ${codigoProducto} no se encontró en el inventario.`);
      }
    }
  });

  // Agregar datos a la hoja mensual
  if (filasVenta.length > 0) {
    const rangoInicio = hojaVentas.getLastRow() + 1;

    // Insertar filas de productos vendidos
    hojaVentas.getRange(rangoInicio, 1, filasVenta.length, filasVenta[0].length).setValues(filasVenta);

    // Agregar la fila de totales
    hojaVentas.getRange(rangoInicio + filasVenta.length, 1, 1, 10).setValues([
      [
        fecha, // Fecha de la venta
        "TOTAL", // Texto de total
        "TOTAL", // Cliente vacío
        "TOTAL", // Código vacío
        "TOTAL", // Cantidad vacía
        "TOTAL", // Precio unitario vacío
        "TOTAL", // Costo unitario vacío
        totalVenta, // Total de la venta
        totalGanancia, // Total de la ganancia
        "TOTAL" // Forma de pago vacía
      ]
    ]);
    generarIDEnHoja("Factura");

    SpreadsheetApp.getUi().alert(`Venta registrada correctamente en la hoja de ${mes}.`);
  } else {
    SpreadsheetApp.getUi().alert("No se encontraron productos válidos en la factura.");
  }
}*/
function revisionVenta() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    "Confirmación",
    "¿Está seguro de que ha revisado la factura y que todo está correctamente?",
    ui.ButtonSet.YES_NO
  );

  if (respuesta === ui.Button.NO) {
    ui.alert("El registro de la venta fue cancelado.");
    return;
  }
}

/*function registrarVenta() {
  

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = ss.getSheetByName("Factura");
  const hojaInventario = ss.getSheetByName("Inventario");
  const hojaCajaGeneral = ss.getSheetByName("Caja General");

  // Abrir el archivo donde están las hojas mensuales
  const archivoVentasId = "1J2pVzi7D30Df2OMgtHkj1-BiUdmvaCE_V4lVevwilh0"; // Reemplaza con el ID del archivo
  const archivoVentas = SpreadsheetApp.openById(archivoVentasId);

  if (!hojaInventario) {
    SpreadsheetApp.getUi().alert("La hoja 'Inventario' no existe.");
    return;
  }

  // Datos generales de la venta
  const fecha = hojaFactura.getRange("C5").getValue(); // Fecha de la factura
  const concepto = hojaFactura.getRange("B3").getValue() + " " + hojaFactura.getRange("C3").getValue();
  const cliente = hojaFactura.getRange("B9").getValue();
  const formaPago = hojaFactura.getRange("C8").getValue();

  if (!fecha || !concepto || !cliente || !formaPago) {
    SpreadsheetApp.getUi().alert("Por favor, completa todos los campos generales de la factura.");
    return;
  }

  // Determinar la hoja mensual basada en la fecha
  const mes = fecha.toLocaleString("es-ES", { month: "long" }); // Nombre del mes en español
  const hojaVentas = archivoVentas.getSheetByName(mes.charAt(0).toUpperCase() + mes.slice(1)); // Capitalizar el nombre

  if (!hojaVentas) {
    SpreadsheetApp.getUi().alert(`La hoja para el mes de ${mes} no existe en el archivo de ventas.`);
    return;
  }

  // Procesar productos de la factura
  const rangoProductos = hojaFactura.getRange("A15:D24"); // Ajusta el rango para incluir el precio unitario
  const datosFactura = rangoProductos.getValues();
  const datosInventario = hojaInventario.getDataRange().getValues();

  let filasVenta = []; // Filas de productos vendidos
  let totalVenta = 0;
  let totalCosto = 0;
  let totalGanancia = 0;

  datosFactura.forEach(producto => {
    const codigoProducto = producto[0]; // Código del producto
    const cantidadVendida = producto[2]; // Cantidad vendida
    const precioUnitario = producto[3]; // Precio unitario desde la columna D

    if (codigoProducto && cantidadVendida > 0 && precioUnitario) {
      const productoInventario = datosInventario.find(row => row[0] === codigoProducto);

      if (productoInventario) {
        const costoUnitario = productoInventario[6]; // Columna G en inventario
        const costoTotal = costoUnitario * cantidadVendida;
        const subtotalVenta = precioUnitario * cantidadVendida;
        const ganancia = subtotalVenta - costoTotal;

        totalVenta += subtotalVenta;
        totalCosto += costoTotal;
        totalGanancia += ganancia;

        filasVenta.push([
          fecha,
          concepto,
          cliente,
          codigoProducto,
          cantidadVendida,
          precioUnitario,
          costoUnitario, // Costo unitario
          costoTotal,
          ganancia,
          formaPago
        ]);
      } else {
        SpreadsheetApp.getUi().alert(`El producto con código ${codigoProducto} no se encontró en el inventario.`);
      }
    }
  });

  // Agregar datos a la hoja mensual
  if (filasVenta.length > 0) {
    const rangoInicio = hojaVentas.getLastRow() + 1;

    // Insertar filas de productos vendidos
    hojaVentas.getRange(rangoInicio, 1, filasVenta.length, filasVenta[0].length).setValues(filasVenta);

    // Agregar la fila de totales
    hojaVentas.getRange(rangoInicio + filasVenta.length, 1, 1, 10).setValues([
      [
        fecha, // Fecha de la venta
        "TOTAL", // Texto de total
        "TOTAL", // Cliente vacío
        "TOTAL", // Código vacío
        "TOTAL", // Cantidad vacía
        "TOTAL", // Precio unitario vacío
        "TOTAL", // Costo unitario vacío
        totalVenta, // Total de la venta
        totalGanancia, // Total de la ganancia
        "TOTAL" // Forma de pago vacía
      ]
    ]);

    // Agregar datos a la hoja Caja General
    const filaCajaGeneral = hojaCajaGeneral.getLastRow() + 1;
    const ingresoAnterior = hojaCajaGeneral.getRange(filaCajaGeneral - 1, 5).getValue() || 0; // Columna E de la fila anterior
    const ingresoTotal = ingresoAnterior + totalVenta;

    hojaCajaGeneral.getRange(filaCajaGeneral, 1, 1, 5).setValues([
      [fecha, concepto, totalVenta, 0, ingresoTotal] // Columna D es 0; Columna E suma acumulada
    ]);

    generarIDEnHoja("Factura");

    SpreadsheetApp.getUi().alert(`Venta registrada correctamente en la hoja de ${mes} y en 'Caja General'.`);
  } else {
    SpreadsheetApp.getUi().alert("No se encontraron productos válidos en la factura.");
  }
}*/


















function registrarVenta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = ss.getSheetByName("Factura");
  const hojaInventario = ss.getSheetByName("Inventario");
  const hojaCajaGeneral = ss.getSheetByName("Caja General");

  // Abrir el archivo donde están las hojas mensuales
  const archivoVentasId = "18s_q9bG9p7Jp2PqbxQSMMYMIslNbbPm2hHno2nBDFWg"; // Reemplaza con el ID del archivo
  const archivoVentas = SpreadsheetApp.openById(archivoVentasId);

  const archivoClientesId = "1H5sG4EkRVIIxihQbmAUwcwrpmZjiYH06r-1gCYp_QRY"; // Reemplaza con el ID del archivo
  const archivoClientes = SpreadsheetApp.openById(archivoClientesId);

  if (!hojaInventario) {
    SpreadsheetApp.getUi().alert("La hoja 'Inventario' no existe.");
    return;
  }

  // Datos generales de la venta
  const fecha = hojaFactura.getRange("C5").getValue(); // Fecha de la factura
  const concepto = hojaFactura.getRange("B3").getValue() + " " + hojaFactura.getRange("C3").getValue();
  const cliente = hojaFactura.getRange("B9").getValue();
  const formaPago = hojaFactura.getRange("C8").getValue();
  const totalVentaCredito = hojaFactura.getRange("E27").getValue();

  if (!fecha || !concepto || !cliente || !formaPago) {
    SpreadsheetApp.getUi().alert("Por favor, completa todos los campos generales de la factura.");
    return;
  }



  if (formaPago === "CREDITO") {
  const hojaCliente = archivoClientes.getSheetByName(cliente) || archivoClientes.insertSheet(cliente);

  // Verificar si la hoja está vacía y agregar encabezados
  if (hojaCliente.getLastRow() === 0) {
    hojaCliente.appendRow(["Fecha", "Concepto", "Total", "Abonos", "Descuentos", "Saldo"]);
  } else {
    SpreadsheetApp.getUi().alert(`No se ha logrado`);
  }

  // Obtener la última fila con datos
  const ultimaFila = hojaCliente.getLastRow();
  let saldoAnterior = 0;

  // Si ya existen registros, obtener el saldo anterior de la última fila
  if (ultimaFila > 1) {
    saldoAnterior = hojaCliente.getRange(ultimaFila, 6).getValue(); // Columna F (Saldo)
  }

  // Calcular el nuevo saldo
  const abonos = 0;      // Inicialmente 0 porque aún no hay pagos
  const descuentos = 0;  // Inicialmente 0 si no hay descuentos aplicados
  const nuevoSaldo = saldoAnterior + totalVentaCredito - abonos - descuentos;

  // Agregar la fila con la información correcta
  hojaCliente.appendRow([fecha, concepto, totalVentaCredito, abonos, descuentos, nuevoSaldo]);
}
 else if(formaPago === "CONTADO") {

  } else {
    SpreadsheetApp.getUi().alert(`Esta forma de pago no es compatible`);
  }

  // Determinar la hoja mensual basada en la fecha
  const mes = fecha.toLocaleString("es-ES", { month: "long" }); // Nombre del mes en español
  const hojaVentas = archivoVentas.getSheetByName(mes.charAt(0).toUpperCase() + mes.slice(1)); // Capitalizar el nombre

  if (!hojaVentas) {
    SpreadsheetApp.getUi().alert(`La hoja para el mes de ${mes} no existe en el archivo de ventas.`);
    return;
  }

  // Procesar productos de la factura
  const rangoProductos = hojaFactura.getRange("A15:D24"); // Ajusta el rango para incluir el precio unitario
  const datosFactura = rangoProductos.getValues();
  const datosInventario = hojaInventario.getDataRange().getValues();

  let filasVenta = []; // Filas de productos vendidos
  let totalVenta = 0;
  let totalCosto = 0;
  let totalGanancia = 0;

  datosFactura.forEach(producto => {
    const codigoProducto = producto[0]; // Código del producto
    const cantidadVendida = producto[2]; // Cantidad vendida
    const precioUnitario = producto[3]; // Precio unitario desde la columna D

    if (codigoProducto && cantidadVendida > 0 && precioUnitario) {
      const productoInventario = datosInventario.find(row => row[0] === codigoProducto);

      if (productoInventario) {
        const costoUnitario = productoInventario[6]; // Columna G en inventario
        const costoTotal = costoUnitario * cantidadVendida;
        const subtotalVenta = precioUnitario * cantidadVendida;
        const ganancia = subtotalVenta - costoTotal;

        totalVenta += subtotalVenta;
        totalCosto += costoTotal;
        totalGanancia += ganancia;

        filasVenta.push([
          fecha,
          concepto,
          cliente,
          codigoProducto,
          cantidadVendida,
          precioUnitario,
          costoUnitario, // Costo unitario
          costoTotal,
          ganancia,
          formaPago
        ]);
      } else {
        SpreadsheetApp.getUi().alert(`El producto con código ${codigoProducto} no se encontró en el inventario.`);
      }
    }
  });

  // Agregar datos a la hoja mensual
  if (filasVenta.length > 0) {
    const rangoInicio = hojaVentas.getLastRow() + 1;

    // Insertar filas de productos vendidos
    hojaVentas.getRange(rangoInicio, 1, filasVenta.length, filasVenta[0].length).setValues(filasVenta);

    // Agregar la fila de totales
    const filaTotal = rangoInicio + filasVenta.length;
    hojaVentas.getRange(filaTotal, 1, 1, 10).setValues([
      [
        fecha, // Fecha de la venta
        "TOTAL", // Texto de total
        "TOTAL", // Cliente vacío
        "TOTAL", // Código vacío
        "TOTAL", // Cantidad vacía
        "TOTAL", // Precio unitario vacío
        "TOTAL", // Costo unitario vacío
        totalVenta, // Total de la venta
        totalGanancia, // Total de la ganancia
        "TOTAL" // Forma de pago vacía
      ]
    ]);

    // Pintar la fila de totales en verde
    hojaVentas.getRange(filaTotal, 1, 1, 10).setBackground("#00ff00");

    // Excluir las celdas verdes de las columnas H e I al calcular el total
    const rangoH = hojaVentas.getRange(2, 8, filaTotal - 1).getValues(); // Columna H (sin totales)
    const rangoI = hojaVentas.getRange(2, 9, filaTotal - 1).getValues(); // Columna I (sin totales)
    const coloresH = hojaVentas.getRange(2, 8, filaTotal - 1).getBackgrounds(); // Colores de H
    const coloresI = hojaVentas.getRange(2, 9, filaTotal - 1).getBackgrounds(); // Colores de I

    let sumaH = 0;
    let sumaI = 0;

    rangoH.forEach((valor, i) => {
      if (coloresH[i][0] !== "#00ff00") {
        sumaH += parseFloat(valor[0]) || 0;
      }
    });

    rangoI.forEach((valor, i) => {
      if (coloresI[i][0] !== "#00ff00") {
        sumaI += parseFloat(valor[0]) || 0;
      }
    });

    hojaVentas.getRange("M1").setValue(sumaH); // Total para H
    hojaVentas.getRange("M2").setValue(sumaI); // Total para I

    // Agregar datos a la hoja Caja General
    const filaCajaGeneral = hojaCajaGeneral.getLastRow() + 1;
    const ingresoAnterior = hojaCajaGeneral.getRange(filaCajaGeneral - 1, 5).getValue() || 0; // Columna E de la fila anterior
    const ingresoTotal = ingresoAnterior + totalVenta;

    hojaCajaGeneral.getRange(filaCajaGeneral, 1, 1, 5).setValues([
      [fecha, concepto, totalVenta, 0, ingresoTotal] // Columna D es 0; Columna E suma acumulada
    ]);

    generarIDEnHoja("Factura");

    SpreadsheetApp.getUi().alert(`Venta registrada correctamente en la hoja de ${mes} y en 'Caja General'.`);
  } else {
    SpreadsheetApp.getUi().alert("No se encontraron productos válidos en la factura.");
  }
}













// Función para obtener el nombre del mes
function obtenerMes(fecha) {
  const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  return meses[fecha.getMonth()];
}


function procesarVentaYActualizarInventario() {
  // Ejecutar la función registrarVenta
  try {
    revisionVenta();
    registrarVenta(); // Registrar la venta en la hoja mensual

  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error al registrar la venta: ${error.message}`);
    return; // Detener si hay un error en esta etapa
  }

  // Ejecutar la función registrarFactura
  try {
    registrarFactura(); // Actualizar el inventario basado en la factura
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error al registrar la factura: ${error.message}`);
    return; // Detener si hay un error en esta etapa
  }

  // Confirmación final
  SpreadsheetApp.getUi().alert("Proceso completado: Venta registrada y factura procesada.");
}

// CAJA GENERAL COMPRAS

/*function registrarCompras() {
  const ssOrigen = SpreadsheetApp.getActiveSpreadsheet(); // Hoja actual
  const hojaCompra = ssOrigen.getSheetByName("Compra"); // Hoja donde se registran las compras

  // Obtener los datos de la compra
  const fecha = hojaCompra.getRange("C5").getValue();
  const proveedor = hojaCompra.getRange("B9").getValue(); // Proveedor
  const concepto = hojaCompra.getRange("B3").getValue() + " " + hojaCompra.getRange("C3").getValue();
  const totalCompra = hojaCompra.getRange("E27").getValue();
  const formaPago = hojaCompra.getRange("C8").getValue();

  // Validar que todos los campos necesarios tengan datos
  if (!fecha || !proveedor || !concepto || !totalCompra || !formaPago) {
    SpreadsheetApp.getUi().alert("Por favor, asegúrate de llenar todos los campos requeridos.");
    return;
  }

  // Conectar a la hoja de cálculo destino
  const idHojaDestino = "1fxHz1T0Fw7l4feNpO1w1Z0HK5nWsGvzE2g3sDC67Vx8"; // Reemplaza con el ID de la hoja de cálculo destino para compras
  const ssDestino = SpreadsheetApp.openById(idHojaDestino);

  // Determinar la hoja del mes correspondiente
  const mes = obtenerMes(new Date(fecha));
  const hojaMes = ssDestino.getSheetByName(mes);

  if (!hojaMes) {
    SpreadsheetApp.getUi().alert(`La hoja para el mes ${mes} no existe en el archivo de compras.`);
    return;
  }

  // Registrar los datos en la hoja mensual del archivo destino
  hojaMes.appendRow([fecha, proveedor, concepto, totalCompra, formaPago]);
  generarIDEnHoja("Compra");
  // Mensaje de confirmación
  //SpreadsheetApp.getUi().alert("La compra ha sido registrada correctamente en la hoja correspondiente del archivo destino.");
}*/

function registrarCompras() {
  /*const ui = SpreadsheetApp.getUi();

  // Confirmar si se desea realizar la compra
  const respuesta = ui.alert(
    "Confirmación de Compra",
    "¿Ha revisado la factura y está seguro de registrar esta compra?",
    ui.ButtonSet.YES_NO
  );

  if (respuesta === ui.Button.NO) {
    ui.alert("El registro de la compra fue cancelado.");
    return;
  }*/

  const ssOrigen = SpreadsheetApp.getActiveSpreadsheet(); // Hoja actual
  const hojaCompra = ssOrigen.getSheetByName("Compra"); // Hoja donde se registran las compras

  // Obtener los datos de la compra
  const fecha = hojaCompra.getRange("C5").getValue();
  const proveedor = hojaCompra.getRange("B9").getValue(); // Proveedor
  const concepto = hojaCompra.getRange("B3").getValue() + " " + hojaCompra.getRange("C3").getValue();
  const totalCompra = hojaCompra.getRange("E27").getValue();
  const formaPago = hojaCompra.getRange("C8").getValue();

  // Validar que todos los campos necesarios tengan datos
  if (!fecha || !proveedor || !concepto || !totalCompra || !formaPago) {
    SpreadsheetApp.getUi().alert("Por favor, asegúrate de llenar todos los campos requeridos.");
    return;
  }

  // Conectar a la hoja de cálculo destino
  const idHojaDestino = "1fxHz1T0Fw7l4feNpO1w1Z0HK5nWsGvzE2g3sDC67Vx8"; // Reemplaza con el ID de la hoja de cálculo destino para compras
  const ssDestino = SpreadsheetApp.openById(idHojaDestino);

  // Determinar la hoja del mes correspondiente
  const mes = obtenerMes(new Date(fecha));
  const hojaMes = ssDestino.getSheetByName(mes);

  if (!hojaMes) {
    SpreadsheetApp.getUi().alert(`La hoja para el mes ${mes} no existe en el archivo de compras.`);
    return;
  }

  // Registrar los datos en la hoja mensual del archivo destino
  hojaMes.appendRow([fecha, proveedor, concepto, totalCompra, formaPago]);

  // Registrar en "Caja General"
  const hojaCajaGeneral = ssOrigen.getSheetByName("Caja General");
  if (!hojaCajaGeneral) {
    SpreadsheetApp.getUi().alert("La hoja 'Caja General' no existe.");
    return;
  }

  const filaCajaGeneral = hojaCajaGeneral.getLastRow() + 1;
  const saldoAnterior = hojaCajaGeneral.getRange(filaCajaGeneral - 1, 5).getValue() || 0; // Columna E de la fila anterior
  const nuevoSaldo = saldoAnterior - totalCompra; // Restar el egreso

  hojaCajaGeneral.getRange(filaCajaGeneral, 1, 1, 5).setValues([
    [fecha, concepto, "", totalCompra, nuevoSaldo] // Columna C vacía, D es el egreso, E es el saldo acumulado
  ]);

  generarIDEnHoja("Compra");

  // Mensaje de confirmación
  SpreadsheetApp.getUi().alert("La compra ha sido registrada correctamente en la hoja correspondiente y en 'Caja General'.");
}




// Función para obtener el nombre del mes
function obtenerMes(fecha) {
  const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  return meses[fecha.getMonth()];
}

/*function registrarCompras(nuevoID) {
  const ssOrigen = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = ssOrigen.getSheetByName("Factura");

  // Datos de la factura
  const fecha = hojaFactura.getRange("C5").getValue();
  const concepto = hojaFactura.getRange("B3").getValue() + " " + hojaFactura.getRange("C3").getValue();
  const proveedor = hojaFactura.getRange("B9").getValue(); // Proveedor
  const totalCompra = hojaFactura.getRange("E27").getValue();
  const formaPago = hojaFactura.getRange("C8").getValue();

  // Validar campos
  if (!fecha || !concepto || !proveedor || !totalCompra || !formaPago) {
    SpreadsheetApp.getUi().alert("Por favor, asegúrate de llenar todos los campos requeridos.");
    return;
  }

  // Conectar a la hoja de cálculo destino
  const idHojaDestino = "1fxHz1T0Fw7l4feNpO1w1Z0HK5nWsGvzE2g3sDC67Vx8"; // Reemplaza con el ID de la hoja de cálculo destino
  const ssDestino = SpreadsheetApp.openById(idHojaDestino);

  // Determinar la hoja del mes correspondiente
  const mes = obtenerMes(new Date(fecha));
  const hojaMes = ssDestino.getSheetByName(mes);

  if (!hojaMes) {
    SpreadsheetApp.getUi().alert(`La hoja para el mes ${mes} no existe en el archivo destino.`);
    return;
  }

  // Registrar los datos en la hoja mensual del archivo destino
  hojaMes.appendRow([nuevoID, fecha, concepto, proveedor, totalCompra, formaPago]);
}*/


function confirmarYRegistrarCompras() {
  const ui = SpreadsheetApp.getUi();

  // Confirmar si se desea proceder
  const respuesta = ui.alert(
    "Confirmación de Compra",
    "¿Ha revisado las facturas y está seguro de registrar estas compras?",
    ui.ButtonSet.YES_NO
  );

  if (respuesta === ui.Button.NO) {
    ui.alert("El registro de las compras fue cancelado.");
    return;
  }

  // Llamar a las funciones específicas si la respuesta es 'Sí'
  //registrarCompras();
  //registrarCompra();
}



function ejecutarRegistro() {
  try {
    
    confirmarYRegistrarCompras();
    // Llamar a la función para manejar el stock
    registrarCompra();

    // Llamar a la función para registrar la compra en la hoja "Compras"
    registrarCompras();
    
    // Confirmación al usuario
    SpreadsheetApp.getUi().alert("El registro de la compra y la actualización del stock se han realizado correctamente.");
  } catch (error) {
    // Manejo de errores
    SpreadsheetApp.getUi().alert(`Ocurrió un error: ${error.message}`);
  }
}

// GASTOS

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Menú para registrar clientes
  ui.createMenu("Registro de Clientes")
    .addItem("Registrar Cliente", "mostrarFormularioClientes")
    .addToUi();
  
  // Menú para registrar gastos
  ui.createMenu("Registro de Gastos")
    .addItem("Registrar Gasto", "mostrarFormularioGasto")
    .addToUi();
}


function mostrarFormularioGasto() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("FormularioGasto")
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Registrar Gasto");
}

function registrarGastoCaja(fecha, concepto, monto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCajaGeneral = ss.getSheetByName("Caja General");

  if (!hojaCajaGeneral) {
    throw new Error("La hoja 'Caja General' no existe.");
  }

  const ultimaFila = hojaCajaGeneral.getLastRow() + 1;

  hojaCajaGeneral.getRange(ultimaFila, 1).setValue(new Date(fecha)); // Fecha
  hojaCajaGeneral.getRange(ultimaFila, 2).setValue(concepto);       // Concepto
  hojaCajaGeneral.getRange(ultimaFila, 4).setValue(monto);          // Monto
}





// ESTADO DE RESULTADOS

function generarEstadoDeResultados () {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archivoVentasId = "1J2pVzi7D30Df2OMgtHkj1-BiUdmvaCE_V4lVevwilh0"; // Reemplaza con el ID del archivo
  const archivoVentas = SpreadsheetApp.openById(archivoVentasId);

  // Crear o abrir la hoja "Estado de Resultados"
  let hojaResultados = ss.getSheetByName("Estado de Resultados");
  if (!hojaResultados) {
    hojaResultados = ss.insertSheet("Estado de Resultados");
  } else {
    hojaResultados.clear(); // Limpiar contenido previo
  }

  // Configurar encabezados
  hojaResultados.getRange("A1").setValue("Estado de Resultados");
  hojaResultados.getRange("A3:B3").setValues([["Concepto", "Monto"]]);

  let fila = 4; // Primera fila para los resultados

  // Ingresos Totales
  let ingresosTotales = 0;
  const hojasMensuales = archivoVentas.getSheets();
  hojasMensuales.forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    datos.forEach(row => {
      if (typeof row[7] === "number") { // Total de venta (Columna H)
        ingresosTotales += row[7];
      }
    });
  });
  hojaResultados.getRange(fila, 1, 1, 2).setValues([["Ingresos Totales", ingresosTotales]]);
  fila++;

  // Costos Totales
  let costosTotales = 0;
  hojasMensuales.forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    datos.forEach(row => {
      if (typeof row[8] === "number") { // Total costo (Columna G)
        costosTotales += row[8];
      }
    });
  });
  hojaResultados.getRange(fila, 1, 1, 2).setValues([["Costos Totales", costosTotales]]);
  fila++;

  // Utilidad Bruta
  const utilidadBruta = ingresosTotales - costosTotales;
  hojaResultados.getRange(fila, 1, 1, 2).setValues([["Utilidad Bruta", utilidadBruta]]);
  fila++;

  // Gastos Operativos
  const hojaGastos = ss.getSheetByName("Gastos");
  let gastosOperativos = 0;
  if (hojaGastos) {
    const datosGastos = hojaGastos.getDataRange().getValues();
    datosGastos.forEach(row => {
      if (typeof row[3] === "number") { // Suponiendo que los montos están en la columna B
        gastosOperativos += row[1];
      }
    });
  }
  hojaResultados.getRange(fila, 1, 1, 3).setValues([["Gastos Operativos", gastosOperativos]]);
  fila++;

  // Utilidad Operativa
  const utilidadOperativa = utilidadBruta - gastosOperativos;
  hojaResultados.getRange(fila, 1, 1, 2).setValues([["Utilidad Operativa", utilidadOperativa]]);
  fila++;

  // Otros Ingresos/Gastos (pueden ajustarse según necesidades)
  const otrosIngresos = 0; // Cambiar si aplica
  const otrosGastos = 0; // Cambiar si aplica
  hojaResultados.getRange(fila, 1, 1, 2).setValues([["Otros Ingresos/Gastos", otrosIngresos - otrosGastos]]);
  fila++;

  // Utilidad Neta
  const utilidadNeta = utilidadOperativa + (otrosIngresos - otrosGastos);
  hojaResultados.getRange(fila, 1, 1, 2).setValues([["Utilidad Neta", utilidadNeta]]);
  fila++;

  SpreadsheetApp.getUi().alert("Estado de Resultados generado correctamente.");
}
