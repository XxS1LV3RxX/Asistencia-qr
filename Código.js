function doGet(e) {
    var id = e.parameter.id;
    var nombre = e.parameter.nombre;
    var tipo = e.parameter.tipo;

    Logger.log("Parámetros recibidos: " + JSON.stringify(e.parameters));
    Logger.log("ID: " + id);
    Logger.log("Nombre: " + nombre);
    Logger.log("Tipo: " + tipo);

    if (!id || !nombre || !tipo) {
        return mostrarError("Faltan parámetros obligatorios.");
    }


    var hora = new Date();
    var horaFormateada = hora.toLocaleString("es-MX", { timeZone: "America/Mexico_City", hour: "2-digit", minute: "2-digit", second: "2-digit" });

    var diaPermitido = hora.getDay() === 6; 
    var horaApertura = new Date(hora);
    horaApertura.setHours(09, 40, 0);
    var horaLimiteEntrada = new Date(hora);
    horaLimiteEntrada.setHours(10, 16, 0);
    var horaLimiteCierre = new Date(hora);
    horaLimiteCierre.setHours(12, 30, 0);

    var correoUsuario = Session.getActiveUser().getEmail();
    var correoPropietario = Session.getEffectiveUser().getEmail();
    var correoProyecto = "sbregasistencia@gmail.com";

    if (correoUsuario !== correoPropietario) {
        enviarCorreoAlerta(
            correoProyecto,
            "Intento de registro no autorizado",
            `Se ha detectado un intento de registro no autorizado:
            - ID: ${id}
            - Nombre: ${nombre}
            - Tipo: ${tipo}
            - Correo del usuario: ${correoUsuario}
            - Hora del intento: ${horaFormateada}`
        );
        return mostrarError("No estás autorizado para realizar esta acción.");
    }

    if (!diaPermitido || hora < horaApertura || hora > horaLimiteCierre) {
        enviarCorreoAlerta(
            correoProyecto,
            "Intento de registro fuera de horario",
            `Se ha detectado un intento de registro fuera del horario permitido:
            - ID: ${id}
            - Nombre: ${nombre}
            - Tipo: ${tipo}
            - Correo del usuario: ${correoUsuario}
            - Hora del intento: ${horaFormateada}`
        );
        return mostrarRegistroCerrado("No puedes registrar asistencia fuera del horario permitido.");
    }

    var spreadsheet = SpreadsheetApp.openById("1ZJYL8tzo9Kfw235Zdl2UR29mLjFj-pyHElUdJZ-dRHU");
    var sheetAsistencia = spreadsheet.getSheetByName("Asistencia");
    var sheetAsisAnual = spreadsheet.getSheetByName("AsisAnual");

    var range = sheetAsistencia.getDataRange();
    var values = range.getValues();
    var rowIndex = -1;

    for (var i = 0; i < values.length; i++) {
        if (values[i][0] == id) {
            rowIndex = i + 1;
            break;
        }
    }

    if (rowIndex == -1) {
        rowIndex = sheetAsistencia.getLastRow() + 1;
        sheetAsistencia.getRange(rowIndex, 1).setValue(id);
        sheetAsistencia.getRange(rowIndex, 2).setValue(nombre);
    }

    sheetAsistencia.getRange(rowIndex, 4).setValue(new Date());

    var horaEntrada = sheetAsistencia.getRange(rowIndex, 5).getValue();
    var horaSalida = sheetAsistencia.getRange(rowIndex, 6).getValue();
    var semaforo = sheetAsistencia.getRange(rowIndex, 7);

    if (tipo == "entrada") {
    if (horaEntrada) {
        return mostrarError("Ya has registrado tu entrada.");
    }
    sheetAsistencia.getRange(rowIndex, 5).setValue(horaFormateada);

    if (hora > horaLimiteEntrada) {
        semaforo.setValue("");
        semaforo.setBackground("yellow");
    } else {
        semaforo.setValue(""); 
        semaforo.setBackground("green");
    }

    return mostrarMensaje("¡Entrada registrada!", `Bienvenido, <span class="highlight">${nombre}</span> (ID: ${id}).`, hora);
}

    if (tipo == "salida") {
    if (!horaEntrada) {
        return mostrarError("Para registrar tu salida, debes registrar primero tu entrada.");
    }
    if (horaSalida) {
        return mostrarError("Ya has registrado tu salida.");
    }
    sheetAsistencia.getRange(rowIndex, 6).setValue(horaFormateada);

    // Obtener el color actual del semáforo
    var semaforoColor = semaforo.getBackground();


    if (semaforoColor === "yellow") {
        semaforo.setBackground("yellow");
    } 
    else if (semaforoColor === "green") {
        semaforo.setBackground("green");
    } 
    else if (!horaEntrada) {
        semaforo.setBackground("red");
    } 

    var asisAnualRange = sheetAsisAnual.getDataRange();
    var asisAnualValues = asisAnualRange.getValues();
    var asisRowIndex = -1;

    for (var j = 1; j < asisAnualValues.length; j++) {
        if (asisAnualValues[j][0] == nombre) {
            asisRowIndex = j + 1;
            break;
        }
    }

    if (asisRowIndex == -1) {
        asisRowIndex = sheetAsisAnual.getLastRow() + 1;
        sheetAsisAnual.getRange(asisRowIndex, 1).setValue(nombre);
    }

    var puntosASumar;
    var semaforoColor = semaforo.getBackground();

    if (semaforoColor === "yellow"){
      puntosASumar = 0.5;
    }else{
      puntosASumar = 1;
    }

    var mes = hora.getMonth();
    var columnaMes = mes + 2;
    var asistenciaActual = Number(sheetAsisAnual.getRange(asisRowIndex, columnaMes).getValue()) || 0;

    sheetAsisAnual.getRange(asisRowIndex, columnaMes).setValue(asistenciaActual + puntosASumar);

    var totalAsistencias = 0;
    for (var k = 2; k <= 13; k++) {
        totalAsistencias += Number(sheetAsisAnual.getRange(asisRowIndex, k).getValue()) || 0;
    }
    sheetAsisAnual.getRange(asisRowIndex, 14).setValue(totalAsistencias);

    return mostrarMensaje("¡Hasta luego!", `Gracias por asistir, <span class="highlight">${nombre}</span> (ID: ${id}).`, hora);
}
}

function enviarCorreoAlerta(destinatario, asunto, mensaje) {
    MailApp.sendEmail(destinatario, asunto, mensaje);
}

function mostrarMensaje(titulo, mensaje, hora) {
    var htmlOutput = `
    <html>
    <head>
        <style>
            body {
                font-family: 'Arial', sans-serif;
                background-color: #f4f4f9;
                color: #333;
                margin: 0;
                padding: 0;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
            }
            .container {
                background-color: #fff;
                border-radius: 10px;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                padding: 30px;
                width: 90%;
                max-width: 500px;
                text-align: center;
            }
            h1 {
                color: #4CAF50;
                font-size: 2em;
                margin-bottom: 20px;
            }
            .info {
                background-color: #f8f8f8;
                border-radius: 5px;
                padding: 15px;
                margin: 20px 0;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            .info p {
                margin: 10px 0;
                font-size: 1.1em;
                color: #555;
            }
            .footer {
                font-size: 0.9em;
                color: #888;
            }
            .highlight {
                color: #4CAF50;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>${titulo}</h1>
            <div class="info">
                <p><strong>Fecha:</strong> <span class="time">${hora.toLocaleString("es-MX", { timeZone: "America/Mexico_City" })}</span></p>
                <p>${mensaje}</p>
            </div>
            <div class="footer">
                <p>¡Gracias por Asistir!</p>
            </div>
        </div>
    </body>
    </html>
    `;
    return HtmlService.createHtmlOutput(htmlOutput);
}

function mostrarError(mensaje) {
    var htmlOutput = `
    <html>
    <head>
        <style>
            body {
                font-family: 'Arial', sans-serif;
                background-color: #f4f4f9;
                color: #333;
                margin: 0;
                padding: 0;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
            }
            .container {
                background-color: #fff;
                border-radius: 10px;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                padding: 30px;
                width: 90%;
                max-width: 500px;
                text-align: center;
            }
            h1 {
                color: #FF5722;
                font-size: 2em;
                margin-bottom: 20px;
            }
            .info {
                background-color: #f8f8f8;
                border-radius: 5px;
                padding: 15px;
                margin: 20px 0;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            .info p {
                margin: 10px 0;
                font-size: 1.1em;
                color: #555;
            }
            .footer {
                font-size: 0.9em;
                color: #888;
            }
            .highlight {
                color: #FF5722;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>¡Error!</h1>
            <div class="info">
                <p>${mensaje}</p>
            </div>
            <div class="footer">
                <p>Por favor, intenta nuevamente.</p>
            </div>
        </div>
    </body>
    </html>
    `;
    return HtmlService.createHtmlOutput(htmlOutput);
}

function mostrarRegistroCerrado(mensaje) {
    var htmlOutput = `
    <html>
    <head>
        <style>
            body {
                font-family: 'Arial', sans-serif;
                background-color: #f4f4f9;
                color: #333;
                margin: 0;
                padding: 0;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
            }
            .container {
                background-color: #fff;
                border-radius: 10px;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                padding: 30px;
                width: 90%;
                max-width: 500px;
                text-align: center;
            }
            h1 {
                color: #FF5722;
                font-size: 2em;
                margin-bottom: 20px;
            }
            .info {
                background-color: #f8f8f8;
                border-radius: 5px;
                padding: 15px;
                margin: 20px 0;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            .info p {
                margin: 10px 0;
                font-size: 1.1em;
                color: #555;
            }
            .footer {
                font-size: 0.9em;
                color: #888;
            }
            .highlight {
                color: #FF5722;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>¡Registro cerrado!</h1>
            <div class="info">
                <p>${mensaje}</p>
            </div>
            <div class="footer">
                <p>Por favor, vuelve durante el horario permitido.</p>
            </div>
        </div>
    </body>
    </html>
    `;
    return HtmlService.createHtmlOutput(htmlOutput);
}
function marcarFaltantes() {
    var spreadsheet = SpreadsheetApp.openById("1ZJYL8tzo9Kfw235Zdl2UR29mLjFj-pyHElUdJZ-dRHU");
    var sheetAsistencia = spreadsheet.getSheetByName("Asistencia");
    var sheetQR = spreadsheet.getSheetByName("QR");

    var asistenciaData = sheetAsistencia.getDataRange().getValues();
    var qrData = sheetQR.getDataRange().getValues();

    var idsRegistrados = asistenciaData.map(row => row[0]);

    for (var i = 1; i < qrData.length; i++) {
        var id = qrData[i][0];
        var nombre = qrData[i][1];

        var rowIndex = idsRegistrados.indexOf(id);

        if (rowIndex === -1) {
            var newRow = sheetAsistencia.getLastRow() + 1;
            sheetAsistencia.getRange(newRow, 1).setValue(id);
            sheetAsistencia.getRange(newRow, 2).setValue(nombre);

            var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
            sheetAsistencia.getRange(newRow, 4).setValue(fecha);
            sheetAsistencia.getRange(newRow, 7).setBackground("red");
        } else {
            var horaEntrada = asistenciaData[rowIndex][4];
            var horaSalida = asistenciaData[rowIndex][5];
            var semaforoCell = sheetAsistencia.getRange(rowIndex + 1, 7);

            if (horaEntrada && !horaSalida) {
                semaforoCell.setBackground("orange"); 
            } else if (!horaEntrada && !horaSalida) {
                semaforoCell.setBackground("red"); 
            }
        }
    }
}
