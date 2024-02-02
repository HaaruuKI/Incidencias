// leer.js

function leerExcel() {
    var fileInput = document.getElementById('excelFileInput');
  
    if (fileInput.files.length > 0) {
        var file = fileInput.files[0];
        var reader = new FileReader();
        document.getElementById('botonDescarga').style.display = 'block';
        reader.onload = function (e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, { type: 'binary' });
            var sheetName = 'doc 2';  // Nombre de la hoja que quieres leer
            var sheet = workbook.Sheets[sheetName];
            var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
    
            // Limpia la tabla antes de llenarla con nuevos datos
            limpiarTabla();
    
            // Llena los encabezados de la tabla (solo necesitas la columna A)
            llenarEncabezados(['Fecha', 'Hora', 'Codigo/Nombre']);
    
            // Llena los datos de la tabla con la columna A, excluyendo los índices 2, 4 y 6
            llenarDatos(jsonData.slice(1).map(function (fila) {
                // Divide la fila por punto y coma y devuelve un arreglo
                return fila[0].split(';').filter(function (_, indice) {
                    // Filtra los índices 2, 4 y 6
                    return indice !== 2 && indice !== 4 && indice !== 5 && indice !== 6 && indice !== 7;
                });
            }).filter(function (fila) {
                // Filtra las filas donde el cuarto elemento sea igual a 0
                return fila[2] !== '0';
            }));
        };
    
        reader.readAsBinaryString(file);
    } else {
        console.log('Por favor, selecciona un archivo Excel.');
    }
}

function limpiarTabla() {
    var tabla = document.getElementById('tablaRegistros');
    tabla.innerHTML = '<thead></thead><tbody></tbody>';
}

function llenarEncabezados(encabezados) {
    var thead = document.querySelector('#tablaRegistros thead');
  
    var tr = document.createElement('tr');
    encabezados.forEach(function (encabezado) {
        var th = document.createElement('th');
        th.textContent = encabezado;
        tr.appendChild(th);
    });
  
    thead.appendChild(tr);
}

function llenarDatos(datos) {
    var tbody = document.querySelector('#tablaRegistros tbody');
  
    datos.forEach(function (fila) {
        var tr = document.createElement('tr');
        fila.forEach(function (dato) {
            var td = document.createElement('td');
            td.textContent = dato;
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

function generarExcel() {
    // Obtener la tabla HTML
    var tabla = document.getElementById('tablaRegistros');

    // Convertir la tabla a un array de datos
    var datos = [];
    for (var i = 0; i < tabla.rows.length; i++) {
        var fila = [];
        for (var j = 0; j < tabla.rows[i].cells.length; j++) {
            fila.push(tabla.rows[i].cells[j].textContent);
        }
        datos.push(fila);
    }

    // Crear un nuevo libro de trabajo de Excel
    var workbook = XLSX.utils.book_new();

    // Agregar una hoja al libro de trabajo
    var worksheet = XLSX.utils.aoa_to_sheet(datos);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Asistencia");

    // Generar el archivo Excel
    var archivo = XLSX.write(workbook, { type: 'binary' });

    // Crear un enlace para descargar el archivo
    var enlaceDescarga = document.createElement('a');
    enlaceDescarga.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + btoa(archivo);
    enlaceDescarga.download = "asistencia.xlsx";
    enlaceDescarga.click();
}
