import XLSX from 'https://cdn.sheetjs.com/xlsx-0.20.0/package/xlsx.mjs';

//Agrega la informacion de un archivo excel utilizando la hoja seleccionada
const fileInput = document.getElementById('fileInput');
fileInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    const fileName = file.name;
    const fileExtension = fileName.slice((fileName.lastIndexOf(".")));
    console.log(fileExtension)

    const compatibleExtensions = [".xlsx", ".xls", ".ods"]; // Lista de extensiones compatibles

    if (compatibleExtensions.includes(fileExtension)) {
        const ab = await file.arrayBuffer(); // El metodo arrayBuffer() es una estructura de JS que te permite trabajar con datos binarios. En este caso como el archivo lo trae directamente desde el navegador el arraybuffer trabaja con los datos para poder parsearlos luego con el metodo de la libreria SheetJs XLSX.read()
        //Una vez seleccionado el archivo se realiza el parseo del arrayBuffer con el metodo XLSX.read() (en el caso de utilizar una ruta se utiliza el metodo XLSX.readFile(path))
        const wb = XLSX.read(ab);
        const wsnames = wb.SheetNames;// Array con los nombres de las hojas.;

        const selectContainer = document.getElementById("SelectContainer");
        const selectElement = document.createElement("select");
        selectElement.classList.add("form-select");
        selectElement.setAttribute("aria-label", "Default select example");

        const defaultOption = document.createElement("option");
        defaultOption.selected = true;
        defaultOption.textContent = "Elige una hoja";
        selectElement.appendChild(defaultOption);

        for (let i = 0; i < wsnames.length; i++) {
            const option = document.createElement("option");
            option.value = i; // Se agrega el value en la etiqueta para luego poder agregar segun la seleccion que hoja se quiere ver.
            option.textContent = wsnames[i]; // Usa el valor del array
            selectElement.appendChild(option);
        }

        selectContainer.appendChild(selectElement); // Agrega el select al contenedor deseado

        selectElement.addEventListener("change", () => {
            const selectValue = selectElement.value;
            const wsname = wb.SheetNames[selectValue];
            const ws = wb.Sheets[wsname];
            const tableContainer = document.getElementById("TableContainer");

            const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
            const headers = rows[3].slice(2, 7); // Se elijen los elementos para el encabezado de la tabla en este caso empieza en la fila 4 (del archivo excel).

            const tableHTML = `
                <table class="table table-dark table-hover">
                    <thead>
                        <tr>
                            ${headers.map(header => `<th scope="col">${header}</th>`).join("")}
                            <th scope="col"></th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rows.slice(4)
                    .filter(row => row.length > 4)
                    .map(row => row.slice(2, 7))// Aca se elijen los elementos del array que se van a mostrar en cada fila. Desde el elemento 2 hasta el elemento 7.
                    .map(element => `
                                  <tr>
                                      ${element.map(cell => `<td>${cell}</td>`).join("")}
                                      <td><button type="button" class="btn btn-outline-danger" id="addbtn" >Agregar</button>
                                      </td>
                                  </tr>
                              `)
                    .join("")}
                    </tbody>
                </table>
            `;

            tableContainer.innerHTML = tableHTML;
        });
    } else {
        // La extensión no es compatible con la librería SheetJS
        console.log("Archivo no compatible con SheetJS:", fileName);
    }
});



//Exporta una tabla a un archivo en formato xlsx
// document.getElementById("sheetjsexport").addEventListener('click', function () {
//     /* Create worksheet from HTML DOM TABLE */
//     const wb = XLSX.utils.table_to_book(document.getElementById("TableToExport"));
//     /* Export to file (start a download) */
//     XLSX.writeFile(wb, "SheetJSTable.xlsx");
// });


