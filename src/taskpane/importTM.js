import { ReadInfo, getParams, getLevels, getSamples } from "./modules/H4OInfo";
var select = document.getElementById("satellite");

ReadInfo()
  .then((satellites) => {
    console.log("satellites en el otro", satellites);
    satellites.forEach(function (satellite) {
      var option = document.createElement("option");
      option.text = satellite;

      select.add(option);
    });
  })
  .catch((error) => {
    console.error(error);
  });

// Obtén las referencias de los elementos de fecha
const startInput = document.getElementById("startDate");
const endInput = document.getElementById("endDate");

// Obtiene la fecha actual
const today = new Date();
const todayFormatted = today.toISOString().slice(0, 10);

// Establece los valores de los campos de fecha
startInput.value = todayFormatted + "T00:00";
endInput.value = todayFormatted + "T00:00";

// Función para crear un elemento de lista
const createListItem = (text) => {
  const li = document.createElement("li");
  li.textContent = text;
  return li;
};

// Función para mover un elemento de una lista a otra
const moveItem = (sourceList, targetList) => {
  const selectedItems = Array.from(sourceList.querySelectorAll("li.selected"));
  selectedItems.forEach((item) => {
    item.classList.remove("selected");
    targetList.appendChild(item);
  });

  sortList(sourceList);
};

// Función para ordenar una lista de elementos <li>
const sortList = (list) => {
  const items = Array.from(list.getElementsByTagName("li"));
  items.sort((a, b) => a.textContent.localeCompare(b.textContent));
  list.innerHTML = ""; // Limpiar la lista
  items.forEach((item) => list.appendChild(item));
};

// Cargar los datos desde la API y agregarlos a la lista de origen
const loadSourceList = (satelliteName, filterText) => {
  const sourceList = document.getElementById("sourceList");
  sourceList.innerHTML = ""; // Limpiar la lista anterior
  getParams(satelliteName)
    .then((params) => {
      params.forEach((param) => {
        if (!filterText || param.toLowerCase().includes(filterText.toLowerCase())) {
          const li = createListItem(param);
          li.addEventListener("click", () => {
            li.classList.toggle("selected");
          });
          sourceList.appendChild(li);
        }
      });
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};

const loadLevels = (satelliteName) => {
  const levelSelect = document.getElementById("level");
  levelSelect.innerHTML = ""; // Limpiar las opciones anteriores

  getLevels(satelliteName)
    .then((levels) => {
      levels.forEach((levelText) => {
        const option = document.createElement("option");
        option.value = levelText;
        option.textContent = levelText;
        levelSelect.appendChild(option);
      });
    })
    .catch((error) => {
      console.error("Error:", error);
    });
};

// Evento al hacer clic en el botón para mover a la derecha
document.getElementById("moveRight").addEventListener("click", () => {
  const sourceList = document.getElementById("sourceList");
  const targetList = document.getElementById("targetList");
  moveItem(sourceList, targetList);
  sortList(sourceList);
});

// Evento al hacer clic en el botón para mover a la izquierda
document.getElementById("moveLeft").addEventListener("click", () => {
  const sourceList = document.getElementById("sourceList");
  const targetList = document.getElementById("targetList");
  moveItem(targetList, sourceList);
  sortList(sourceList);
});

// Evento al cambiar el valor del select
document.getElementById("satellite").addEventListener("change", (event) => {
  const satelliteName = event.target.value;
  const filterText = document.getElementById("paramSelect").value;
  loadSourceList(satelliteName, filterText);
  loadLevels(satelliteName);
  // Limpiar la lista de destino
  const targetList = document.getElementById("targetList");
  targetList.innerHTML = "";
});

// Evento al escribir en el input de filtro
document.getElementById("paramSelect").addEventListener("input", (event) => {
  const filterText = event.target.value;
  const satelliteName = document.getElementById("satellite").value;
  loadSourceList(satelliteName, filterText);
});

// Cargar la lista de origen al cargar la página con el valor inicial del select
const initialSatelliteName = document.getElementById("satellite").value;
loadSourceList(initialSatelliteName, "");

let num = 0;

// Agregar un controlador de eventos al botón
document.querySelector(".import-button").addEventListener("click", function () {
  // Obtener los valores de los campos del formulario
  const satellite = document.getElementById("satellite").value;
  const level = document.getElementById("level").value.charAt(0);
  const start = formatDate(document.getElementById("startDate").value);
  const end = formatDate(document.getElementById("endDate").value);
  const showHeader = document.getElementById("headerCheckbox").checked;
  console.log("Nivel: ", level);
  // Obtener los elementos de la lista "listSelected"
  const listSelectedItems = document.querySelectorAll("#targetList li");
  const listSelectedValues = Array.from(listSelectedItems)
    .map((item) => item.textContent)
    .join(" ");

  // Crear un objeto con los valores recolectados
  const formData = {
    satellite: satellite,
    level: level,
    start: start,
    end: end,
    showHeader: showHeader,
    listSelected: listSelectedValues,
  };

  console.log("ShowHeader: ", showHeader);
  // Lógica adicional para procesar los datos o enviarlos a un servidor
  //setup();
  if (level === "1") {
    writeToExcelLevel1(formData, num, showHeader);
    num++;
  } else {
    writeToExcel(formData, num, showHeader);
    num++;
  }
});

// Función para formatear la fecha
const formatDate = (dateString) => {
  const date = new Date(dateString);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
};

const writeToExcelLevel1 = (formData, num, showHeader) => {
  Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let selectedRange = context.workbook.getSelectedRange();
    selectedRange.load("address, columnIndex, rowIndex");

    return context.sync().then(function () {
      var address = selectedRange.address;
      return new Promise(function (resolve, reject) {
        getSamples(formData)
          .then(function (samples) {
            if (showHeader) {
              const dataPairs = [
                ["Satellite", formData.satellite],
                ["Level", formData.level],
                ["Result", samples.Data.length.toString()],
                ["Parameters", formData.listSelected.toString()],
                ["Start time", formatDate(formData.start).toString()],
                ["End time", formatDate(formData.end).toString()],
                ["Import time", formatDate(new Date().toLocaleString()).toString()],
              ];

              var startCell = sheet.getRange(address);
              let row;
              for (row = 0; row < 7; row++) {
                for (let col = 0; col < 2; col++) {
                  startCell.getCell(row, col).clear(); // Limpiar el formato de la celda
                  startCell.getCell(row, col).values = [[dataPairs[row][col].toString()]];
                  if (col === 0) {
                    startCell.getCell(row, col).format.font.bold = true;
                  }
                  startCell.getCell(row, col).format.horizontalAlignment = "Center";
                  startCell.getCell(row, col).format;
                  startCell.getCell(row, col).format.autofitColumns();
                  startCell.getCell(row, col).format.autofitRows();
                }
              }

              selectedRange = startCell.getCell(row + 1, 0);
            }
            const tableRange = selectedRange.getAbsoluteResizedRange(samples.Data.length + 1, 6);
            let table = sheet.tables.add(tableRange, true);
            table.name = "DataSheet" + num;

            const headerRange = table.getHeaderRowRange();
            headerRange.values = [samples.ParameterName];

            const dataRange = table.getDataBodyRange();
            dataRange.values = samples.Data.map((item) => [
              item.Name,
              item.Timestamp,
              item.Raw,
              item.Eng,
              item.Validity,
              item.OOL,
            ]);

            // table.format.horizontalAlignment = "Center";

            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();

            // Obtener las columnas de la tabla
            const dataColumns = table.columns;
            context.load(dataColumns, "items");

            resolve();

            return context.sync().then(function () {
              if (dataColumns.items.length > 0) {
                // Centrar texto en las columnas
                dataColumns.items.forEach(function (column) {
                  column.getRange().format.horizontalAlignment = "Center";
                });
                return context.sync();
              } else {
                // No hay columnas en la tabla, manejar el caso apropiadamente
                console.log("No se encontraron columnas en la tabla.");
              }
            });
          })
          .catch(function (error) {
            reject(error);
          });
      });
    });
  });
};

const writeToExcel = (formData, num, showHeader) => {
  Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let selectedRange = context.workbook.getSelectedRange();
    selectedRange.load("address, columnIndex, rowIndex");

    return context.sync().then(function () {
      var address = selectedRange.address;
      return new Promise(function (resolve, reject) {
        getSamples(formData)
          .then(function (samples) {
            if (showHeader) {
              const dataPairs = [
                ["Satellite", formData.satellite],
                ["Level", formData.level],
                ["Result", samples.Data.length.toString()],
                ["Parameters", formData.listSelected.toString()],
                ["Start time", formatDate(formData.start).toString()],
                ["End time", formatDate(formData.end).toString()],
                ["Import time", formatDate(new Date().toLocaleString()).toString()],
              ];

              var startCell = sheet.getRange(address);
              let row;
              for (row = 0; row < 7; row++) {
                for (let col = 0; col < 2; col++) {
                  startCell.getCell(row, col).clear(); // Limpiar el formato de la celda
                  startCell.getCell(row, col).values = [[dataPairs[row][col].toString()]];
                  if (col === 0) {
                    startCell.getCell(row, col).format.font.bold = true;
                  }
                  startCell.getCell(row, col).format.horizontalAlignment = "Center";
                  startCell.getCell(row, col).format.autofitColumns();
                  startCell.getCell(row, col).format.autofitRows();
                }
              }

              selectedRange = startCell.getCell(row + 1, 0);
            }

            const tableRange = selectedRange.getAbsoluteResizedRange(samples.Data.length + 1, 8);
            let table = sheet.tables.add(tableRange, true);
            table.name = "DataSheet" + num;

            const headerRange = table.getHeaderRowRange();
            headerRange.values = [samples.ParameterName];

            const dataRange = table.getDataBodyRange();
            dataRange.values = samples.Data.map((item) => [
              item.Name,
              item.Start,
              item.End,
              item.Mean,
              item.Max,
              item.Min,
              item.Validity,
              item.StdDev,
            ]);

            //table.columns.format.horizontalAlignment = "Center";

            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();

            resolve();

            // Obtener las columnas de la tabla
            const dataColumns = table.columns;
            context.load(dataColumns, "items");
            return context.sync().then(function () {
              if (dataColumns.items.length > 0) {
                // Centrar texto en las columnas
                dataColumns.items.forEach(function (column) {
                  column.getRange().format.horizontalAlignment = "Center";
                });
                return context.sync();
              } else {
                // No hay columnas en la tabla, manejar el caso apropiadamente
                console.log("No se encontraron columnas en la tabla.");
              }
            });
          })
          .catch(function (error) {
            reject(error);
          });
      });
    });
  });
};
