import { SatellitesQuery, ParametersQuery, LevelsQuery, BuildServiceAddress, SamplesQuery } from "./H4OSvc";

export const ReadInfo = () => {
  return new Promise((resolve, reject) => {
    const satellites = [];
    try {
      const headers = {
        "X-Requested-With": "XMLHttpRequest",
      };
      const hostLS = localStorage.getItem("host");
      const portLS = localStorage.getItem("port");
      console.log("hostLS", hostLS);
      console.log("portLS", portLS);
      fetch(SatellitesQuery(hostLS, portLS), { headers })
        .then((response) => response.text())
        .then((html) => {
          const tempElement = document.createElement("div");
          tempElement.innerHTML = html;
          const satelliteRows = tempElement.querySelectorAll("table tr");
          const sats = Array.from(satelliteRows).map((row) => row.querySelector("td").textContent);
          sats.forEach((sat) => {
            satellites.push(sat);
          });
          Office.initialize = function (reason) {
            console.log("Office.initialize");
            // El código aquí se ejecutará cuando Office.js esté inicializado

            if (reason === Office.InitializationReason.DocumentOpened) {
              // Este código se ejecutará cuando se abra un documento de Excel

              // Aquí puedes llamar a tu función enableGetTMbutton() u otras operaciones relacionadas con el complemento
              enableGetTMbutton();
            }
          };

          resolve(satellites);
        })
        .catch((error) => {
          console.error("Error:", error);
          reject(error);
        });
    } catch (error) {
      console.error("Error " + error.code + ": " + error.message);
      reject(error);
    }
  });
};
// const host = "demo-swarm"
// const port = 7000

export const getParams = (satelliteName) => {
  return new Promise((resolve, reject) => {
    try {
      const hostLS = localStorage.getItem("host");
      const portLS = localStorage.getItem("port");
      fetch(ParametersQuery(hostLS, portLS, satelliteName))
        .then((response) => response.text())
        .then((html) => {
          const params = parseParamsFromHTML(html);

          resolve(params);
        })
        .catch((error) => {
          console.error("Error:", error);
          reject(error);
        });
    } catch (error) {
      console.error("Error:", error);
      reject(error);
    }
  });
};

// Función para obtener los levels de la API
export const getLevels = (satelliteName) => {
  return new Promise((resolve, reject) => {
    try {
      const hostLS = localStorage.getItem("host");
      const portLS = localStorage.getItem("port");
      fetch(LevelsQuery(hostLS, portLS, satelliteName))
        .then((response) => response.text())
        .then((html) => {
          const levels = parseLevelsFromHTML(html);
          resolve(levels);
        })
        .catch((error) => {
          console.error("Error:", error);
          reject(error);
        });
    } catch (error) {
      console.error("Error:", error);
      reject(error);
    }
  });
};

// Función para obtener los niveles de la API
export const getSamples = (formData) => {
  return new Promise((resolve, reject) => {
    try {
      const hostLS = localStorage.getItem("host");
      const portLS = localStorage.getItem("port");
      fetch(SamplesQuery(hostLS, portLS, formData))
        .then((response) => response.text())
        .then((html) => {
          const samples = parseSamplesFromHtml(html);
          console.log("Samples: ", samples);
          resolve(samples);
        })
        .catch((error) => {
          console.error("Error:", error);
          reject(error);
        });
    } catch (error) {
      console.error("Error:", error);
      reject(error);
    }
  });
};


const parseSamplesFromHtml = (html) => {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");

  // // 1. Obtener los headers de la primera tabla
  // const headerRows = doc.querySelectorAll("table:nth-of-type(1) tr");
  // const headers = Array.from(headerRows).map((row) => row.querySelector("th").textContent);

  // 2. Obtener los nombres de parámetros del primer tr de la segunda tabla
  const parameterRow = doc.querySelector("table:nth-of-type(2) tr:first-of-type");
  const parameterNameCells = parameterRow.querySelectorAll("th");
  const parameterNames = Array.from(parameterNameCells).map((cell) => cell.textContent);

  // 3. Obtener los datos correspondientes a cada nombre de parámetro
  const dataRows = doc.querySelectorAll("table:nth-of-type(2) tr:not(:first-of-type)");
  const data = Array.from(dataRows).map((row) => {
    const cells = row.querySelectorAll("td");
    return Array.from(cells).reduce((obj, cell, index) => {
      const parameterName = parameterNames[index];
      obj[parameterName] = cell.textContent.trim();
      return obj;
    }, {});
  });

  return {
    // Headers: headers,
    ParameterName: parameterNames,
    Data: data,
  };
};

// const parseSamplesFromHtml = (html) => {
//   const headers = [];
//   const data = [];

//   const tempElement = document.createElement("div");
//   tempElement.innerHTML = html;

//   const headerCells = tempElement.querySelectorAll("table:first-of-type th");
//   headerCells.forEach((cell) => {
//     headers.push(cell.textContent);
//   });

//   const dataRows = tempElement.querySelectorAll("table:last-of-type tr:not(:first-child)");
//   dataRows.forEach((row) => {
//     const rowData = [];
//     const cells = row.querySelectorAll("td");
//     cells.forEach((cell) => {
//       rowData.push(cell.textContent);
//     });
//     data.push(rowData);
//   });

//   return { headers, data };
// };

const parseParamsFromHTML = (html) => {
  const params = [];
  const tempElement = document.createElement("div");
  tempElement.innerHTML = html;
  const paramRows = tempElement.querySelectorAll("table tr");

  for (let i = 0; i < paramRows.length; i++) {
    const paramData = paramRows[i].querySelector("td");
    if (paramData) {
      const param = paramData.textContent;
      params.push(param);
    }
  }

  return params;
};

const parseLevelsFromHTML = (html) => {
  const levels = [];
  const tempElement = document.createElement("div");
  tempElement.innerHTML = html;
  const levelRows = tempElement.querySelectorAll("table tr");

  for (let i = 0; i < levelRows.length; i++) {
    const levelData = levelRows[i].querySelectorAll("td");
    if (levelData && levelData.length === 2) {
      const level = levelData[0].textContent;
      const levelName = levelData[1].textContent;
      const levelText = `${level}   -   ${levelName}`;
      levels.push(levelText);
    }
  }

  return levels;
};

const enableGetTMbutton = async () => {
  console.log("enableGetTMbutton");

  await Office.onReady();

  if (Office.context.mailbox) {
    const runtimeId = Office.context.mailbox.diagnostics.hostName;
    const reasons = Office.Runtime.UpdateReason.Registered;

    Office.runtime.requestUpdate({
      addinId: runtimeId,
      reasons: reasons,
    });

    const button = { id: "TaskpaneButton2", enabled: true };
    const parentGroup = { id: "CommandsGroup2", controls: [button] };
    const parentTab = { id: "H4O.Tab", groups: [parentGroup] };
    const ribbonUpdater = { tabs: [parentTab] };
    console.log(ribbonUpdater);
    Office.ribbon.requestUpdate(ribbonUpdater);
  } else {
    console.log("Office.context is undefined.");
  }
};




//for (sat of sats.resultRange.cells.values) {
//   // Put the parameters list for a satellite in the meta info worksheet
//   if (sat[0] !== "") {
//     params = info.queryTables.add("URL;" + ParametersQuery(host, port, sat[0]), info.getRange("D1"));
//     params.name = ParamsTableName(sat[0]);
//     params.backgroundQuery = false;
//     params.refresh();

//     // Put the levels list for a satellite in the meta info worksheet
//     levels = info.queryTables.add("URL;" + LevelsQuery(host, port, sat[0]), info.getRange("D1"));
//     levels.name = LevelsTableName(sat[0]);
//     levels.backgroundQuery = false;
//     levels.refresh();
//   }
// }

// Everything is ok until now. Replace the old worksheet with the new.
// ForceDeleteWorksheet(INFO_SHEET_NAME);
// info.name = INFO_SHEET_NAME;
