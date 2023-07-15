import { ReadInfo, Enab } from "./modules/H4OInfo.js";
import { showErrorDialog } from "./modules/H4O.js";

function getDecimalSeparator() {
  const n = 1.1;
  return n.toLocaleString().substring(1, 2);
}

const IsValidPort = (port) => {
  if (port.trim().length === 0) {
    return false;
  } else if (isNaN(port)) {
    return false;
  } else if (parseInt(port) !== parseFloat(port)) {
    return false;
  } else if (parseInt(port) < 0 || parseInt(port) > 65535) {
    return false;
  }

  return true;
};

const cmdConnect = (host, port) => {
  console.log("Valor de txtHost:", host);
  console.log("Valor de txtPort:", port);
  if (txtHost.value.trim().length === 0) {
    console.log("Host field is empty");
    showErrorDialog("Host field is empty");
    return false;
  } else if (txtPort.value.trim().length === 0) {
    console.log("Port field is empty");

    showErrorDialog("Port field is empty");
    return false;
  } else if (!IsValidPort(txtPort.value)) {
    console.log("The port field is not a number between 0 and 65535");

    showErrorDialog("The port field is not a number between 0 and 65535");
    return false;
  } else {
    ReadInfo(txtHost.value.trim(), txtPort.value.trim());
    return true;
  }
};

Office.onReady()
  .then(function () {
    if (Office.context.host === Office.HostType.Excel) {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("run").onclick = run;
      document.getElementById("enableButton").onclick = runDisabled;
    }
  })
  .catch(function (error) {
    console.log(error);
  });

export async function run() {
  try {
    await Excel.run(async (context) => {
      const txtHostInput = document.getElementById("txtHost");
      const txtPortInput = document.getElementById("txtPort");
      const txtHostValue = txtHostInput.value;
      const txtPortValue = txtPortInput.value;
      const res = cmdConnect(txtHostValue, txtPortValue);
      localStorage.setItem("host", txtHost.value.trim());
      localStorage.setItem("port", txtPort.value.trim());

      console.log("CondiciónHost: ", txtHost.value.trim());
      console.log("CondiciónPort: ", txtPort.value.trim());
      console.log("Res: ", res);
      if (res) {
        Office.ribbon.requestUpdate({
          tabs: [
            {
              id: "H4O.Tab",
              groups: [
                {
                  id: "CommandsGroup2",
                  controls: [
                    {
                      id: "TaskpaneButton2",
                      enabled: true,
                    },
                  ],
                },
              ],
            },
          ],
        });
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function runDisabled() {
  try {
    await Excel.run(async (context) => {
      Office.ribbon.requestUpdate({
        tabs: [
          {
            id: "H4O.Tab",
            groups: [
              {
                id: "CommandsGroup2",
                controls: [
                  {
                    id: "TaskpaneButton2",
                    enabled: true,
                  },
                ],
              },
            ],
          },
        ],
      });
      await context.sync();
    });
  } catch {}
}
