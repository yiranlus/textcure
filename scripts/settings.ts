import { applyTranslation } from "./utils";

((window, undefined) => {
  let inputAntidotePort: HTMLInputElement | null;
  let inputForceSetPort: HTMLInputElement | null;

  window.Asc.plugin.init = function () {
    inputAntidotePort = document.getElementById("antidotePort") as HTMLInputElement;
    inputForceSetPort = document.getElementById("forceSetPort") as HTMLInputElement;

    if (inputAntidotePort) {
      const antidotePort = localStorage.getItem("ANTIDOTE_PORT");
      if (antidotePort)
        inputAntidotePort.value = antidotePort;
    }
    if (inputForceSetPort) {
      const forceSetPort = localStorage.getItem("FORCE_SET_PORT");
      if (forceSetPort === "true")
        inputForceSetPort.checked = true;
    }
  };

  window.Asc.plugin.button = (id: string, windowId: string) => {
    const antidotePort = Number(inputAntidotePort?.value);
    const forceSetPort = inputForceSetPort?.checked;

    // Send value back to main plugin context (optional)
    localStorage.setItem("ANTIDOTE_PORT", antidotePort.toString());
    localStorage.setItem("FORCE_SET_PORT", forceSetPort?"true":"false");

    window.Asc.plugin.executeCommand("close", "");
  };

  window.Asc.plugin.onTranslate = () => {
    applyTranslation(window.Asc, "lblAntidotePort", "Websocket Port:");
    applyTranslation(window.Asc, "lblForceSetPort", "Ignore Antidote Connector");
  }
})(window, undefined);
