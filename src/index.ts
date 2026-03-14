import {
  AntidoteConnector,
  ConnectixAgent,
} from "@druide-informatique/antidote-api-js";
import { WordProcessorAgentOnlyOfficeDocument } from "./processor-agent";

(function(window, undefined){
  const wordProcessorAgent = new WordProcessorAgentOnlyOfficeDocument(window.Asc);

  function getFullUrl(name: string): string {
    const location = window.location;
    const start = location.pathname.lastIndexOf("/") + 1;
    const file = location.pathname.slice(start);
    return location.href.replace(file, name);
  }

  const connectionErrorModal = {
    url: getFullUrl("connection-error.html"),  // Same HTML as config variationnt
    description: window.Asc.plugin.tr("Error"),
    isVisual: true,
    EditorsSupport: ["word"],
    isModal : true,
    isInsideMode : false,
    initDataType : "none",
    initData : "",
    size: [350, 150],
    buttons: [
      {
        text: window.Asc.plugin.tr("Close"),
        primary: true
      }
    ]
  };

  function getAntidotePort() {
    const antidotePort = localStorage.getItem("ANTIDOTE_PORT");
    if (antidotePort) {
      return Number(antidotePort);
    }

    throw new Error("Antidote port is not set.")
  }

  function launchCorrector() {
    AntidoteConnector.announcePresence();
    const agent = new ConnectixAgent(
      wordProcessorAgent,
      AntidoteConnector.isDetected() ?
      AntidoteConnector.getWebSocketPort :
      async () => getAntidotePort()
    );

    agent.connectWithAntidote()
      .then(() => agent.launchCorrector())
      .catch(error => {
        const errorDialog = new window.Asc.PluginWindow();
        errorDialog.show(connectionErrorModal);
        window.Asc.plugin.connectionErrorModalId = errorDialog.id;
      })
  }

  window.Asc.plugin.init = () => {
    let firstLoad = true;
    wordProcessorAgent.updateParagraphs()
      .then(() => {
        firstLoad = false;
        launchCorrector();
      });

    window.Asc.plugin.attachEditorEvent("onParagraphText", (data: any) => {
      if (!wordProcessorAgent.updatingByAntidote && !firstLoad) {
        console.log("From onParagraphText", data)
        wordProcessorAgent.updateParagraphs();
      }
    });
  };

  window.Asc.plugin.button = (id: string, windowId: string) => {
    if (windowId === window.Asc.plugin.connectionErrorModalId) {
      window.Asc.plugin.executeCommand("close", "");
    }
  };

})(window, undefined);
