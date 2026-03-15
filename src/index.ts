import {
  AntidoteConnector,
  ConnectixAgent,
} from "@druide-informatique/antidote-api-js";

import * as utils from "./utils";
import { WordProcessorAgentOnlyOffice } from "./processor-agent/base";
import { WordProcessorAgentOnlyOfficeDocument } from "./processor-agent/document";
// import { Range, WordProcessorAgentOnlyOfficeSelection } from "./processor-agent/selection";

((window, undefined) => {
  let wordProcessorAgent: WordProcessorAgentOnlyOffice | null;

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
      wordProcessorAgent!,
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
    utils.callCommand(
      window.Asc,
      () => {
        const oDocument = Api.GetDocument();
        const oDocumentInfo = oDocument.GetDocumentInfo();
        const title = oDocumentInfo.Title;

        const oRange = oDocument.GetRangeBySelect();
        const range = oRange ? {
          start: oRange.GetStartPos(),
          end: oRange.GetEndPos()
        } : null;

        return { title, range };
      }
    )
    .then(async res => {
      const { title, range } = res;
      // if (range) {
        // wordProcessorAgent = new WordProcessorAgentOnlyOfficeSelection(window.Asc, title, range);
      // } else {
        wordProcessorAgent = new WordProcessorAgentOnlyOfficeDocument(window.Asc, title);
      // }
      await wordProcessorAgent.updateText();
    })
    .then(() => {
      firstLoad = false;
      launchCorrector();
    });

    window.Asc.plugin.attachEditorEvent("onParagraphText", (data: any) => {
      // console.log("The not firstLoad: ", !firstLoad);
      // console.log("The expression: ", !firstLoad && wordProcessorAgent
      //   && !wordProcessorAgent.updatingByAntidote);
      if (!firstLoad && wordProcessorAgent
        && !wordProcessorAgent.updatingByAntidote) {

        // Check if currently the text is updated by Antidote,
        // if not, wait sometime and then recheck to ensure that the
        // replacingQueue is empty
        setTimeout(() => {
          if (!firstLoad && wordProcessorAgent
            && !wordProcessorAgent.updatingByAntidote) {
            // console.log("From onParagraphText", data)
            wordProcessorAgent!.updateText();
          }
        }, 100);
      }
    });
  };

  window.Asc.plugin.button = (id: string, windowId: string) => {
    if (windowId === window.Asc.plugin.connectionErrorModalId) {
      window.Asc.plugin.executeCommand("close", "");
    }
  };

})(window, undefined);
