import {
  AntidoteConnector,
  ConnectixAgent,
} from "@druide-informatique/antidote-api-js";

import * as utils from "./utils";
import { WordProcessorAgentOnlyOffice } from "./processor-agent/base";
import { WordProcessorAgentOnlyOfficeDocument } from "./processor-agent/document";
import { WordProcessorAgentOnlyOfficeDocumentSelection } from "./processor-agent/document-selection";
import { WordProcessorAgentOnlyOfficeUniversalSelection } from "./processor-agent/universal-selection";

((window, undefined) => {
  let isInitialized = false;
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
  let connectionErrorModalId: string | null;

  function getAntidotePort() {
    const antidotePort = localStorage.getItem("ANTIDOTE_PORT");
    if (antidotePort) {
      return Number(antidotePort);
    }

    throw new Error("Antidote port is not set.")
  }

  function getForceSetPort() {
    const forceSetPort = localStorage.getItem("FORCE_SET_PORT");
    if (forceSetPort === "true")
      return true;
    return false;
  }

  function launchCorrector() {
    AntidoteConnector.announcePresence();

    if (AntidoteConnector.isDetected()) {
      console.log("Antidote Connector is detected")
    }

    const agent = new ConnectixAgent(
      wordProcessorAgent!,
      (AntidoteConnector.isDetected() && !getForceSetPort())?
      AntidoteConnector.getWebSocketPort :
      async () => getAntidotePort()
    );

    agent.connectWithAntidote()
      .then(() => agent.launchCorrector())
      .catch(error => {
        const errorDialog = new window.Asc.PluginWindow();
        errorDialog.show(connectionErrorModal);
        connectionErrorModalId = errorDialog.id;

        console.log(error);
      })
  }

  window.Asc.plugin.init = (text: string) => {
    const alternativeText = (text.length === 0) ? null : text;

    if (wordProcessorAgent && wordProcessorAgent.isAvailable) {
      // On every selection change

      if (!wordProcessorAgent.updatingByAntidote) {
        if (wordProcessorAgent instanceof WordProcessorAgentOnlyOfficeDocumentSelection) {
          setTimeout(() => {
            if (wordProcessorAgent && !wordProcessorAgent.updatingByAntidote) {
              wordProcessorAgent.updateText();
            }
          }, 200);
        } else if (wordProcessorAgent instanceof WordProcessorAgentOnlyOfficeUniversalSelection) {
          setTimeout(() => {
            (wordProcessorAgent as WordProcessorAgentOnlyOfficeUniversalSelection).setAlternativeText(alternativeText);
            if (wordProcessorAgent && !wordProcessorAgent.updatingByAntidote) {
              wordProcessorAgent.updateText();
            }
          }, 200);
        }
      }
    } else {
      // Otherwise, create an WordProcessorAgent instance
      let promise: Promise<void> | null = null;
      switch (window.Asc.plugin.info.editorType) {
        case "word":
          promise = utils.callCommand(
            window.Asc,
            () => {
              const oDocument = Api.GetDocument();
              const oDocumentInfo = oDocument.GetDocumentInfo();
              const title = oDocumentInfo.Title;

              const oRange = oDocument.GetRangeBySelect();
              const start = oRange ? oRange.GetStartPos() : null;
              const end = oRange ? oRange.GetEndPos() : null;

              const hasSelection = (start !== end);

              return { title, hasSelection };
            },
            false,
            false,
          )
            .then(async ({ title, hasSelection }) => {
              if (hasSelection) {
                wordProcessorAgent = new WordProcessorAgentOnlyOfficeDocumentSelection(window.Asc, title);
              } else {
                wordProcessorAgent = new WordProcessorAgentOnlyOfficeDocument(window.Asc, title);
              }
            });
          break;
        case "slide":
          promise = utils.callCommand(
            window.Asc,
            () => {
              const oPresentation = Api.GetPresentation();
              const oDocumentInfo = oPresentation.GetDocumentInfo();
              const title = oDocumentInfo.Title;

              return title;
            },
            false,
            false
          )
            .then(title => {
              wordProcessorAgent = new WordProcessorAgentOnlyOfficeUniversalSelection(window.Asc, title);
              (wordProcessorAgent as WordProcessorAgentOnlyOfficeUniversalSelection)
                .setAlternativeText(alternativeText);
            });
          break;
        case "cell":
          promise = utils.callCommand(
            window.Asc,
            () => {
              const oDocumentInfo = Api.GetDocumentInfo();
              const title = oDocumentInfo.Title;

              return title;
            },
            false,
            false
          )
            .then(title => {
              wordProcessorAgent = new WordProcessorAgentOnlyOfficeUniversalSelection(window.Asc, title);
              (wordProcessorAgent as WordProcessorAgentOnlyOfficeUniversalSelection)
                .setAlternativeText(alternativeText);
            });
          break;
      }

      if (promise) {
        promise.then(() => wordProcessorAgent!.updateText()).then(launchCorrector);
      }
    }

    isInitialized = true;
  };

  window.Asc.plugin.button = (id: string, windowId: string) => {
    if (connectionErrorModalId && windowId === connectionErrorModalId) {
      window.Asc.plugin.executeCommand("close", "");
    }
  };

})(window, undefined);
