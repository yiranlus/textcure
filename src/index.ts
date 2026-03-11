import {
  WordProcessorAgent,
  AntidoteConnector,
  ConnectixAgent,
  ParamsReplace,
  ParamsAllowEdit,
  WordProcessorConfiguration,
  ParamsGetZonesToCorrect,
  TextZoneConnectix,
  DocumentType
} from "@druide-informatique/antidote-api-js";

type Segment = {
  relPos: number,
  text: string
}

type Paragraph = {
  globalPos: number,
  segments: Segment[]
}

type SegmentIndex = {
  par: number,
  seg: number
}

type TextRange = {
  start: number,
  end: number
}

class WordProcessorAgentOnlyOfficeDocument extends WordProcessorAgent {
  title: string;
  Asc: any;
  updateByAntidote: boolean;
  paragraphs: Paragraph[] | null;

  constructor(Asc: any) {
    console.log("created WordProcessorAgent");
    super();
    this.Asc = Asc;
    this.updateByAntidote = false;

    this.title = "";
    this.paragraphs = null;
  }

  sessionEnded() {
    this.Asc.plugin.executeCommand("close", "");
    super.sessionEnded();
  }

  findIndex(pos: number, eager: boolean=false): SegmentIndex {
    if (!this.paragraphs) {
      throw new Error("No text found");
    }

    let i = 0, j = 0;
    if (eager) {
      while (
        i + 1 < this.paragraphs.length &&
        this.paragraphs[i + 1].globalPos <= pos)
          i++;
        while (
          j + 1 < this.paragraphs[i].segments.length &&
          this.paragraphs[i].segments[j+1].relPos <= pos - this.paragraphs[i].globalPos)
        j++;
    } else {
      while (
        i + 1 < this.paragraphs.length &&
        this.paragraphs[i + 1].globalPos < pos)
      i++;
      while (
        j + 1 < this.paragraphs[i].segments.length &&
        this.paragraphs[i].segments[j+1].relPos < pos - this.paragraphs[i].globalPos)
      j++;
    }

    return {
      par: i,
      seg: j,
    }
  }

  correctIntoWordProcessor(params: ParamsReplace): boolean {
    if (!this.paragraphs) {
      throw new Error("No text found");
    }

    this.updateByAntidote = true;

    let segmentIndex = this.findIndex(params.positionStartReplace, true);
    let i = segmentIndex.par, j = segmentIndex.seg;

    let textElement = this.paragraphs[segmentIndex.par].segments[segmentIndex.seg];
    let globalPos = this.paragraphs[segmentIndex.par].globalPos;
    let relPos = this.paragraphs[segmentIndex.par].segments[segmentIndex.seg].relPos;
    let textRange: TextRange = {
      start: params.positionStartReplace - globalPos - relPos,
      end: params.positionReplaceEnd - globalPos - relPos,
    }

    this.Asc.scope.paramsReplace = {
      segmentIndex,
      textRange,
      newString: params.newString
    }
    // console.log("paramsReplace: ", this.Asc.scope.paramsReplace);

    try {
      this.Asc.plugin.callCommand(() => {
        var segmentIndex: SegmentIndex = Asc.scope.paramsReplace.segmentIndex;
        var textRange: TextRange = Asc.scope.paramsReplace.textRange;
        var newString = Asc.scope.paramsReplace.newString;

        var oDocument = Api.GetDocument();
        var oElement = oDocument.GetElement(segmentIndex.par).GetElement(segmentIndex.seg);
        var oRange = oElement.GetRange(textRange.start, textRange.end);

        // console.log(`the Run: "${oElement.GetText()}"`)
        // console.log(`Text in the range: "${oRange.GetText()}"`);
        // console.log(`newText in the range: "${newString}"`);

        oRange.Delete();
        oRange.AddText(newString);
        return {
          text: oElement.GetText(),
          diff: newString.length - (textRange.end - textRange.start)
        }
      },
      false,
      true,
      (res: {text: string, diff: number}) => {

        if (!this.paragraphs) {
          throw new Error("No text found");
        }

        this.paragraphs[segmentIndex.par].segments[j].text = res.text;
        for (let j = segmentIndex.seg + 1; j < this.paragraphs[segmentIndex.par].segments.length; j++) {
          this.paragraphs[segmentIndex.par].segments[j].relPos += res.diff;
        }
        for (let i = segmentIndex.par + 1; i < this.paragraphs.length; i++) {
          this.paragraphs[i].globalPos += res.diff;
        }
      });
    } catch (error) {
      console.log("error: ", error);
      return false;
    }
    this.updateByAntidote = false;

    return true;
  }

  configuration(): WordProcessorConfiguration {
    return {
      documentTitle: this.title,
      carriageReturn: "\r\n",
      activeMarkup: DocumentType.text
    };
  }

  allowEdit(params: ParamsAllowEdit): boolean {
    let indexStart = this.findIndex(params.positionStart, true);
    let indexEnd = this.findIndex(params.positionEnd);

    // console.log("params: ", params);
    // console.log("Index: ", indexStart, indexEnd);

    return (
      indexStart.par === indexEnd.par &&
      indexStart.seg === indexEnd.seg
    );
  }

  textZonesAvailable(): boolean {
    return !!this.paragraphs;
  }

  zonesToCorrect(_params: ParamsGetZonesToCorrect): TextZoneConnectix[] {
    // console.log("zonesToCorrect called");
    const text = (
      this.paragraphs ?
      this.paragraphs.map(el =>
        el.segments.map((el: any) => el.text).join("")
      ).join("\r\n\r\n") :
      "Please wait..."
    );
    return [{
      text,
      zoneId: "",
      zoneIsFocused: true,
    }];
  }

  updateParagraphs() {
    // console.log("Update text array");
    this.paragraphs = null;
    this.Asc.plugin.callCommand(() => {
      const oDocument = Api.GetDocument();
      const oDocumentInfo = oDocument.GetDocumentInfo();
      const title = oDocumentInfo.Title;

      let paragraphs: Paragraph[] = [], globalPos = 0;
      for (let i = 0; i < oDocument.GetElementsCount(); i++) {
        let oElement1 = oDocument.GetElement(i);

        let segments: Segment[] = [], relPos = 0;
        for (let j = 0; j < oElement1.GetElementsCount(); j++) {
          let oElement2 = oElement1.GetElement(j);

          if (oElement2) {
            let text = oElement2.GetText();
            segments.push({ relPos, text })
            relPos += text.length
          } else {
            segments.push({ relPos, text: "" })
          }

        }
        paragraphs.push({ globalPos, segments: segments });
        globalPos += relPos;
      }
      // console.log("paragraphs: ", paragraphs);

      return { title, paragraphs };
    },
    false,
    false,
    (res: {title: string, paragraphs: Paragraph[]}) => {
      for (let i = 1; i < res.paragraphs.length; i++) {
        // add new line length "\r\n"
        res.paragraphs[i].globalPos += 4 * i;
      }
      this.title = res.title;
      this.paragraphs = res.paragraphs;
    });
  }
}

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
    // console.log("antidotePort: ", antidotePort)
    if (antidotePort) {
      return Number(antidotePort);
    }

    throw new Error("Antidote port is not set.")
  }

  function launchCorrector() {
    AntidoteConnector.announcePresence();
    console.log("Status of AntidoteConnector: ", AntidoteConnector.isDetected());

    const agent = new ConnectixAgent(
      wordProcessorAgent,
      AntidoteConnector.isDetected() ?
      AntidoteConnector.getWebSocketPort :
      async () => getAntidotePort()
    );

    agent.connectWithAntidote()
      .then(() => agent.launchCorrector())
      .catch(error => {
        window.Asc.plugin.executeMethod("ShowWindow", ["iframe_asc.{E649827B-6CD5-477F-A7A7-C6952C813ADE}", connectionErrorModal]);

        console.log("Error Encountered: ", error)
      })
  }

  window.Asc.plugin.init = () => {
    window.Asc.plugin.attachEditorEvent("onDocumentContentReady", () => {
      wordProcessorAgent.updateParagraphs();
    });

    window.Asc.plugin.attachEditorEvent("onParagraphText", (data: any) => {
      if (!wordProcessorAgent.updateByAntidote) {
        wordProcessorAgent.updateParagraphs();
      }
    });

    launchCorrector();
  };

  window.Asc.plugin.button = function(id: string, windowId: string) {
    if (windowId === "iframe_asc.{E649827B-6CD5-477F-A7A7-C6952C813ADE}") {
      window.Asc.plugin.executeCommand("close", "");
    }
  };

})(window, undefined);
