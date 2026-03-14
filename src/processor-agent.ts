import {
  WordProcessorAgent,
  ParamsReplace,
  ParamsAllowEdit,
  WordProcessorConfiguration,
  ParamsGetZonesToCorrect,
  TextZoneConnectix,
  DocumentType,
  StyleInfo,
  TextStyle
} from "@druide-informatique/antidote-api-js";

type Paragraph = {
  globalPos: number,
  text?: string
}

type TextRange = {
  start: number,
  end: number
}

export class EmptyDataError extends Error {
  constructor() {
    super("Data is empty");
    this.name = 'EmptyDataError';
    Object.setPrototypeOf(this, EmptyDataError.prototype);
  }
}

export class WordProcessorAgentOnlyOfficeDocument extends WordProcessorAgent {
  title: string;
  Asc: any;
  updatingByAntidote: boolean;
  paragraphs: Paragraph[] | null;

  constructor(Asc: any) {
    super();
    this.Asc = Asc;
    this.updatingByAntidote = false;

    this.title = "";
    this.paragraphs = null;
  }

  sessionEnded() {
    this.Asc.plugin.executeCommand("close", "");
    super.sessionEnded();
  }

  findIndex(pos: number, eager: boolean=false): number {
    if (!this.paragraphs) {
      throw new Error("Data is empty");
    }

    let elementIndex = 0;
    if (eager) {
      while (
        elementIndex + 1 < this.paragraphs.length &&
        this.paragraphs[elementIndex + 1].globalPos <= pos)
          elementIndex++;
    } else {
      while (
        elementIndex + 1 < this.paragraphs.length &&
        this.paragraphs[elementIndex + 1].globalPos < pos)
      elementIndex++;
    }

    return elementIndex;
  }

  correctIntoWordProcessor(params: ParamsReplace): boolean {
    this.updatingByAntidote = true;

    let elementIndex = this.findIndex(params.positionStartReplace, true);

    let text = this.paragraphs![elementIndex].text!;
    let newText = (
      text.substring(0, params.positionStartReplace - this.paragraphs![elementIndex].globalPos) +
      params.newString +
      text.substring(params.positionReplaceEnd - this.paragraphs![elementIndex].globalPos)
      ).replace(/(\r\n)*$/, "");

    this.Asc.scope.paramsReplace = { elementIndex, newText };

    try {
      this.Asc.plugin.callCommand(() => {
        const { elementIndex, newText } = Asc.scope.paramsReplace;

        var oDocument = Api.GetDocument();
        var oElement = oDocument.GetElement(elementIndex);

        var oldText = oElement.GetText();

        oElement.Select();
        Api.ReplaceTextSmart([newText]);

        return {
          text: oElement.GetText(),
          diff: oElement.GetText().length - oldText.length
        }
      },
      false,
      true,
      (res: {text: string, diff: number}) => {
        this.paragraphs![elementIndex].text = res.text;
        for (let i = elementIndex + 1; i < this.paragraphs!.length; i++) {
          this.paragraphs![i].globalPos += res.diff;
        }
        this.updatingByAntidote = false;
      });
    } catch (error) {
      return false;
    }

    return true;
  }

  configuration(): WordProcessorConfiguration {
    return {
      documentTitle: this.title,
      activeMarkup: DocumentType.text
    };
  }

  allowEdit(params: ParamsAllowEdit): boolean {
    return true;
  }

  textZonesAvailable(): boolean {
    return !!this.paragraphs;
  }

  zonesToCorrect(_params: ParamsGetZonesToCorrect): TextZoneConnectix[] {
    const text = (
      this.paragraphs ?
        this.paragraphs.map(el => el.text).join("\r\n\r\n") :
      "Please wait..."
    );
    return [{
      text,
      zoneId: "",
      zoneIsFocused: true,
    }];
  }

  updateParagraphs() {
    this.paragraphs = null;
    this.Asc.plugin.callCommand(() => {
      const oDocument = Api.GetDocument();
      const oDocumentInfo = oDocument.GetDocumentInfo();
      const title = oDocumentInfo.Title;

      let paragraphs: Paragraph[] = [], globalPos = 0;
      for (let i = 0; i < oDocument.GetElementsCount(); i++) {
        const element = oDocument.GetElement(i);
        if (element.GetClassType() === "paragraph") {
          const text = element.GetText().replace(/(\r\n)*$/, "");
          paragraphs.push({ globalPos, text });
          globalPos += text.length;
        } else {
          paragraphs.push({ globalPos });
        }
      }

      return { title, paragraphs };
    },
    false,
    false,
    (res: {title: string, paragraphs: Paragraph[]}) => {
      for (let i = 1; i < res.paragraphs.length; i++) {
        res.paragraphs[i].globalPos += 4 * i;
      }
      this.title = res.title;
      this.paragraphs = res.paragraphs;
    });
  }
}
