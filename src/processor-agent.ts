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
import { Mutex } from "async-mutex";

type Paragraph = {
  globalPos: number,
  text?: string
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

  replacingQueue: ParamsReplace[];
  mutexQueue: Mutex;
  mutexDocument: Mutex;

  constructor(Asc: any) {
    super();
    this.Asc = Asc;
    this.updatingByAntidote = false;

    this.title = "";
    this.paragraphs = null;

    this.replacingQueue = [];
    this.mutexQueue = new Mutex();
    this.mutexDocument = new Mutex();
  }

  sessionEnded() {
    this.Asc.plugin.executeCommand("close", "");
    super.sessionEnded();
  }

  findIndex(pos: number, eager: boolean = false): number {
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
    if (!this.paragraphs) return false;

    this.mutexQueue.runExclusive(() => {
      // console.log("Locking Queue");
      this.replacingQueue.push(params);
    })
      .then(() => {
        // console.log("Calling Apply Corrections")
        this.applyCorrections();
      })
      .catch(error => {
        console.log(error);
      })

    return true;
  }

  async applyCorrections() {
    await this.mutexDocument.runExclusive(async () => {
      this.updatingByAntidote = true;

      let params = await this.mutexQueue.runExclusive(() => {
        // console.log("Retriving Item from the Queue")
        return this.replacingQueue.shift();
      });

      while (params) {
        // console.log("Applying a Correction")
        await this._correctIntoWordProcessor(params!);

        params = await this.mutexQueue.runExclusive(() => {
          // console.log("Retriving Item from the Queue")
          return this.replacingQueue.shift();
        })
      }

      this.updatingByAntidote = false;
    });
  }

  async _correctIntoWordProcessor(params: ParamsReplace) {
    // Waiting to previous action to finish
    // console.log("ParasReplace: ", params);

    let elementIndex = this.findIndex(params.positionStartReplace, true);

    let text = this.paragraphs![elementIndex].text!;
    let newText = (
      text.substring(0, params.positionStartReplace - this.paragraphs![elementIndex].globalPos) +
      params.newString +
      text.substring(params.positionReplaceEnd - this.paragraphs![elementIndex].globalPos)
    ).replace(/(\r\n)*$/, "");

    this.Asc.scope.paramsReplace = { elementIndex, text: newText };

    return new Promise<void>(resolve => {
      this.Asc.plugin.callCommand(() => {
        const { elementIndex, text } = Asc.scope.paramsReplace;

        var oDocument = Api.GetDocument();
        var oElement = oDocument.GetElement(elementIndex);

        var oldText = oElement.GetText({ "Numbering": false }).replace(/(\r\n)*$/, "");

        oElement.Select();
        Api.ReplaceTextSmart([text]);

        const newText = oElement.GetText({ "Numbering": false }).replace(/(\r\n)*$/, "");

        return {
          text: newText,
          diff: newText.length - oldText.length
        }
      },
      false,
      true,
      (res: { text: string, diff: number }) => {
        this.paragraphs![elementIndex].text = res.text;
        for (let i = elementIndex + 1; i < this.paragraphs!.length; i++) {
          this.paragraphs![i].globalPos += res.diff;
        }
        resolve();
      });
    });
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
    if (this.replacingQueue.length > 0
      && !this.mutexQueue.isLocked()
      && !this.mutexDocument.isLocked())
      return false;
    return !!this.paragraphs;
  }

  zonesToCorrect(_params: ParamsGetZonesToCorrect): TextZoneConnectix[] {
    const text = this.paragraphs!.map(el => el.text).join("\r\n\r\n");
    return [{
      text,
      zoneId: "",
      zoneIsFocused: true,
    }];
  }

  updateParagraphs() {
    console.log("UpdateParagraphs called.");
    return this.mutexDocument.runExclusive(() => this._updateParagraphs());
  }

  _updateParagraphs(): Promise<void> {
    console.log("_updateParagraphs called");
    this.paragraphs = null;

    return new Promise<void>(resolve => {
      this.Asc.plugin.callCommand(() => {
        const oDocument = Api.GetDocument();
        const oDocumentInfo = oDocument.GetDocumentInfo();
        const title = oDocumentInfo.Title;

        let paragraphs: Paragraph[] = [], globalPos = 0;
        for (let i = 0; i < oDocument.GetElementsCount(); i++) {
          const element = oDocument.GetElement(i);
          if (element.GetClassType() === "paragraph") {
            const text = element.GetText({ "Numbering": false }).replace(/(\r\n)*$/, "");
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
        (res: { title: string, paragraphs: Paragraph[] }) => {
          for (let i = 1; i < res.paragraphs.length; i++) {
            res.paragraphs[i].globalPos += 4 * i;
          }
          this.title = res.title;
          this.paragraphs = res.paragraphs;

          resolve();
        });
    });
  }
}
