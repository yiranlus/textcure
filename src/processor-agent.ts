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

type Segment = {
  relPos: number,
  text?: string
  readonly?: boolean,
  bold?: boolean,
  italic?: boolean,
  strike?: boolean,
  verticalAlign?: string
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
  updateByAntidote: boolean;
  paragraphs: Paragraph[] | null;

  constructor(Asc: any) {
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
      throw new Error("Data is empty");
    }

    let par = 0, seg = 0;
    if (eager) {
      while (
        par + 1 < this.paragraphs.length &&
        this.paragraphs[par + 1].globalPos <= pos)
          par++;
        while (
          seg + 1 < this.paragraphs[par].segments.length &&
          this.paragraphs[par].segments[seg+1].relPos <= pos - this.paragraphs[par].globalPos)
        seg++;
    } else {
      while (
        par + 1 < this.paragraphs.length &&
        this.paragraphs[par + 1].globalPos < pos)
      par++;
      while (
        seg + 1 < this.paragraphs[par].segments.length &&
        this.paragraphs[par].segments[seg+1].relPos < pos - this.paragraphs[par].globalPos)
      seg++;
    }

    return { par, seg };
  }

  correctIntoWordProcessor(params: ParamsReplace): boolean {
    if (!this.paragraphs) {
      throw new EmptyDataError();
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
          throw new EmptyDataError();
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
    if (!this.paragraphs) {
      throw new Error("No text found");
    }

    let indexStart = this.findIndex(params.positionStart, true);
    let indexEnd = this.findIndex(params.positionEnd);

    // console.log("params: ", params);
    // console.log("Index: ", indexStart, indexEnd);

    return (
      indexStart.par === indexEnd.par &&
      indexStart.seg === indexEnd.seg &&
      !this.paragraphs[indexStart.par].segments[indexStart.seg].readonly
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
    const styleInfos: StyleInfo[] = [];
    this.paragraphs?.forEach(paragraph =>
      paragraph.segments.forEach(segment => {
        const positionStart = paragraph.globalPos + segment.relPos;
        const positionEnd = positionStart + segment.text!.length;

        if (segment.bold)
          styleInfos.push({ positionStart, positionEnd, style: TextStyle.bold });
        if (segment.italic)
          styleInfos.push({ positionStart, positionEnd, style: TextStyle.italic });
        if (segment.strike)
          styleInfos.push({ positionStart, positionEnd, style: TextStyle.strike });
        if (segment.verticalAlign === "superscript") {
          styleInfos.push({ positionStart, positionEnd, style: TextStyle.superscript });
        } else if (segment.verticalAlign === "subscript") {
          styleInfos.push({ positionStart, positionEnd, style: TextStyle.subscript });
        }
      })
    );
    return [{
      text,
      zoneId: "",
      zoneIsFocused: true,
      styleInfo: styleInfos
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
        let oParagraph = oDocument.GetElement(i);

        let segments: Segment[] = [], relPos = 0;

        if (oParagraph.GetClassType() === "paragraph") {
          for (let j = 0; j < oParagraph.GetElementsCount(); j++) {
            let oSegment = oParagraph.GetElement(j);

            let segment: Segment = { relPos };
            if (oSegment.GetClassType() === "run") {
              const text = oSegment.GetText();
              segment.text = text;
              relPos += text.length;

              if (oSegment.GetBold())
                segment.bold = true;
              if (oSegment.GetItalic())
                segment.italic = true;
              if (oSegment.GetStrikeout() || oSegment.GetDoubleStrikeout())
                segment.strike = true;
              if (oSegment.GetVertAlign() === "superscript") {
                segment.verticalAlign = "superscript";
              } else if (oSegment.GetVertAlign() === "subscript") {
                segment.verticalAlign = "subscript";
              }
            } else if (oSegment.GetClassType() === "hyperlink") {
              const text = oSegment.GetDisplayedText();
              segment.text = text;
              segment.readonly = true;
              relPos += text.length

              segment.italic = true;
            }
            segments.push(segment);

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
