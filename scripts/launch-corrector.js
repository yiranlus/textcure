import { WordProcessorAgent, AntidoteConnector, ConnectixAgent } from "./antidote.min.js";

// const antidotePort = 49157;

class WordProcessorAgentOnlyOfficeDocument extends WordProcessorAgent {
    constructor(Asc, title, textArray) {
        super();
        this.Asc = Asc;
        this.title = title;
        this.textArray = textArray;
        for (let i = 1; i < textArray.length; i++) {
            this.textArray[i].globalPos += 4 * i;
        }
    }

    sessionEnded() {
        this.Asc.plugin.executeCommand("close", "");
    }

    findIndex(pos) {
        let i = 0, j = 0;
        while (
            i + 1 < this.textArray.length &&
            this.textArray[i + 1].globalPos < pos)
            i++;
        while (
            j + 1 < this.textArray[i].textArray.length &&
            this.textArray[i].textArray[j+1].relPos < pos - this.textArray[i].globalPos)
            j++;

        return {
            index1: i,
            index2: j,
        }
    }

    correctIntoWordProcessor(params) {
        let start = params.positionStartReplace;
        let end = params.positionReplaceEnd

        let index = this.findIndex(start);
        let i = index.index1, j = index.index2;

        let textElement = this.textArray[i].textArray[j];

        this.Asc.scope.paramsReplace = {
            index1: i,
            index2: j,
            start: start - this.textArray[i].globalPos - this.textArray[i].textArray[j].relPos,
            end: end - this.textArray[i].globalPos - this.textArray[i].textArray[j].relPos,
            newString: params.newString
        }

        try {
            this.Asc.plugin.callCommand(() => {
                var index1 = Asc.scope.paramsReplace.index1;
                var index2 = Asc.scope.paramsReplace.index2;
                var start = Asc.scope.paramsReplace.start;
                var end = Asc.scope.paramsReplace.end;
                var newString = Asc.scope.paramsReplace.newString;

                var oDocument = Api.GetDocument();
                var oElement = oDocument.GetElement(index1).GetElement(index2);
                var oRange = oElement.GetRange(start, end);

                oRange.Delete();
                oRange.AddText(newString);
                return {
                    newText: oElement.GetText(),
                    diff: newString.length - (end - start)
                }
            },
            false,
            true,
            (res) => {
                this.textArray[index.index1].textArray[j].text = res.newText;
                for (let j = index.index2 + 1; j < this.textArray[index.index1].textArray.length; j++) {
                    this.textArray[index.index1].textArray[j].relPos += res.diff;
                }
                for (let i = index.index1 + 1; i < this.textArray.length; i++) {
                    this.textArray[i].globalPos += res.diff;
                }
            });
        } catch (error) {
            console.log("error: ", error);
            return false;
        }

        return true;
    }

    configuration() {
        return {
            documentTitle: this.title,
            carriageReturn: "\r\n",
            activeMarkup: "text"
        };
    }

    allowEdit(params) {
        let indexStart = this.findIndex(params.positionStart);
        let indexEnd = this.findIndex(params.positionEnd - 1);

        // console.log("params: ", params);
        // console.log("Index: ", indexStart, indexEnd);

        return (
            indexStart.index1 === indexEnd.index1 &&
            indexStart.index2 === indexEnd.index2
        );
    }

    zonesToCorrect(_params) {
        const text = this.textArray.map((el, index) =>
            el.textArray.map(el => el.text).join("")
        ).join("\r\n\r\n");
        return [{
            text,
            zoneId: "0",
            zoneIsFocused: true,
        }];
    }
}

(function(window, undefined){
    function getFullUrl(name) {
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
        isVisual : true,
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
        var antidotePort = localStorage.getItem("ANTIDOTE_PORT");
        console.log("antidotePort: ", antidotePort)
        if (antidotePort) {
            return Number(antidotePort);
        }

        throw new Error("Antidote port is not set.")
    }

    function launchCorrector(Asc, title, textArray) {
        AntidoteConnector.announcePresence();
        console.log("Status of AntidoteConnector: ", AntidoteConnector.isDetected());

        const agent = new ConnectixAgent(
            new WordProcessorAgentOnlyOfficeDocument(Asc, title, textArray),
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

    window.Asc.plugin.init = function() {
        window.Asc.plugin.callCommand(() => {
            const oDocument = Api.GetDocument();
            const oDocumentInfo = oDocument.GetDocumentInfo();

            let start = 0, end = 0;

            let text = "";

            let textArray = [], globalPos = 0;
            for (let i = 0; i < oDocument.GetElementsCount(); i++) {
                let oElement1 = oDocument.GetElement(i);

                let subTextArray = [], relPos = 0;
                for (let j = 0; j < oElement1.GetElementsCount(); j++) {
                    let oElement2 = oElement1.GetElement(j);

                    if (oElement2) {
                        let text = oElement2.GetText();
                        subTextArray.push({ relPos, text })
                        relPos += text.length
                    } else {
                        subTextArray.push({ relPos, text: null })
                    }

                }
                textArray.push({ globalPos, textArray: subTextArray });
                globalPos += relPos;
            }

            return {
                title: oDocumentInfo.Title,
                textArray
            }
        },
        false,
        false,
        (res) => {
            launchCorrector(window.Asc, res.title, res.textArray);
        })
    };

    window.Asc.plugin.button = function(id, windowId) {
        if (windowId === "iframe_asc.{E649827B-6CD5-477F-A7A7-C6952C813ADE}") {
            window.Asc.plugin.executeCommand("close", "");
        }
    };
})(window, undefined);
