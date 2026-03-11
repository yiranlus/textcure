let inputEl: HTMLInputElement | null;

window.Asc.plugin.init = function () {
    console.log("Init Settings");
    inputEl = document.getElementById("antidotePort") as HTMLInputElement;

    if (inputEl) {
        const antidotePort = localStorage.getItem("ANTIDOTE_PORT");
        if (antidotePort)
            inputEl.value = antidotePort;
        inputEl.focus();
    }
};

window.Asc.plugin.button = (id: string, windowId: string) => {
    if (!inputEl) {
        inputEl = document.getElementById("antidotePort") as HTMLInputElement;
    }
    console.log("Value of input El: ", inputEl)
    const value = inputEl ? Number(inputEl.value) : 0;


    // Send value back to main plugin context (optional)
    localStorage.setItem("ANTIDOTE_PORT", value.toString());
    console.log("Saved ANTIDOTE_PORT:", value);

    // Close modal
    // if (windowId) {
    console.log("windowId: ", windowId);
    window.Asc.plugin.executeCommand("close", "");
    // }
};
