import { applyTranslation } from "./utils";

window.Asc.plugin.init = () => {
}

window.Asc.plugin.onTranslate = () => {
  applyTranslation(window.Asc, "antidote-connction-error-heading", "Antidote Connection Error");
  applyTranslation(window.Asc, "antidote-connection-error-message", "Please make sure that the port number is correct or that Antidote Connector is installed.");
};
