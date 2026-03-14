export function applyTranslation(Asc: Window["Asc"], id: string, text: string) {
  const element = document.getElementById(id);
  if (element) {
    element.innerHTML = Asc.plugin.tr(text);
  }
}
