const popup = document.querySelector("sl-alert") as any;

export function showErrorPopup(errorMessage: string) {
  popup.querySelector("span").innerHTML = errorMessage;
  popup.toast();
}