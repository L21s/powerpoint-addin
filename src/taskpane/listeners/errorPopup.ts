import {popup} from "../taskpane";

export function showErrorPopup(errorMessage: string) {
  popup.querySelector("span").innerHTML = errorMessage;
  popup.toast();
}