import { getImageAsBase64 } from "../utils/imageUtils";
import {showErrorPopup} from "./errorPopup";

export async function handleLogoImageInsert(button: HTMLElement) {
  const selectedImageSrc = button.getElementsByTagName("img")[0].src;
  const currentDropdownImage = document.getElementById(
      selectedImageSrc.includes("Text") ? "currentWithText" : "currentWithoutText"
  ) as HTMLImageElement;

  currentDropdownImage.src = selectedImageSrc;
  if (selectedImageSrc.includes("White")) {
    currentDropdownImage.classList.add("white-shadow");
  } else {
    currentDropdownImage.classList.remove("white-shadow");
  }

  Office.context.document.setSelectedDataAsync(
      ((await getImageAsBase64(selectedImageSrc)) as string).split(",")[1],
      { coercionType: Office.CoercionType.Image },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          const errorMessage = "Action failed. Error: " + asyncResult.error.message;
          showErrorPopup(errorMessage);
        }
      }
  );
}