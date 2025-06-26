import { getImageAsBase64 } from "../utils/imageUtils";
import {showErrorPopup} from "./errorPopup";

const logoDropdownOptions = document.querySelectorAll(".logo-dropdown, .logo-dropdown-option");

export function initializeLogoDropdown() {
  logoDropdownOptions.forEach((button: HTMLElement) => {
    button.onclick = async () => handleLogoImageInsert(button)
  });
}

async function handleLogoImageInsert(button: HTMLElement) {
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