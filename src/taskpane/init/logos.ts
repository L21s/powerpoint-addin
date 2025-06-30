import {logoDropdownOptions} from "../taskpane";
import {handleLogoImageInsert} from "../listeners/logos";

export function initializeLogoDropdown() {
    logoDropdownOptions.forEach((button: HTMLElement) => {
        button.onclick = async () => handleLogoImageInsert(button)
    });
}