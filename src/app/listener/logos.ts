import {logoDropdownOptions} from "../taskpane";
import {handleLogoImageInsert} from "../actions/logos";

export function initializeLogoDropdownListener() {
    logoDropdownOptions.forEach((button: HTMLElement) => {
        button.onclick = async () => handleLogoImageInsert(button)
    });
}