import {backgroundColorPicker, fixedColors, shapeOptions} from "../taskpane";
import {addColoredBackground, chooseNewColor} from "../listeners/imageBackgroundEditor";
import {ShapeTypeKey} from "../types";

export function initializeImageBackgroundEditor() {
    shapeOptions.forEach((button: HTMLElement) => {
        button.onclick = () => addColoredBackground(button.getAttribute("data-value") as ShapeTypeKey);
    });

    backgroundColorPicker.addEventListener("change", async (e) => {
        chooseNewColor((e.target as HTMLInputElement).value);
    });

    fixedColors.forEach((button: HTMLElement) => {
        button.onclick = () => chooseNewColor(button.getAttribute("data-color"));
    });
}