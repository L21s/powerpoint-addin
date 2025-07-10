import {backgroundColorPicker, deleteBackground, fixedColors, paintBucket, paintBucketColor, shapeOptions} from "../taskpane";
import {addColoredBackground, chooseNewColor, getIconGroupWith} from "../actions/backgroundFills";
import {ShapeTypeKey} from "../shared/types";

export function initializeImageBackgroundEditorListener() {
  shapeOptions.forEach((button: HTMLElement) => {
    button.onclick = () => addColoredBackground(button.getAttribute("data-value") as ShapeTypeKey);
  });

  backgroundColorPicker.addEventListener("change", async (e) => {
    await chooseNewColor((e.target as HTMLInputElement).value);
    (e.target as HTMLInputElement).closest("sl-dropdown")["open"] = false;
  });

  fixedColors.forEach((button: HTMLElement) => {
    button.onclick = () => chooseNewColor(button.getAttribute("data-color"));
  });

  paintBucket.onclick = async () => {
    await chooseNewColor(paintBucketColor.getAttribute("data-color"));
  };

  deleteBackground.onclick = async () => {
    await PowerPoint.run(async (context) => {
      const iconGroup = await getIconGroupWith(context);
      if (iconGroup.background) iconGroup.background.delete();
    });
  };
}