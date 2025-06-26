import { getSelectedShapeWith } from "../utils/powerPointUtil";
import { ShapeType, ShapeTypeKey } from "../types";
import ShapeZOrder = PowerPoint.ShapeZOrder;

const shapeOptions = document.querySelectorAll(".shape-option");
const backgroundColorPicker = document.getElementById("background-color-picker");
const fixedColors = document.querySelectorAll(".fixed-color");
const paintBucketColor = document.getElementById("paint-bucket-color");

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

function addColorToRecentColors(colorValue: string) {
  let recentColors = [];

  fixedColors.forEach((button: HTMLElement) => {
    recentColors.push(button.getAttribute("data-color"));
  });

  if (!recentColors.includes(colorValue)) {
    recentColors.unshift(colorValue);
    recentColors.pop();

    for (let index = 0; index < recentColors.length; index++) {
      (fixedColors[index] as HTMLElement).style.backgroundColor = recentColors[index];
      (fixedColors[index] as HTMLElement).setAttribute("data-color", recentColors[index]);
    }
  }
}

async function addColoredBackground(shapeSelectValue: ShapeTypeKey) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);
    const colorValue = paintBucketColor.getAttribute("data-color");
    const background: PowerPoint.Shape = slide.shapes.addGeometricShape(
      ShapeType[shapeSelectValue ? shapeSelectValue : "Rectangle"]
    );

    background.left = selectedShape.left;
    background.top = selectedShape.top;
    background.width = selectedShape.width;
    background.height = selectedShape.height;
    background.fill.setSolidColor(colorValue ? colorValue : "lightgreen");
    background.lineFormat.visible = false;

    addColorToRecentColors(colorValue);

    background.setZOrder(ShapeZOrder.sendToBack);
    await context.sync();
    slide.shapes.addGroup([background, selectedShape]);
    await context.sync();
  });
}

function chooseNewColor(color: string) {
  paintBucketColor.style.color = color;
  paintBucketColor.setAttribute("data-color", color);
}