import { getSelectedShape } from "./powerPointUtil";
import { ShapeType, ShapeTypeKey } from "./types";

function addColorToRecent(colorValue: string) {
  const recentColorElements = document.querySelectorAll(".fixed-color");
  let recentColors = [];

  recentColorElements.forEach((button: HTMLElement) => {
    recentColors.push(button.getAttribute("data-color"));
  });

  if (!recentColors.includes(colorValue)) {
    recentColors.unshift(colorValue);
    recentColors.pop();

    for (let index = 0; index < recentColors.length; index++) {
      (recentColorElements[index] as HTMLElement).style.backgroundColor = recentColors[index];
      (recentColorElements[index] as HTMLElement).setAttribute("data-color", recentColors[index]);
    }
  }
}

async function addColoredBackground(shapeSelectValue: ShapeTypeKey) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const selectedShape: PowerPoint.Shape = await getSelectedShape();
    const colorValue = document.getElementById("paint-bucket-color").getAttribute("data-color");
    const background: PowerPoint.Shape = slide.shapes.addGeometricShape(
      ShapeType[shapeSelectValue ? shapeSelectValue : "Rectangle"]
    );

    background.left = selectedShape.left;
    background.top = selectedShape.top;
    background.width = selectedShape.width;
    background.height = selectedShape.height;
    background.fill.setSolidColor(colorValue ? colorValue : "lightgreen");

    addColorToRecent(colorValue);
  });

  /**
   * Note: something like the code stated below should be used right here ... but sadly, the Powerpoint-Context still does not offer this.
   * This code is supposed to stay as reminder that maybe in future days Microsoft may offer this.
   * Better: Select Icon --> Bring to front.
   */
  /*await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const shape = sheet.shapes.getItem("MyShape"); // use shape name or .getItemAt(0)
      shape.setZOrder("SendBackward");
      await context.sync();
  });
   */
}

function chooseNewColor(color: string) {
  const paintBucketIcon =  document.getElementById("paint-bucket-color");
  paintBucketIcon.style.color = color;
  paintBucketIcon.setAttribute("data-color", color);
}

export function registerIconBackgroundTools() {
  document.querySelectorAll(".shape-option").forEach((button: HTMLElement) => {
    button.onclick = () => {
      addColoredBackground(button.getAttribute("data-value") as ShapeTypeKey);
    };
  });

  document.getElementById("background-color-picker").addEventListener("change", async (e) => {
    chooseNewColor((e.target as HTMLInputElement).value);
  });

  document.querySelectorAll(".fixed-color").forEach((button: HTMLElement) => {
    button.onclick = () => {
      chooseNewColor(button.getAttribute("data-color"));
    };
  });
}
