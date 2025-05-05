import { getSelectedShape } from "./powerPointUtil";
import { ShapeType, ShapeTypeKey } from "./types";

export async function addColoredBackground(shapeSelectValue: ShapeTypeKey) {
  await PowerPoint.run(async (context) => {
    // #0. Get slide and selected shape
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const selectedShape: PowerPoint.Shape = await getSelectedShape();

    // #1. Get current background color from paint bucket icon
    const colorValue = RGBAToHex(document.getElementById("current-color").style.color);

    // #2. Build background with given background shape (shapeSelectValue)
    const background: PowerPoint.Shape = slide.shapes.addGeometricShape(
      ShapeType[shapeSelectValue ? shapeSelectValue : "Rectangle"]
    );
    background.left = selectedShape.left;
    background.top = selectedShape.top;
    background.width = selectedShape.width;
    background.height = selectedShape.height;
    background.fill.setSolidColor(colorValue ? colorValue : "lightgreen");

    // #3. After inserting background, add color as recently used color in the dropdown
    const recentColorElements = document.querySelectorAll(".fixed-color");
    let recentColors = [];

    recentColorElements.forEach((button: HTMLElement) => {
      recentColors.push(RGBAToHex(button.style.backgroundColor));
    });

    // only add the color if it's not already added
    if (!recentColors.includes(colorValue)) {
      recentColors.unshift(colorValue);
      recentColors.pop();

      for (let index = 0; index < recentColors.length; index++) {
        (recentColorElements[index] as HTMLElement).style.backgroundColor = recentColors[index];
      }
    }
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

export function chooseNewColor(color: string) {
  // apply new color to paint bucket icon
  document.getElementById("current-color").style.color = color;
}

export function RGBAToHex(rgba: string) {
  return `#${rgba
    .match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+\.{0,1}\d*))?\)$/)
    .slice(1)
    .map((n, i) =>
      (i === 3 ? Math.round(parseFloat(n) * 255) : parseFloat(n)).toString(16).padStart(2, "0").replace("NaN", "")
    )
    .join("")}`;
}
