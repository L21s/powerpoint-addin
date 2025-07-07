import {runPowerPoint} from "../shared/utils/powerPointUtil";
import ShapeCollection = PowerPoint.ShapeCollection;

export async function insertSticker(color: string) {
  await runPowerPoint((powerPointContext) => {
    const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
    createTextBox(shapes, color);
  });
}

function createTextBox(shapes: ShapeCollection, color: string) {
  const today = new Date();
  const textBox = shapes.addTextBox(localStorage.getItem("initials") + ", " + today.toDateString() + "\n", {
    height: 50,
    left: 50,
    top: 50,
    width: 150,
  });
  textBox.name = "Square";
  textBox.fill.setSolidColor(color);
  setTextProperties(textBox);
  return textBox;
}

function setTextProperties(textBox: PowerPoint.Shape) {
  textBox.textFrame.bottomMargin = 7.087; // 0.25 cm in pt
  textBox.textFrame.topMargin = 7.087;
  textBox.textFrame.leftMargin = 4.2525; // 0.15 cm in pt
  textBox.textFrame.rightMargin = 4.2525;
  textBox.textFrame.textRange.font.bold = true;
  textBox.textFrame.textRange.font.name = "Arial";
  textBox.textFrame.textRange.font.size = 12;
  textBox.textFrame.textRange.font.color = "#5A5A5A";
  textBox.lineFormat.visible = true;
  textBox.lineFormat.color = "#000000";
  textBox.lineFormat.weight = 1.25;
}
