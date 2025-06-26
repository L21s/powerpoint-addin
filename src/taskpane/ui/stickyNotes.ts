import { runPowerPoint } from "../utils/powerPointUtil";

const stickyNotes = document.querySelectorAll(".sticky-note");

export function initializeStickyNotes() {
  stickyNotes.forEach((button) => {
    const color = button.getAttribute("data-color");
    (button as HTMLElement).onclick = () => insertSticker(color);
  });
}

export async function insertSticker(color: string) {
  await runPowerPoint((powerPointContext) => {
    const today = new Date();
    const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;
    const textBox = shapes.addTextBox(localStorage.getItem("initials") + ", " + today.toDateString() + "\n", {
      height: 50,
      left: 50,
      top: 50,
      width: 150,
    });
    textBox.name = "Square";
    textBox.fill.setSolidColor(color);
    textBox.textFrame.bottomMargin = 7.087; // 0.25 cm in pt
    textBox.textFrame.topMargin = 7.087;
    textBox.textFrame.leftMargin = 4.2525; // 0.15 cm in pt
    textBox.textFrame.rightMargin = 4.2525;
    setStickerFontProperties(textBox);
  });
}

function setStickerFontProperties(textbox: PowerPoint.Shape) {
  textbox.textFrame.textRange.font.bold = true;
  textbox.textFrame.textRange.font.name = "Arial";
  textbox.textFrame.textRange.font.size = 12;
  textbox.textFrame.textRange.font.color = "#5A5A5A";
  textbox.lineFormat.visible = true;
  textbox.lineFormat.color = "#000000";
  textbox.lineFormat.weight = 1.25;
}
