import { getSelectedShapeWith } from "./powerPointUtil";
import { ShapeType, ShapeTypeKey } from "./types";
import ShapeZOrder = PowerPoint.ShapeZOrder;

function addColorToRecentColors(colorValue: string) {
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
    let selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);
    const colorValue = document.getElementById("paint-bucket-color").getAttribute("data-color");
    const background: PowerPoint.Shape = slide.shapes.addGeometricShape(
      ShapeType[shapeSelectValue ? shapeSelectValue : "Rectangle"]
    );

    background.left = selectedShape.left;
    background.top = selectedShape.top;
    background.width = selectedShape.width;
    background.height = selectedShape.height;
    background.fill.setSolidColor(colorValue ? colorValue : "lightgreen");
    background.lineFormat.visible = false;
    background.setZOrder(ShapeZOrder.sendToBack);

    addColorToRecentColors(colorValue);

    /*
    async function removeBackground(selectedGroup: PowerPoint.Shape) {
      selectedGroup.group.load("shapes");
      await context.sync();

      const groupItems = selectedGroup.group.shapes.items;
      const imageItem = groupItems[groupItems.length - 1];
      groupItems[0].delete();
      return imageItem;
    }
    */

    let testShape: PowerPoint.Shape;
    try {
      selectedShape.load("parentGroup");
      await context.sync();
      testShape = selectedShape.parentGroup;
    } catch {
      testShape = selectedShape;
    }

    testShape.load("type");
    await context.sync();
    if (testShape.type === "Group") selectedShape = await removeBackground(testShape);

    slide.shapes.addGroup([background, selectedShape]);
    await context.sync();
  });
}

function chooseNewColor(color: string) {
  const paintBucketIcon = document.getElementById("paint-bucket-color");
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
