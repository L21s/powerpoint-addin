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
    const iconElement = await removeBackground(context);
    slide.shapes.addGroup([background, iconElement]);
    await context.sync();
  });
}

export async function removeBackground(context: PowerPoint.RequestContext) {
  let selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);

  try {
    selectedShape.load("parentGroup");
    await context.sync();
    selectedShape = selectedShape.parentGroup;
  } catch {}

  selectedShape.load("type");
  await context.sync();

  if (selectedShape.type === "Group") {
    selectedShape.group.load("shapes");
    await context.sync();

    const groupItems = selectedShape.group.shapes.items;
    selectedShape = groupItems[groupItems.length - 1];
    groupItems[0].delete();
  }
  return selectedShape;
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

  document.getElementById("delete-background").onclick = async () => {
    await PowerPoint.run(async (context) => {
      await removeBackground(context);
      await context.sync();
    });
  };
}
