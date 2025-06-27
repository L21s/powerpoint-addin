import { getSelectedShapeWith, getParentGroupWith } from "./powerPointUtil";
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
    const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);
    const colorValue = document.getElementById("paint-bucket-color").getAttribute("data-color");
    const background: PowerPoint.Shape = slide.shapes.addGeometricShape(ShapeType[shapeSelectValue]);

    background.name = shapeSelectValue;
    background.left = selectedShape.left;
    background.top = selectedShape.top;
    background.width = selectedShape.width;
    background.height = selectedShape.height;
    background.fill.setSolidColor(colorValue ? colorValue : "lightgreen");
    background.lineFormat.visible = false;
    background.setZOrder(ShapeZOrder.sendToBack);

    addColorToRecentColors(colorValue);
    const iconGroup = await getGroupElements(context);
    if (iconGroup.background) iconGroup.background.delete();
    slide.shapes.addGroup([background, iconGroup.icon]);
    await context.sync();
  });
}

async function getGroupElements(context: PowerPoint.RequestContext) {
  const selectedShape: PowerPoint.Shape = await getParentGroupWith(context);

  if (selectedShape.type === "Group") {
    selectedShape.group.load("shapes");
    await context.sync();

    const groupItems = selectedShape.group.shapes.items;
    return { icon: groupItems[groupItems.length - 1], background: groupItems[0] };
  } else {
    return { icon: selectedShape, background: null };
  }
}

async function chooseNewColor(color: string) {
  const paintBucketIcon = document.getElementById("paint-bucket-color");
  paintBucketIcon.style.color = color;
  paintBucketIcon.setAttribute("data-color", color);

  await PowerPoint.run(async (context) => {
    let oldBackgroundShape: ShapeTypeKey = "Rectangle";
    const iconGroup = await getGroupElements(context);

    if (iconGroup.background) {
      iconGroup.background.load("name");
      await context.sync();
      oldBackgroundShape = iconGroup.background.name.split(" ")[0] as ShapeTypeKey;
    }
    await addColoredBackground(oldBackgroundShape);
  });
}

export function registerIconBackgroundTools() {
  document.querySelectorAll(".shape-option").forEach((button: HTMLElement) => {
    button.onclick = async () => {
      await addColoredBackground(button.getAttribute("data-value") as ShapeTypeKey);
    };
  });

  document.getElementById("background-color-picker").addEventListener("change", async (e) => {
    await chooseNewColor((e.target as HTMLInputElement).value);
  });

  document.querySelectorAll(".fixed-color").forEach((button: HTMLElement) => {
    button.onclick = async () => {
      await chooseNewColor(button.getAttribute("data-color"));
    };
  });

  document.getElementById("paint-bucket").onclick = async (e) => {
    await chooseNewColor(document.getElementById("paint-bucket-color").style.color);
  };

  document.getElementById("delete-background").onclick = async () => {
    await PowerPoint.run(async (context) => {
      const iconGroup = await getGroupElements(context);
      if (iconGroup.background) iconGroup.background.delete();
    });
  };
}
