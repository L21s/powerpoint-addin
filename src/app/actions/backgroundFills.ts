import ShapeZOrder = PowerPoint.ShapeZOrder;
import {fixedColors, paintBucketColor} from "../taskpane";
import {ShapeTypeKey} from "../shared/types";
import {getSelectedShapeWith} from "../shared/utils/powerPointUtil";
import {ShapeType} from "../shared/consts";

export async function addColoredBackground(shapeSelectValue: ShapeTypeKey) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);
    const colorValue = paintBucketColor.getAttribute("data-color");
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

    const iconGroup = await getIconGroupWith(context);
    if (iconGroup.background) iconGroup.background.delete();
    slide.shapes.addGroup([background, iconGroup.icon]);
    await context.sync();
  });
}

export async function chooseNewColor(color: string) {
  paintBucketColor.style.color = color;
  paintBucketColor.setAttribute("data-color", color);

  await PowerPoint.run(async (context) => {
    let oldBackgroundShape: ShapeTypeKey = "Rectangle";
    const iconGroup = await getIconGroupWith(context);

    if (iconGroup.background) {
      iconGroup.background.load("name");
      await context.sync();
      oldBackgroundShape = iconGroup.background.name.split(" ")[0] as ShapeTypeKey;
    }
    await addColoredBackground(oldBackgroundShape);
  });
}

export async function getIconGroupWith(context: PowerPoint.RequestContext) {
  const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);
  let selectedGroup: PowerPoint.Shape;

  try {
    selectedShape.load("parentGroup");
    await context.sync();
    selectedGroup = selectedShape.parentGroup;
  } catch {
    selectedGroup = selectedShape;
  }

  if (selectedGroup.type === "Group") {
    selectedGroup.group.load("shapes");
    await context.sync();

    const groupItems = selectedGroup.group.shapes.items;
    return {icon: groupItems[groupItems.length - 1], background: groupItems[0]};
  } else {
    return {icon: selectedShape, background: null};
  }
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