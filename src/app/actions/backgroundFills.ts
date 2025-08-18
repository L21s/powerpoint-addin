import ShapeZOrder = PowerPoint.ShapeZOrder;
import {fixedColors, paintBucketColor} from "../taskpane";
import {ShapeTypeKey} from "../shared/types";
import {getSelectedShapeWith} from "../shared/utils/powerPointUtil";
import {FALLBACK_COLOR, ShapeType} from "../shared/consts";

async function initializeNewBackground(context: PowerPoint.RequestContext, shapeSelectValue: ShapeTypeKey, colorValue: string) {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const background: PowerPoint.Shape = slide.shapes.addGeometricShape(ShapeType[shapeSelectValue]);
  const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);

  background.name = shapeSelectValue;
  background.left = selectedShape.left;
  background.top = selectedShape.top;
  background.width = selectedShape.width;
  background.height = selectedShape.height;
  background.fill.setSolidColor(colorValue ? colorValue : FALLBACK_COLOR);
  background.lineFormat.visible = false;
  background.setZOrder(ShapeZOrder.sendToBack);

  return background;
}

// wohin damit? macht das so als function wirkich Sinn?
async function updateOrCreateIconGroupWith(context: PowerPoint.RequestContext, background: PowerPoint.Shape) {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const iconGroup = await getIconGroupWith(context);

  if (iconGroup.background) iconGroup.background.delete();
  slide.shapes.addGroup([background, iconGroup.icon]);
  await context.sync();
}

export async function addColoredBackground(shapeSelectValue: ShapeTypeKey) {
  await PowerPoint.run(async (context) => {
    const colorValue = paintBucketColor.getAttribute("data-color");
    const background = await initializeNewBackground(context, shapeSelectValue, colorValue);

    addColorToRecentColors(colorValue);
    await updateOrCreateIconGroupWith(context, background);
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
      console.log(iconGroup.background.name);
      oldBackgroundShape = iconGroup.background.name.split(" ")[0] as ShapeTypeKey;
    }
    await addColoredBackground(oldBackgroundShape);
  });
}

export async function getIconGroupWith(context: PowerPoint.RequestContext) {
  const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);
  const group = await getGroupFromSelectedShape(context, selectedShape);

  if (group) {
    group.group.load("shapes");
    await context.sync();

    const groupItems = group.group.shapes.items;
    return {icon: groupItems[groupItems.length - 1], background: groupItems[0]};
  }
  return {icon: selectedShape, background: null};
}

async function getGroupFromSelectedShape(context: PowerPoint.RequestContext, shape: PowerPoint.Shape): Promise<PowerPoint.Shape | null> {
  try {
    shape.load("parentGroup");
    await context.sync();
    return shape.parentGroup;
  } catch {
    return shape.type === "Group" ? shape : null;
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