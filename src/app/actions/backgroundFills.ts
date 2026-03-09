import ShapeZOrder = PowerPoint.ShapeZOrder;
import {fixedColors, paintBucketColor} from "../taskpane";
import {ShapeTypeKey} from "../shared/types";
import {getSelectedShapeWith} from "../shared/utils/powerPointUtil";
import {FALLBACK_COLOR, ShapeType} from "../shared/consts";

async function initializeNewBackground(context: PowerPoint.RequestContext, shapeSelectValue: ShapeTypeKey, colorValue: string) {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);
  const background: PowerPoint.Shape = slide.shapes.addGeometricShape(ShapeType[shapeSelectValue]);

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

async function updateOrCreateIconGroup(context: PowerPoint.RequestContext, background: PowerPoint.Shape) {
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
    await updateOrCreateIconGroup(context, background);
  });
}

export async function chooseNewColor(color: string) {
  paintBucketColor.style.color = color;
  paintBucketColor.setAttribute("data-color", color);

  await PowerPoint.run(async (context) => {
    const shapeType = await getPreviousBackgroundShapeType(context);
    await addColoredBackground(shapeType);
  });
}

async function getPreviousBackgroundShapeType(context: PowerPoint.RequestContext) {
  const iconGroup = await getIconGroupWith(context);

  if (iconGroup.background) {
    iconGroup.background.load("name");
    await context.sync();
    return iconGroup.background.name as ShapeTypeKey;
  }
  return "Rectangle" as ShapeTypeKey;
}

export async function getIconGroupWith(context: PowerPoint.RequestContext) {
  const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);
  const group = await getGroupFromShape(context, selectedShape);

  if (group) {
    return await extractGroupItems(context, group);
  }
  return {icon: selectedShape, background: null};
}

async function extractGroupItems(context: PowerPoint.RequestContext, group: PowerPoint.Shape) {
  group.group.load("shapes");
  await context.sync();

  const groupItems = group.group.shapes.items;
  return {icon: groupItems[groupItems.length - 1], background: groupItems[0]};
}

async function getGroupFromShape(context: PowerPoint.RequestContext, shape: PowerPoint.Shape): Promise<PowerPoint.Shape | null> {
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