export async function runPowerPoint(updateFunction: (context: PowerPoint.RequestContext) => void) {
  await PowerPoint.run(async (context) => {
    updateFunction(context);
    await context.sync();
  });
}

export async function getSelectedShapeWith(context: PowerPoint.RequestContext): Promise<PowerPoint.Shape> {
  const selectedShape = context.presentation.getSelectedShapes().getItemAt(0);
  selectedShape.load(["left", "top", "width", "height", "type"]);
  await context.sync();
  return selectedShape;
}

export async function getParentGroupWith(context: PowerPoint.RequestContext) {
  const selectedShape: PowerPoint.Shape = await getSelectedShapeWith(context);

  try {
    selectedShape.load("parentGroup");
    await context.sync();
    return selectedShape.parentGroup;
  } catch {
    console.debug("selected shape has no parent group");
    return selectedShape;
  }
}
