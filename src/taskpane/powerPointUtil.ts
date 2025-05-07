
export async function runPowerPoint(updateFunction: (context: PowerPoint.RequestContext) => void) {
  await PowerPoint.run(async (context) => {
    updateFunction(context);
    await context.sync();
  });
}

export async function getSelectedShapeWith(context: PowerPoint.RequestContext): Promise<PowerPoint.Shape> {

  const selectedShape = context.presentation.getSelectedShapes().getItemAt(0);
  selectedShape.load(["left", "top", "width", "height"]);
  await context.sync();
  return selectedShape;
}