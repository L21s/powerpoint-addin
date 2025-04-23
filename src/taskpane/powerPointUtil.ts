export async function runPowerPoint(updateFunction: (context: PowerPoint.RequestContext) => void) {
  await PowerPoint.run(async (context) => {
    updateFunction(context);
    await context.sync();
  });
}

export async function getSelectedShape(): Promise<PowerPoint.Shape> {
  return new Promise((resolve, reject) => {
    runPowerPoint(async (powerPointContext: PowerPoint.RequestContext) => {
      try {
        const selectedShape = powerPointContext.presentation.getSelectedShapes().getItemAt(0);
        selectedShape.load(["left", "top", "width", "height"]);
        await powerPointContext.sync();
        resolve(selectedShape);
      } catch (error) {
        reject(error);
      }
    });
  });
}
