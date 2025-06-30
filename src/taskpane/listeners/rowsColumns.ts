import {runPowerPoint} from "../utils/powerPointUtil";
import Shape = PowerPoint.Shape;

export const rowLineName = "RowLine";
export const columnLineName = "ColumnLine";

const SLIDE_WIDTH = 960;
const SLIDE_HEIGHT = 540;
const SLIDE_MARGIN = 8;
const CONTENT_MARGIN = { top: 126, bottom: 60, right: 54, left: 58 };
const CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_MARGIN.top - CONTENT_MARGIN.bottom;
const CONTENT_WIDTH = SLIDE_WIDTH - CONTENT_MARGIN.right - CONTENT_MARGIN.left;

export async function deleteShapesByName(name: string) {
  await PowerPoint.run(async context => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("shapes");
    await context.sync();

    slide.shapes.items.forEach(shape => {
      if (shape.name === name) {
        shape.delete();
      }
    });

    await context.sync();
  });
}

async function getSingleSelectedShapeOrNull(context: PowerPoint.RequestContext) {
  let selectedShapes = context.presentation.getSelectedShapes();
  let clientResult = selectedShapes.getCount();
  await context.sync();
  let selectedShapesCount = clientResult.value;
  if (selectedShapesCount != 1) {
    return null;
  }

  let selectedShape = selectedShapes.getItemAt(0);
  return selectedShape.load();
}

export async function createRows(numberOfRows: number) {
  await runPowerPoint(async (powerPointContext) => {
    const singleSelectedShapeOrNull = await getSingleSelectedShapeOrNull(powerPointContext);
    if (singleSelectedShapeOrNull) {
      await createRowsForObject(numberOfRows, singleSelectedShapeOrNull, powerPointContext);
    } else {
      await createRowsForSlide(numberOfRows, powerPointContext);
    }
  });
}

export async function createColumns(numberOfRows: number) {
  await runPowerPoint(async (powerPointContext) => {
    const singleSelectedShapeOrNull = await getSingleSelectedShapeOrNull(powerPointContext);
    if (singleSelectedShapeOrNull) {
      await createColumnsForObject(numberOfRows, singleSelectedShapeOrNull, powerPointContext);
    } else {
      await createColumnsForSlide(numberOfRows, powerPointContext);
    }
  });
}

async function createRowsForSlide(numberOfRows: number, powerPointContext: PowerPoint.RequestContext) {
  const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;

  const lineDistance = CONTENT_HEIGHT / numberOfRows;
  let top = CONTENT_MARGIN.top;

  await renderRows(shapes, numberOfRows, top, lineDistance, SLIDE_WIDTH - SLIDE_MARGIN * 2, SLIDE_MARGIN);
  await powerPointContext.sync();
}

async function createColumnsForSlide(numberOfColumns: number, powerPointContext: PowerPoint.RequestContext) {
  const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;

  const lineDistance = CONTENT_WIDTH / numberOfColumns;

  let left = CONTENT_MARGIN.left;

  await renderColumns(shapes, numberOfColumns, left, lineDistance, SLIDE_HEIGHT - SLIDE_MARGIN * 2, SLIDE_MARGIN);
  await powerPointContext.sync();
}

async function createRowsForObject(numberOfColumns: number, selectedShape: Shape, powerPointContext: PowerPoint.RequestContext) {
  await powerPointContext.sync();
  const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;


  const lineDistance = selectedShape.height / numberOfColumns;
  const selectedShapeRight = selectedShape.left + selectedShape.width;
  const lineWidth = SLIDE_WIDTH - CONTENT_MARGIN.right - selectedShapeRight;

  let top = selectedShape.top;

  await renderRows(shapes, numberOfColumns, top, lineDistance, lineWidth, selectedShapeRight);
  await powerPointContext.sync();
}

async function createColumnsForObject(numberOfColumns: number, selectedShape: Shape, powerPointContext: PowerPoint.RequestContext) {
  await powerPointContext.sync();
  const shapes = powerPointContext.presentation.getSelectedSlides().getItemAt(0).shapes;

  const lineDistance = selectedShape.width / numberOfColumns;
  const selectedShapeBottom = selectedShape.top + selectedShape.height;
  const lineHeight = SLIDE_HEIGHT - CONTENT_MARGIN.bottom - selectedShapeBottom;

  let left = selectedShape.left;

  await renderColumns(shapes, numberOfColumns, left, lineDistance, lineHeight, selectedShapeBottom);
  await powerPointContext.sync();
}

async function renderRows(
    shapes: PowerPoint.ShapeCollection,
    numberOfRows: number, initialTop: number,
    lineDistance: number, lineWidth: number, left: number) {
  let top = initialTop;
  for (let _i = 0; _i <= numberOfRows; _i++) {
    const line = shapes.addLine(
        PowerPoint.ConnectorType.straight,
        { height: 0.5, left: left, top: top, width: lineWidth }
    );
    line.name = rowLineName;
    line.lineFormat.color = "#000000";

    top += lineDistance;
  }
}

async function renderColumns(
    shapes: PowerPoint.ShapeCollection,
    numberOfColumns: number, initialLeft: number,
    lineDistance: number, lineHeight: number, top: number) {
  let left = initialLeft;
  for (let _i = 0; _i <= numberOfColumns; _i++) {
    const line = shapes.addLine(
        PowerPoint.ConnectorType.straight,
        {height: lineHeight, left: left, top: top, width: 0.5}
    );
    line.name = columnLineName;
    line.lineFormat.color = "#000000";

    left += lineDistance;
  }
}