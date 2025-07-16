import Shape = PowerPoint.Shape;
import { runPowerPoint } from "../shared/utils/powerPointUtil";
import { SLIDE_HEIGHT, SLIDE_MARGIN, SLIDE_WIDTH } from "../shared/consts";

const ROW_LINE_NAME = "RowLine";
const COLUMN_LINE_NAME = "ColumnLine";
const CONTENT_MARGIN = { top: 126, bottom: 60, right: 54, left: 58 };
const CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_MARGIN.top - CONTENT_MARGIN.bottom;
const CONTENT_WIDTH = SLIDE_WIDTH - CONTENT_MARGIN.right - CONTENT_MARGIN.left;
const LINE_COLOR = "#000000";
const LINE_THICKNESS = 0.5;

export async function createRows(count: number) {
  await runPowerPoint(async ctx => {
    const selectedShape = await getSingleSelectedShapeOrNull(ctx);
    if (selectedShape) {
      await createLinesForShape(ctx, count, selectedShape, true);
    } else {
      await createLinesForSlide(ctx, count, true);
    }
  });
}

export async function createColumns(count: number) {
  await runPowerPoint(async ctx => {
    const selectedShape = await getSingleSelectedShapeOrNull(ctx);
    if (selectedShape) {
      await createLinesForShape(ctx, count, selectedShape, false);
    } else {
      await createLinesForSlide(ctx, count, false);
    }
  });
}

export async function deleteRows() {
  await deleteShapesByName(ROW_LINE_NAME);
}

export async function deleteColumns() {
  await deleteShapesByName(COLUMN_LINE_NAME);
}

async function deleteShapesByName(name: string) {
  await PowerPoint.run(async ctx => {
    const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
    slide.load("shapes");
    await ctx.sync();

    slide.shapes.items.forEach(shape => {
      if (shape.name === name) shape.delete();
    });

    await ctx.sync();
  });
}

async function getSingleSelectedShapeOrNull(ctx: PowerPoint.RequestContext) {
  const selectedShapes = ctx.presentation.getSelectedShapes();
  const countResult = selectedShapes.getCount();
  await ctx.sync();
  const count = countResult.value;

  if (count !== 1) {
    return null;
  }

  const shape = selectedShapes.getItemAt(0);
  return shape.load();
}

async function createLinesForSlide(ctx: PowerPoint.RequestContext, count: number, isRow: boolean) {
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  const shapes = slide.shapes;

  if (isRow) {
    const rowHeight = CONTENT_HEIGHT / count;
    await renderRows(shapes, count, CONTENT_MARGIN.top, rowHeight, SLIDE_WIDTH - SLIDE_MARGIN * 2, SLIDE_MARGIN);
  } else {
    const columnWidth = CONTENT_WIDTH / count;
    await renderColumns(shapes, count, CONTENT_MARGIN.left, columnWidth, SLIDE_HEIGHT - SLIDE_MARGIN * 2, SLIDE_MARGIN);
  }
  await ctx.sync();
}

async function createLinesForShape(ctx: PowerPoint.RequestContext, count: number, shape: Shape, isRow: boolean) {
  await ctx.sync();
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  const shapes = slide.shapes;

  if (isRow) {
    const rowHeight = shape.height / count;
    const right = shape.left + shape.width;
    const width = SLIDE_WIDTH - CONTENT_MARGIN.right - right;
    await renderRows(shapes, count, shape.top, rowHeight, width, right);
  } else {
    const columnWidth = shape.width / count;
    const bottom = shape.top + shape.height;
    const height = SLIDE_HEIGHT - CONTENT_MARGIN.bottom - bottom;
    await renderColumns(shapes, count, shape.left, columnWidth, height, bottom);
  }
  await ctx.sync();
}

async function renderRows(
    shapes: PowerPoint.ShapeCollection,
    count: number,
    topStart: number,
    rowHeight: number,
    lineWidth: number,
    left: number
) {
  for (let i = 0; i <= count; i++) {
    const line = shapes.addLine(
        PowerPoint.ConnectorType.straight,
        { height: LINE_THICKNESS, left, top: topStart + i * rowHeight, width: lineWidth }
    );
    line.name = ROW_LINE_NAME;
    line.lineFormat.color = LINE_COLOR;
  }
}

async function renderColumns(
    shapes: PowerPoint.ShapeCollection,
    count: number,
    leftStart: number,
    columnWidth: number,
    lineHeight: number,
    top: number
) {
  for (let i = 0; i <= count; i++) {
    const line = shapes.addLine(
        PowerPoint.ConnectorType.straight,
        { height: lineHeight, left: leftStart + i * columnWidth, top, width: LINE_THICKNESS }
    );
    line.name = COLUMN_LINE_NAME;
    line.lineFormat.color = LINE_COLOR;
  }
}
