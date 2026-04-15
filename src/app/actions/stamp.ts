import {SLIDE_HEIGHT, SLIDE_WIDTH} from "../shared/consts";
import {StampPosition} from "../shared/enums";
import {StampOptions} from "../shared/types";

const STAMP_SHAPE_NAME = "Stamp";
const STAMP_SETTINGS_KEY = "stampOptions";
const STAMP_EDGE_OFFSET = 10;
const STAMP_HORIZONTAL_PADDING = 20;
const STAMP_VERTICAL_PADDING = 20;
const STAMP_FONT_SIZE = 18;
const STAMP_TEXT_COLOR = "#ffffff";
const SYNC_INTERVAL_MS = 2000;

export const DEFAULT_STAMP_TEXT = "ENTWURF";
export const DEFAULT_STAMP_BACKGROUND = "#d92d20";
export const DEFAULT_STAMP_POSITION = StampPosition.Top;

let syncTimer: number | null = null;

export async function addStamp(options: StampOptions) {
  saveStampOptions(options);

  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      await deleteStampFromSlide(slide, context);
    }
    await context.sync();

    for (const slide of slides) {
      const shape = createStampShape(slide);
      setStampText(shape, options);
      applyStampStyle(shape, options);
      await autoResizeShape(context, shape);
      positionStampShape(shape, options.position);
    }

    await context.sync();
  });

  startStampSync();
}

export async function removeStamp() {
  stopStampSync();
  clearStampOptions();

  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      await deleteStampFromSlide(slide, context);
    }

    await context.sync();
  });
}

export function getSavedStampOptions(): StampOptions | null {
  const raw = Office.context.document.settings.get(STAMP_SETTINGS_KEY);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw as string) as Partial<StampOptions>;
    if (typeof parsed?.text !== "string" || typeof parsed?.backgroundColor !== "string") {
      return null;
    }
    const position = isStampPosition(parsed.position) ? parsed.position : DEFAULT_STAMP_POSITION;
    return {text: parsed.text, backgroundColor: parsed.backgroundColor, position};
  } catch {
    return null;
  }
}

function isStampPosition(value: unknown): value is StampPosition {
  return value === StampPosition.Top || value === StampPosition.Left || value === StampPosition.Right;
}

export async function syncStampToAllSlides(): Promise<void> {
  const options = getSavedStampOptions();
  if (!options) return;

  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);
    const missing: PowerPoint.Slide[] = [];

    for (const slide of slides) {
      if (!(await stampExistsInSlide(slide, context))) {
        missing.push(slide);
      }
    }

    if (missing.length === 0) return;

    for (const slide of missing) {
      const shape = createStampShape(slide);
      setStampText(shape, options);
      applyStampStyle(shape, options);
      await autoResizeShape(context, shape);
      positionStampShape(shape, options.position);
    }

    await context.sync();
  });
}

export function startStampSync() {
  if (syncTimer !== null) return;
  syncTimer = window.setInterval(() => {
    syncStampToAllSlides().catch(() => {
      // Swallow errors from transient PowerPoint state; next tick will retry.
    });
  }, SYNC_INTERVAL_MS);
}

export function stopStampSync() {
  if (syncTimer !== null) {
    window.clearInterval(syncTimer);
    syncTimer = null;
  }
}

function saveStampOptions(options: StampOptions) {
  Office.context.document.settings.set(STAMP_SETTINGS_KEY, JSON.stringify(options));
  Office.context.document.settings.saveAsync();
}

function clearStampOptions() {
  Office.context.document.settings.remove(STAMP_SETTINGS_KEY);
  Office.context.document.settings.saveAsync();
}

async function getSlides(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide[]> {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  return slides.items;
}

function createStampShape(slide: PowerPoint.Slide): PowerPoint.Shape {
  const shape = slide.shapes.addTextBox(STAMP_SHAPE_NAME);
  shape.name = STAMP_SHAPE_NAME;
  shape.textFrame.wordWrap = false;
  shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
  return shape;
}

function setStampText(shape: PowerPoint.Shape, options: StampOptions) {
  const range = shape.textFrame.textRange;
  range.text = options.position === StampPosition.Top ? options.text : toVerticalText(options.text);
  range.font.color = STAMP_TEXT_COLOR;
  range.font.bold = true;
  range.font.size = STAMP_FONT_SIZE;
  range.paragraphFormat.horizontalAlignment = "Center";
}

function toVerticalText(text: string): string {
  return text.split("").join("\n");
}

function applyStampStyle(shape: PowerPoint.Shape, options: StampOptions) {
  shape.fill.setSolidColor(options.backgroundColor);
}

async function autoResizeShape(context: PowerPoint.RequestContext, shape: PowerPoint.Shape) {
  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
  shape.load(["width", "height"]);
  await context.sync();
  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
}

function positionStampShape(shape: PowerPoint.Shape, position: StampPosition) {
  switch (position) {
    case StampPosition.Top:
      shape.width += STAMP_HORIZONTAL_PADDING;
      shape.left = (SLIDE_WIDTH - shape.width) / 2;
      shape.top = STAMP_EDGE_OFFSET;
      break;
    case StampPosition.Left:
      shape.height += STAMP_VERTICAL_PADDING;
      shape.left = STAMP_EDGE_OFFSET;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      break;
    case StampPosition.Right:
      shape.height += STAMP_VERTICAL_PADDING;
      shape.left = SLIDE_WIDTH - shape.width - STAMP_EDGE_OFFSET;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      break;
  }
}

async function deleteStampFromSlide(slide: PowerPoint.Slide, context: PowerPoint.RequestContext) {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  shapes.items
    .filter((shape) => shape.name === STAMP_SHAPE_NAME)
    .forEach((shape) => shape.delete());
}

async function stampExistsInSlide(slide: PowerPoint.Slide, context: PowerPoint.RequestContext): Promise<boolean> {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  return shapes.items.some((shape) => shape.name === STAMP_SHAPE_NAME);
}
