import {SLIDE_HEIGHT, SLIDE_WIDTH} from "../shared/consts";
import {StorerPosition} from "../shared/enums";
import {StorerOptions} from "../shared/types";

const STORER_SHAPE_NAME = "Storer";
const STORER_SETTINGS_KEY = "storerOptions";
const STORER_EDGE_OFFSET = 10;
const STORER_HORIZONTAL_PADDING = 20;
const STORER_VERTICAL_PADDING = 20;
const STORER_FONT_SIZE = 18;
const STORER_TEXT_COLOR = "#ffffff";
const SYNC_INTERVAL_MS = 2000;

export const DEFAULT_STORER_TEXT = "ENTWURF";
export const DEFAULT_STORER_BACKGROUND = "#d92d20";
export const DEFAULT_STORER_POSITION = StorerPosition.Top;

let syncTimer: number | null = null;

export async function addStorer(options: StorerOptions) {
  saveStorerOptions(options);

  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      await deleteStorerFromSlide(slide, context);
    }
    await context.sync();

    for (const slide of slides) {
      const shape = createStorerShape(slide);
      setStorerText(shape, options);
      applyStorerStyle(shape, options);
      await autoResizeShape(context, shape);
      positionStorerShape(shape, options.position);
    }

    await context.sync();
  });

  startStorerSync();
}

export async function removeStorer() {
  stopStorerSync();
  clearStorerOptions();

  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      await deleteStorerFromSlide(slide, context);
    }

    await context.sync();
  });
}

export function getSavedStorerOptions(): StorerOptions | null {
  const raw = Office.context.document.settings.get(STORER_SETTINGS_KEY);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw as string) as Partial<StorerOptions>;
    if (typeof parsed?.text !== "string" || typeof parsed?.backgroundColor !== "string") {
      return null;
    }
    const position = isStorerPosition(parsed.position) ? parsed.position : DEFAULT_STORER_POSITION;
    return {text: parsed.text, backgroundColor: parsed.backgroundColor, position};
  } catch {
    return null;
  }
}

function isStorerPosition(value: unknown): value is StorerPosition {
  return value === StorerPosition.Top || value === StorerPosition.Left || value === StorerPosition.Right;
}

export async function syncStorerToAllSlides(): Promise<void> {
  const options = getSavedStorerOptions();
  if (!options) return;

  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);
    const missing: PowerPoint.Slide[] = [];

    for (const slide of slides) {
      if (!(await storerExistsInSlide(slide, context))) {
        missing.push(slide);
      }
    }

    if (missing.length === 0) return;

    for (const slide of missing) {
      const shape = createStorerShape(slide);
      setStorerText(shape, options);
      applyStorerStyle(shape, options);
      await autoResizeShape(context, shape);
      positionStorerShape(shape, options.position);
    }

    await context.sync();
  });
}

export function startStorerSync() {
  if (syncTimer !== null) return;
  syncTimer = window.setInterval(() => {
    syncStorerToAllSlides().catch(() => {
      // Swallow errors from transient PowerPoint state; next tick will retry.
    });
  }, SYNC_INTERVAL_MS);
}

export function stopStorerSync() {
  if (syncTimer !== null) {
    window.clearInterval(syncTimer);
    syncTimer = null;
  }
}

function saveStorerOptions(options: StorerOptions) {
  Office.context.document.settings.set(STORER_SETTINGS_KEY, JSON.stringify(options));
  Office.context.document.settings.saveAsync();
}

function clearStorerOptions() {
  Office.context.document.settings.remove(STORER_SETTINGS_KEY);
  Office.context.document.settings.saveAsync();
}

async function getSlides(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide[]> {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  return slides.items;
}

function createStorerShape(slide: PowerPoint.Slide): PowerPoint.Shape {
  const shape = slide.shapes.addTextBox(STORER_SHAPE_NAME);
  shape.name = STORER_SHAPE_NAME;
  shape.textFrame.wordWrap = false;
  shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
  return shape;
}

function setStorerText(shape: PowerPoint.Shape, options: StorerOptions) {
  const range = shape.textFrame.textRange;
  range.text = options.position === StorerPosition.Top ? options.text : toVerticalText(options.text);
  range.font.color = STORER_TEXT_COLOR;
  range.font.bold = true;
  range.font.size = STORER_FONT_SIZE;
  range.paragraphFormat.horizontalAlignment = "Center";
}

function toVerticalText(text: string): string {
  return text.split("").join("\n");
}

function applyStorerStyle(shape: PowerPoint.Shape, options: StorerOptions) {
  shape.fill.setSolidColor(options.backgroundColor);
}

async function autoResizeShape(context: PowerPoint.RequestContext, shape: PowerPoint.Shape) {
  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
  shape.load(["width", "height"]);
  await context.sync();
  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
}

function positionStorerShape(shape: PowerPoint.Shape, position: StorerPosition) {
  switch (position) {
    case StorerPosition.Top:
      shape.width += STORER_HORIZONTAL_PADDING;
      shape.left = (SLIDE_WIDTH - shape.width) / 2;
      shape.top = STORER_EDGE_OFFSET;
      break;
    case StorerPosition.Left:
      shape.height += STORER_VERTICAL_PADDING;
      shape.left = STORER_EDGE_OFFSET;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      break;
    case StorerPosition.Right:
      shape.height += STORER_VERTICAL_PADDING;
      shape.left = SLIDE_WIDTH - shape.width - STORER_EDGE_OFFSET;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      break;
  }
}

async function deleteStorerFromSlide(slide: PowerPoint.Slide, context: PowerPoint.RequestContext) {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  shapes.items
    .filter((shape) => shape.name === STORER_SHAPE_NAME)
    .forEach((shape) => shape.delete());
}

async function storerExistsInSlide(slide: PowerPoint.Slide, context: PowerPoint.RequestContext): Promise<boolean> {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  return shapes.items.some((shape) => shape.name === STORER_SHAPE_NAME);
}
