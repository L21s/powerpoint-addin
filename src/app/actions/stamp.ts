import {SLIDE_HEIGHT, SLIDE_WIDTH} from "../shared/consts";
import {StampPosition} from "../shared/enums";
import {StampOptions} from "../shared/types";

const STAMP_SHAPE_NAME = "Stamp";
const STAMP_SETTINGS_KEY = "stampOptions";
const STAMP_HORIZONTAL_PADDING = 20;
const STAMP_VERTICAL_PADDING = 20;
const STAMP_FONT_SIZE = 18;

export const DEFAULT_STAMP_TEXT = "ENTWURF";
export const DEFAULT_STAMP_BACKGROUND = "#d92d20";
export const DEFAULT_STAMP_TEXT_COLOR = "#ffffff";
export const DEFAULT_STAMP_POSITION = StampPosition.Top;

let syncHandlerAttached = false;
// Cached parsed copy of the persisted options, reused by the
// DocumentSelectionChanged handler to avoid re-parsing settings JSON on
// every selection change. Tri-state:
//   undefined — not yet read from document settings
//   null      — no stamp currently configured (fresh document, or after Delete)
//   StampOptions — the parsed options
let cachedOptions: StampOptions | null | undefined;

/**
 * Stamps every slide in the presentation. Called when the user clicks
 * "Hinzufügen". Persists the options so subsequent slide-change events
 * can keep new slides in sync.
 */
export async function addStamp(options: StampOptions) {
  saveStampOptions(options);

  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      await removeStampFromSlide(slide, context);
    }
    await context.sync();

    for (const slide of slides) {
      await addStampToSlide(slide, options, context);
    }

    await context.sync();
  });

  startStampSync();
}

/**
 * Removes the stamp from every slide and stops syncing new slides.
 * Called when the user clicks the trash button.
 */
export async function removeStamp() {
  stopStampSync();
  clearStampOptions();

  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      await removeStampFromSlide(slide, context);
    }

    await context.sync();
  });
}

export function getSavedStampOptions(): StampOptions | null {
  if (cachedOptions === undefined) {
    cachedOptions = readStampOptionsFromSettings();
  }
  return cachedOptions;
}

function readStampOptionsFromSettings(): StampOptions | null {
  const raw = Office.context.document.settings.get(STAMP_SETTINGS_KEY);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw as string) as Partial<StampOptions>;
    if (
      typeof parsed?.text !== "string" ||
      typeof parsed?.textColor !== "string" ||
      typeof parsed?.backgroundColor !== "string" ||
      !isStampPosition(parsed.position)
    ) {
      return null;
    }
    return {
      text: parsed.text,
      textColor: parsed.textColor,
      backgroundColor: parsed.backgroundColor,
      position: parsed.position,
    };
  } catch (error) {
    console.warn("Stamp: failed to parse persisted options; ignoring", error);
    return null;
  }
}

function isStampPosition(value: unknown): value is StampPosition {
  return (
    value === StampPosition.Top ||
    value === StampPosition.Bottom ||
    value === StampPosition.Left ||
    value === StampPosition.Right
  );
}

export function startStampSync() {
  if (syncHandlerAttached) return;
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSlideChanged,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        syncHandlerAttached = true;
      }
    }
  );
}

export function stopStampSync() {
  if (!syncHandlerAttached) return;
  Office.context.document.removeHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    {handler: onSlideChanged},
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        syncHandlerAttached = false;
      }
    }
  );
}

/**
 * Handler that fires whenever the selection in the document changes.
 * PowerPoint auto-selects newly inserted slides, so we use this as a
 * slide-added signal: if the currently selected slide is missing its
 * stamp, add it.
 */
function onSlideChanged() {
  addStampToSelectedSlideIfMissing().catch(() => {
    // Swallow errors from transient PowerPoint state; the next event will retry.
  });
}

async function addStampToSelectedSlideIfMissing(): Promise<void> {
  const options = getSavedStampOptions();
  if (!options) return;

  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedSlides();
    selected.load("items");
    await context.sync();

    const slide = selected.items[0];
    if (!slide) return;
    if (await stampExistsInSlide(slide, context)) return;

    await addStampToSlide(slide, options, context);
    await context.sync();
  });
}

/**
 * Adds a stamp shape to a single slide using the given options.
 * Caller is responsible for calling `context.sync()` to persist.
 */
async function addStampToSlide(
  slide: PowerPoint.Slide,
  options: StampOptions,
  context: PowerPoint.RequestContext
): Promise<void> {
  const shape = slide.shapes.addTextBox(options.text);
  shape.name = STAMP_SHAPE_NAME;
  shape.textFrame.wordWrap = false;
  shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
  // PowerPoint.js doesn't expose a shape-lock/protection API (only Excel
  // does via `lockAspectRatio`). A true edit-lock would require either
  // moving the shape onto the slide master or low-level OOXML manipulation
  // — both much larger changes. For now the stamp is freely editable; the
  // DocumentSelectionChanged handler at least re-adds it if it gets deleted.

  const range = shape.textFrame.textRange;
  // PowerPoint.js (as of 2026-04) exposes neither `Shape.rotation` nor a
  // vertical text-orientation setting on `TextFrame`, so vertical stamps
  // stack each character on its own line as a fallback.
  range.text = isVerticalPosition(options.position) ? toVerticalText(options.text) : options.text;
  range.font.color = options.textColor;
  range.font.bold = true;
  range.font.size = STAMP_FONT_SIZE;
  range.paragraphFormat.horizontalAlignment = "Center";

  shape.fill.setSolidColor(options.backgroundColor);

  await autoResizeShape(context, shape);
  positionStampShape(shape, options.position);
}

function toVerticalText(text: string): string {
  return text.split("").join("\n");
}

function isVerticalPosition(position: StampPosition): boolean {
  return position === StampPosition.Left || position === StampPosition.Right;
}

async function removeStampFromSlide(slide: PowerPoint.Slide, context: PowerPoint.RequestContext) {
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

function saveStampOptions(options: StampOptions) {
  Office.context.document.settings.set(STAMP_SETTINGS_KEY, JSON.stringify(options));
  Office.context.document.settings.saveAsync();
  cachedOptions = options;
}

function clearStampOptions() {
  Office.context.document.settings.remove(STAMP_SETTINGS_KEY);
  Office.context.document.settings.saveAsync();
  cachedOptions = null;
}

async function getSlides(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide[]> {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  return slides.items;
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
      shape.top = 0;
      break;
    case StampPosition.Bottom:
      shape.width += STAMP_HORIZONTAL_PADDING;
      shape.left = (SLIDE_WIDTH - shape.width) / 2;
      shape.top = SLIDE_HEIGHT - shape.height;
      break;
    case StampPosition.Left:
      shape.height += STAMP_VERTICAL_PADDING;
      shape.left = 0;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      break;
    case StampPosition.Right:
      shape.height += STAMP_VERTICAL_PADDING;
      shape.left = SLIDE_WIDTH - shape.width;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      break;
  }
}
