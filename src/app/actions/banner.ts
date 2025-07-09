import {BannerPosition} from "../shared/enums";
import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
} from "../shared/consts";
import {BannerOptions} from "../shared/types";
import {addBannerButton, removeBannerButton} from "../taskpane";

const BANNER_SHAPE_NAME = "Banner";
const TOP_BANNER_WIDTH_PADDING = 15;
const SIDE_BANNER_HEIGHT_PADDING = 20;

export async function addBanner(options: BannerOptions) {
  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      const shape = createBannerShape(slide);
      setBannerText(shape, options);
      applyBannerStyle(shape, options);
      await autoResizeShape(context, shape);
      positionBannerShape(shape, options.position);
    }

    await context.sync();
  });
}

export async function removeBanner() {
  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      await deleteBannerFromSlide(slide, context);
    }

    await context.sync();
  });
}

export async function checkBannerExists(): Promise<boolean> {
  return PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      const exists = await bannerExistsInSlide(slide, context);
      if (exists) return true;
    }

    return false;
  });
}

export function toggleBannerButtons(bannerExists: boolean) {
  addBannerButton.style.display = bannerExists ? "none" : "inline-block";
  removeBannerButton.style.display = bannerExists ? "inline-block" : "none";
}

async function getSlides(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide[]> {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  return slides.items;
}

function createBannerShape(slide: PowerPoint.Slide): PowerPoint.Shape {
  const shape = slide.shapes.addTextBox(BANNER_SHAPE_NAME);
  shape.name = BANNER_SHAPE_NAME;
  shape.textFrame.wordWrap = false;
  shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
  return shape;
}

function setBannerText(shape: PowerPoint.Shape, options: BannerOptions) {
  const { text, textColor, position } = options;
  const range = shape.textFrame.textRange;

  range.text = position === BannerPosition.Top ? text : toVerticalText(text);
  range.font.color = textColor;
  range.paragraphFormat.horizontalAlignment = "Center";
}

function applyBannerStyle(shape: PowerPoint.Shape, options: BannerOptions) {
  shape.fill.setSolidColor(options.backgroundColor);
}

async function autoResizeShape(context: PowerPoint.RequestContext, shape: PowerPoint.Shape) {
  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
  shape.load(["width", "height"]);
  await context.sync();
  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
}

function positionBannerShape(shape: PowerPoint.Shape, position: BannerPosition) {
  switch (position) {
    case BannerPosition.Top:
      shape.left = (SLIDE_WIDTH - shape.width) / 2;
      shape.top = 0;
      shape.width += TOP_BANNER_WIDTH_PADDING;
      break;
    case BannerPosition.Left:
      shape.left = 0;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      shape.height += SIDE_BANNER_HEIGHT_PADDING;
      break;
    case BannerPosition.Right:
      shape.left = SLIDE_WIDTH - shape.width;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      shape.height += SIDE_BANNER_HEIGHT_PADDING;
      break;
  }
}

function toVerticalText(text: string): string {
  return text.split("").join("\n");
}

async function deleteBannerFromSlide(slide: PowerPoint.Slide, context: PowerPoint.RequestContext) {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  shapes.items
      .filter((shape) => shape.name === BANNER_SHAPE_NAME)
      .forEach((shape) => shape.delete());
}

async function bannerExistsInSlide(slide: PowerPoint.Slide, context: PowerPoint.RequestContext): Promise<boolean> {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  return shapes.items.some((shape) => shape.name === BANNER_SHAPE_NAME);
}
