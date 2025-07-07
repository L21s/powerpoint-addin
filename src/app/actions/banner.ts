const BANNER_SHAPE_NAME = "CustomBanner";
const SLIDE_WIDTH = 960;
const SLIDE_HEIGHT = 540;

interface BannerOptions {
  text: string;
  textColor: string;
  backgroundColor: string;
  position: "Top" | "Left" | "Right";
}

export async function addBanner(options: BannerOptions) {
  await PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      const shape = createBannerShape(slide);
      configureBannerText(shape, options);
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
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      shapes.items
          .filter((shape) => shape.name === BANNER_SHAPE_NAME)
          .forEach((shape) => shape.delete());
    }

    await context.sync();
  });
}

export async function checkBannerExists(): Promise<boolean> {
  return PowerPoint.run(async (context) => {
    const slides = await getSlides(context);

    for (const slide of slides) {
      const shapes = slide.shapes;
      shapes.load("items/name");
      await context.sync();

      const hasBanner = shapes.items.some((shape) => shape.name === BANNER_SHAPE_NAME);
      if (hasBanner) return true;
    }

    return false;
  });
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
  shape.tags.add("banner", "true");
  return shape;
}

function configureBannerText(shape: PowerPoint.Shape, options: BannerOptions) {
  const { text, textColor, position } = options;
  const range = shape.textFrame.textRange;

  range.text = position === "Top" ? text : toVerticalText(text);
  range.font.color = textColor;

  range.paragraphFormat.horizontalAlignment = "Center";
  shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
  shape.textFrame.wordWrap = false;
}

function applyBannerStyle(shape: PowerPoint.Shape, options: BannerOptions) {
  shape.fill.setSolidColor(options.backgroundColor);
}

async function autoResizeShape(context: PowerPoint.RequestContext, shape: PowerPoint.Shape) {
  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
  await context.sync();

  shape.load(["width", "height"]);
  await context.sync();

  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
}

function positionBannerShape(shape: PowerPoint.Shape, position: BannerOptions["position"]) {
  switch (position) {
    case "Top":
      shape.left = (SLIDE_WIDTH - shape.width) / 2;
      shape.top = 0;
      shape.width += 15;
      break;
    case "Left":
      shape.left = 0;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      shape.height += 20;
      break;
    case "Right":
      shape.left = SLIDE_WIDTH - shape.width;
      shape.top = (SLIDE_HEIGHT - shape.height) / 2;
      shape.height += 20;
      break;
  }
}

function toVerticalText(text: string): string {
  return text.split("").join("\n");
}
