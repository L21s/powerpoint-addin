import {getSelectedShape} from "./powerPointUtil";
import {ShapeType, ShapeTypeKey} from "./types";

export async function addColoredBackground() {
    await PowerPoint.run(async (context) => {

            // #0. Get slide and selected shape
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const selectedShape: PowerPoint.Shape = await getSelectedShape();

            // #1. Get selected background shape
            const shapeSelect       = document.getElementById('background-shape-selector') as HTMLSelectElement;
            const shapeSelectValue  = shapeSelect.value as ShapeTypeKey;

            // #2. Get selected background color
            const colorSelect       = document.getElementById('background-color-picker') as HTMLInputElement;
            const colorSelectValue  = colorSelect.value as string;

            // #3. Build background
            const background: PowerPoint.Shape = slide.shapes.addGeometricShape(ShapeType[shapeSelectValue ? shapeSelectValue : "Rectangle"]);
            background.left     = selectedShape.left
            background.top      = selectedShape.top;
            background.width    = selectedShape.width;
            background.height   = selectedShape.height;
            background.fill.setSolidColor(colorSelectValue ? colorSelectValue : 'lightgreen');
    })

    /**
     * Note: something like the code stated below should be used right here ... but sadly, the Powerpoint-Context still does not offer this.
     * This code is supposed to stay as reminder that maybe in future days Microsoft may offer this.
     * Better: Select Icon --> Bring to front.
     */
    /*await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shape = sheet.shapes.getItem("MyShape"); // use shape name or .getItemAt(0)
        shape.setZOrder("SendBackward");
        await context.sync();
    });
     */
}
