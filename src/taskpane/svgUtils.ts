import {getSelectedShape} from "./powerPointUtil";



export async function addDefinedBackgroundToSVGShape(shapeType: ShapeTypeKey) {
    await PowerPoint.run(async (context) => {

        const selectedShape: PowerPoint.Shape = await getSelectedShape();
        const slide = context.presentation.getSelectedSlides().getItemAt(0);

        const background: PowerPoint.Shape = slide.shapes.addGeometricShape(ShapeType[shapeType]);
        background.left     = selectedShape.left
        background.top      = selectedShape.top;
        background.width    = selectedShape.width;
        background.height   = selectedShape.height;
        background.fill.setSolidColor("lightgreen");

    })

    /**
     * Note: something like the code stated below should be used right here ... but sadly, the Powerpoint-Context still does not offer this .
     * This code is supposed to stay as reminder that maybe in future days Microsoft may offer this.
     */
    /*await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const shape = sheet.shapes.getItem("MyShape"); // use shape name or .getItemAt(0)
        shape.setZOrder("SendBackward");
        await context.sync();
    });
     */
}

const ShapeType = {
    Rectangle: PowerPoint.GeometricShapeType.rectangle,
    Ellipse: PowerPoint.GeometricShapeType.ellipse,
    Diamond: PowerPoint.GeometricShapeType.diamond,
    Triangle: PowerPoint.GeometricShapeType.triangle,
    Parallelogram: PowerPoint.GeometricShapeType.parallelogram,
} as const;

export type ShapeTypeKey = keyof typeof ShapeType; // "Rectangle" | "Ellipse" | ...
export type ShapeTypeValue = (typeof ShapeType)[ShapeTypeKey]; // PowerPoint.GeometricShapeType
