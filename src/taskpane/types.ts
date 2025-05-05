export type FetchIconResponse = {
  id: string;
  url: string;
};


export const ShapeType = {
  Rectangle: PowerPoint.GeometricShapeType.rectangle,
  Ellipse: PowerPoint.GeometricShapeType.ellipse,
  Diamond: PowerPoint.GeometricShapeType.diamond,
  Triangle: PowerPoint.GeometricShapeType.triangle,
  Pentagon: PowerPoint.GeometricShapeType.pentagon,
  Hexagon: PowerPoint.GeometricShapeType.hexagon,
} as const;

export type ShapeTypeKey = keyof typeof ShapeType; // "Rectangle" | "Ellipse" | ...
export type ShapeTypeValue = (typeof ShapeType)[ShapeTypeKey]; // PowerPoint.GeometricShapeType