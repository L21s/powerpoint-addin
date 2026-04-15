import {ShapeType} from "./consts";
import {StampPosition} from "./enums";

export type FetchIconResponse = {
  id: string;
  url: string;
};

export type Employee = {
  id: string;
  name: string;
};

export type StampOptions = {
  text: string;
  textColor: string;
  backgroundColor: string;
  position: StampPosition;
};

export type ShapeTypeKey = keyof typeof ShapeType; // "Rectangle" | "Ellipse" | ...
