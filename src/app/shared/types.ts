import {ShapeType} from "./consts";
import {BannerPosition} from "./enums";

export type FetchIconResponse = {
  id: string;
  url: string;
};

export type Employee = {
  id: string;
  name: string;
};

export type BannerOptions = {
  text: string;
  textColor: string;
  backgroundColor: string;
  position: BannerPosition;
};

export type ShapeTypeKey = keyof typeof ShapeType; // "Rectangle" | "Ellipse" | ...
