export interface Shadow {
  h: number;
  v: number;
  blur: number;
  color: string;
}

export interface ColorFill {
  type: "color";
  value: string;
}

export interface ImageFill {
  type: "image";
  value: {
    picBase64: string;
    opacity: number;
  };
}

export interface GradientFill {
  type: "gradient";
  value: {
    path: "line" | "circle" | "rect" | "shape";
    rot: number;
    colors: {
      pos: string;
      color: string;
    }[];
  };
}

export type Fill = ColorFill | ImageFill | GradientFill;

export interface Border {
  borderColor: string;
  borderWidth: number;
  borderType: "solid" | "dashed" | "dotted";
}

export interface BaseAttribute {
  id: string
  left: number;
  top: number;
  width: number;
  height: number;
  order: number;
}

export interface Shape extends BaseAttribute {
  type: "shape";
  borderColor: string;
  borderWidth: number;
  borderType: "solid" | "dashed" | "dotted";
  borderStrokeDasharray: string;
  shadow?: Shadow;
  fill: Fill;
  content: string;
  isFlipV: boolean;
  isFlipH: boolean;
  rotate: number;
  shapType: string;
  vAlign: string;
  path?: string;
  name: string;
}

export interface Text extends BaseAttribute {
  type: "text";
  borderColor: string;
  borderWidth: number;
  borderType: "solid" | "dashed" | "dotted";
  borderStrokeDasharray: string;
  shadow?: Shadow;
  fill: Fill;
  isFlipV: boolean;
  isFlipH: boolean;
  isVertical: boolean;
  rotate: number;
  content: string;
  vAlign: string;
  name: string;
}

export interface Image extends BaseAttribute {
  type: "image";
  src: string;
  rotate: number;
  isFlipH: boolean;
  isFlipV: boolean;
  rect?: {
    t?: number;
    b?: number;
    l?: number;
    r?: number;
  };
  geom: string;
  borderColor: string;
  borderWidth: number;
  borderType: "solid" | "dashed" | "dotted";
  borderStrokeDasharray: string;
  filters?: {
    sharpen?: number;
    colorTemperature?: number;
    saturation?: number;
    brightness?: number;
    contrast?: number;
  };
}

export interface TableCell {
  text: string;
  rowSpan?: number;
  colSpan?: number;
  vMerge?: number;
  hMerge?: number;
  fillColor?: string;
  fontColor?: string;
  fontBold?: boolean;
  borders: {
    top?: Border;
    bottom?: Border;
    left?: Border;
    right?: Border;
  };
}
export interface Table extends BaseAttribute {
  type: "table";
  data: TableCell[][];
  borders: {
    top?: Border;
    bottom?: Border;
    left?: Border;
    right?: Border;
  };
  rowHeights: number[];
  colWidths: number[];
}

export type ChartType =
  | "lineChart"
  | "line3DChart"
  | "barChart"
  | "bar3DChart"
  | "pieChart"
  | "pie3DChart"
  | "doughnutChart"
  | "areaChart"
  | "area3DChart"
  | "scatterChart"
  | "bubbleChart"
  | "radarChart"
  | "surfaceChart"
  | "surface3DChart"
  | "stockChart";

export interface ChartValue {
  x: string;
  y: number;
}
export interface ChartXLabel {
  [key: string]: string;
}
export interface ChartItem {
  key: string;
  values: ChartValue[];
  xlabels: ChartXLabel;
}
export type ScatterChartData = [number[], number[]];
export interface CommonChart extends BaseAttribute {
  type: "chart";
  data: ChartItem[];
  colors: string[];
  chartType: Exclude<ChartType, "scatterChart" | "bubbleChart">;
  barDir?: "bar" | "col";
  marker?: boolean;
  holeSize?: string;
  grouping?: string;
  style?: string;
}
export interface ScatterChart extends BaseAttribute {
  type: "chart";
  data: ScatterChartData;
  colors: string[];
  chartType: "scatterChart" | "bubbleChart";
}
export type Chart = CommonChart | ScatterChart;

export interface Video extends BaseAttribute {
  type: "video";
  blob?: string;
  src?: string;
}

export interface Audio extends BaseAttribute {
  type: "audio";
  blob: string;
}

export interface Diagram extends BaseAttribute {
  type: "diagram";
  elements: (Shape | Text)[];
}

export interface Math extends BaseAttribute {
  type: "math";
  latex: string;
  picBase64: string;
  text?: string;
}

export type BaseElement =
  | Shape
  | Text
  | Image
  | Table
  | Chart
  | Video
  | Audio
  | Diagram
  | Math;

export interface Group extends BaseAttribute {
  type: "group";
  rotate: number;
  elements: BaseElement[];
  isFlipH: boolean;
  isFlipV: boolean;
}
export type Element = BaseElement | Group;

export interface SlideTransition {
  type: string;
  duration: number;
  direction: string | null;
}

export interface Slide {
  fill: Fill;
  elements: Element[];
  layoutElements: Element[];
  note: string;
  transition?: SlideTransition | null;
}

export interface Options {
  slideFactor?: number;
  fontsizeFactor?: number;
}

export const pptxToJson: (
  file: ArrayBuffer,
  uploadFun: (
    fileBlob: Blob,
    fileExt: string
  ) => Promise<{ url: string; fileSize: number }> | null
) => Promise<{
  slides: Slide[];
  themeColors: string[];
  size: {
    width: number;
    height: number;
  };
}>;
