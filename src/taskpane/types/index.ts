export interface Swatch {
  shade: string;
  hex: string;
  rgb: string;
}

export interface Color {
  name: string;
  swatches: Swatch[];
}
