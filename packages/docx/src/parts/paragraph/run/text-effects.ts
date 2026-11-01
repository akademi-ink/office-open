/**
 * Word 2010 run text effects (w14) — the `EG_RPrOpenType` and
 * `EG_RPrTextEffects` extension slot of `w:rPr`.
 *
 * These are the elements Word writes by default when a user applies
 * Format › Font › Advanced (ligatures, number form/spacing, stylistic sets) or
 * Text Effects (glow, shadow, reflection, 3-D format, text fill). They live in
 * the `w14` extension namespace, which a non-Word consumer may ignore
 * wholesale (that is what `mc:Ignorable="w14"` declares).
 *
 * Only the elements the model understands are typed here; anything else a
 * source carries (`w14:textOutline`, a future effect) falls back to
 * `RunStylePropertiesOptions.w14RawXml` and round-trips byte-faithfully.
 *
 * Reference: wml-2010.xsd, `EG_RPrOpenType`, `EG_RPrTextEffects`,
 * `CT_FillTextEffect`, `CT_Glow`, `CT_Shadow`, `CT_Reflection`, `CT_Props3D`,
 * `CT_Scene3D`, `CT_Ligatures`, `CT_NumForm`, `CT_NumSpacing`, `CT_StyleSet`.
 *
 * @module
 */

import {
  convertToEmu,
  emitAngle,
  emitPercent,
  parseAngle,
  parsePercentAttr,
  xsdCompoundLine,
  xsdLightRigDirection,
  xsdLineCap,
  xsdPenAlignment,
  xsdRectAlignment,
} from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type {
  BevelOptions,
  CameraPreset,
  ColorTransformOptions,
  CompoundLine,
  LightRigDirection,
  LightRigType,
  LineCap,
  PenAlignment,
  PresetDash,
  PresetMaterial,
  RectAlignment,
  SphereCoords,
} from "@office-open/core/drawing";
import {
  ANGLE_TRANSFORMS,
  BOOLEAN_TRANSFORMS,
  PERCENT_TRANSFORMS,
} from "@office-open/core/drawing";
import { attr, attrBool, attrNum, escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { stringifyElement } from "../../../util/stringify-element";

/** ST_Ligatures — which optional ligature classes the font may apply. */
export type LigaturesValue =
  | "none"
  | "standard"
  | "contextual"
  | "historical"
  | "discretional"
  | "standardContextual"
  | "standardHistorical"
  | "contextualHistorical"
  | "standardDiscretional"
  | "contextualDiscretional"
  | "historicalDiscretional"
  | "standardContextualHistorical"
  | "standardContextualDiscretional"
  | "standardHistoricalDiscretional"
  | "contextualHistoricalDiscretional"
  | "standardContextualHistoricalDiscretional";

/** ST_NumForm — figure style for numerals. */
export type NumFormValue = "default" | "lining" | "oldStyle";

/** ST_NumSpacing — spacing class for numerals. */
export type NumSpacingValue = "default" | "proportional" | "tabular";

/**
 * A w14 color: the reduced `EG_ColorChoice` the extension schema allows
 * (`w14:srgbClr` or `w14:schemeClr`, plus color transforms).
 */
export interface TextEffectColorOptions {
  /** sRGB hex value (w14:srgbClr `@w14:val`) — one of value/schemeColor. */
  value?: string;
  /** Theme color token, e.g. "accent1" (w14:schemeClr `@w14:val`). */
  schemeColor?: string;
  /** Color transforms in source order (integer percent, degrees, or switches). */
  transforms?: ColorTransformOptions;
}

/** One gradient stop for a w14 gradient fill (w14:gs). */
export interface TextGradientStopOptions {
  /** Stop position as integer percent (0 = start, 100 = end). */
  position: number;
  /** Stop color. */
  color: TextEffectColorOptions;
}

/** Linear shade of a w14 gradient (w14:lin, CT_LinearShadeProperties). */
export interface TextLinearShadeOptions {
  /** Gradient direction in degrees. */
  angle?: number;
  /** Scale the angle to the shape's bounding box (w14:scaled). */
  scaled?: boolean;
}

/** Path shade of a w14 gradient (w14:path, CT_PathShadeProperties). */
export interface TextPathShadeOptions {
  /** Shade path: circle, rectangle or shape outline. */
  path?: "circle" | "rect" | "shape";
  /** Rectangle the gradient fills toward, integer percent offsets. */
  fillToRect?: { left?: number; top?: number; right?: number; bottom?: number };
}

/**
 * Typed `w14:gradFill` (CT_GradientFillProperties) — the subset the w14
 * text-effect schema allows: 2-10 stops plus a linear or path shade.
 */
export interface TextGradientOptions {
  /** Gradient color stops (w14:gsLst/w14:gs, 2-10). */
  stops?: TextGradientStopOptions[];
  /** Linear shade (w14:lin) — mutually exclusive with `path`. */
  linear?: TextLinearShadeOptions;
  /** Path shade (w14:path) — mutually exclusive with `linear`. */
  path?: TextPathShadeOptions;
}

/** Text fill for `w14:textFill` (CT_FillTextEffect). */
export interface TextFillOptions {
  /** No fill (w14:noFill). */
  none?: true;
  /** Solid fill (w14:solidFill wrapping one color). */
  solid?: TextEffectColorOptions;
  /** Gradient fill (w14:gradFill) — typed stops and shade. */
  gradient?: TextGradientOptions;
  /**
   * Verbatim `w14:gradFill` XML for a source the typed model cannot express
   * (a degenerate stop list or an unknown child). Round-trip only — do not
   * hand-author.
   */
  gradientXml?: string;
}

/**
 * Text outline effect for `w14:textOutline` (CT_TextOutlineEffect) — outline
 * width/style plus the fill painted inside the stroke.
 */
export interface TextOutlineOptions {
  /** Outline width in EMUs (number) or UniversalMeasure. */
  width?: number | UniversalMeasure;
  /** Line cap style (a:ST_LineCap). */
  cap?: LineCap;
  /** Compound line style (a:ST_CompoundLine). */
  compound?: CompoundLine;
  /** Stroke alignment relative to the outline path (a:ST_PenAlignment). */
  alignment?: PenAlignment;
  /** Stroke fill (none/solid/gradient), same shape as the text fill. */
  fill?: TextFillOptions;
  /** Preset dash style (w14:prstDash). */
  dash?: PresetDash;
  /** Join style: round/bevel, or a miter limit as integer percent. */
  join?: "round" | "bevel" | { miterLimit: number };
}

/** Glow effect for `w14:glow` (CT_Glow). */
export interface TextGlowOptions {
  /** Glow radius in EMUs (number) or UniversalMeasure (mm/cm/in/pt/pc/pi/px). */
  radius?: number | UniversalMeasure;
  /** Glow color. */
  color: TextEffectColorOptions;
}

/** Shadow effect for `w14:shadow` (CT_Shadow). */
export interface TextShadowOptions {
  /** Blur radius in EMUs (number) or UniversalMeasure. */
  blurRadius?: number | UniversalMeasure;
  /** Shadow distance in EMUs (number) or UniversalMeasure. */
  distance?: number | UniversalMeasure;
  /** Direction angle in degrees (0–360). */
  direction?: number;
  /** Horizontal scale as integer percent (100 = 100%). */
  scaleX?: number;
  /** Vertical scale as integer percent. */
  scaleY?: number;
  /** Horizontal skew angle in degrees. */
  skewX?: number;
  /** Vertical skew angle in degrees. */
  skewY?: number;
  /** Shadow alignment (ST_RectAlignment). */
  alignment?: RectAlignment;
  /** Shadow color. */
  color: TextEffectColorOptions;
}

/** Reflection effect for `w14:reflection` (CT_Reflection) — attributes only. */
export interface TextReflectionOptions {
  /** Blur radius in EMUs (number) or UniversalMeasure. */
  blurRadius?: number | UniversalMeasure;
  /** Start opacity as integer percent (100 = fully opaque). */
  startAlpha?: number;
  /** Start position as integer percent (0-100). */
  startPosition?: number;
  /** End opacity as integer percent. */
  endAlpha?: number;
  /** End position as integer percent. */
  endPosition?: number;
  /** Distance from the text in EMUs (number) or UniversalMeasure. */
  distance?: number | UniversalMeasure;
  /** Direction angle in degrees. */
  direction?: number;
  /** Fade direction angle in degrees. */
  fadeDirection?: number;
  /** Horizontal scale as integer percent. */
  scaleX?: number;
  /** Vertical scale as integer percent. */
  scaleY?: number;
  /** Horizontal skew angle in degrees. */
  skewX?: number;
  /** Vertical skew angle in degrees. */
  skewY?: number;
  /** Reflection alignment (ST_RectAlignment). */
  alignment?: RectAlignment;
}

/** 3-D shape properties for `w14:props3d` (CT_Props3D). */
export interface TextProps3DOptions {
  /** Top bevel (w14:bevelT). */
  bevelT?: BevelOptions;
  /** Bottom bevel (w14:bevelB). */
  bevelB?: BevelOptions;
  /** Extrusion color (w14:extrusionClr). */
  extrusionColor?: TextEffectColorOptions;
  /** Contour color (w14:contourClr). */
  contourColor?: TextEffectColorOptions;
  /** Extrusion height in EMUs. */
  extrusionH?: number;
  /** Contour width in EMUs. */
  contourW?: number;
  /** Material preset (ST_PresetMaterialType). */
  prstMaterial?: PresetMaterial;
}

/** 3-D scene for `w14:scene3d` (CT_Scene3D) — camera and light rig. */
export interface TextScene3DOptions {
  /** Camera preset (w14:camera `@prst`, ST_PresetCameraType). */
  camera: CameraPreset;
  /** Light rig preset (w14:lightRig `@rig`, ST_LightRigType). */
  lightRig: LightRigType;
  /** Light direction (w14:lightRig `@dir`, ST_LightRigDirection). */
  lightDirection: LightRigDirection;
  /** Light rig rotation (w14:lightRig/w14:rot). */
  lightRigRotation?: SphereCoords;
}

/**
 * w14 run text effects. Every member is optional; a source that carries none
 * omits the field entirely.
 */
export interface TextEffectsOptions {
  /** Optional ligature classes the font may apply (w14:ligatures). */
  ligatures?: LigaturesValue;
  /** Numeral figure style (w14:numForm). */
  numForm?: NumFormValue;
  /** Numeral spacing (w14:numSpacing). */
  numSpacing?: NumSpacingValue;
  /** Stylistic sets to toggle (w14:stylisticSets/w14:styleSet). */
  stylisticSets?: { id: number; enabled?: boolean }[];
  /** Contextual alternates (w14:cntxtAlts on/off). */
  contextualAlternatives?: boolean;
  /** Glow effect (w14:glow). */
  glow?: TextGlowOptions;
  /** Shadow effect (w14:shadow). */
  shadow?: TextShadowOptions;
  /** Reflection effect (w14:reflection). */
  reflection?: TextReflectionOptions;
  /** Text fill (w14:textFill). */
  textFill?: TextFillOptions;
  /** Text outline (w14:textOutline) — stroke width/style and its fill. */
  textOutline?: TextOutlineOptions;
  /** 3-D shape properties (w14:props3d). */
  props3d?: TextProps3DOptions;
  /** 3-D scene — camera and lighting (w14:scene3d). */
  scene3d?: TextScene3DOptions;
}

// ── Color ──

/**
 * `EG_ColorTransform` members the w14 schema allows, spelled out rather than
 * interpolated: the vocabulary is finite, and an unknown child in a source
 * file must be skipped instead of leaking into the transform record.
 */
const COLOR_TRANSFORM_TAGS: Readonly<Record<string, string>> = {
  tint: "w14:tint",
  shade: "w14:shade",
  alpha: "w14:alpha",
  hueMod: "w14:hueMod",
  sat: "w14:sat",
  satOff: "w14:satOff",
  satMod: "w14:satMod",
  lum: "w14:lum",
  lumOff: "w14:lumOff",
  lumMod: "w14:lumMod",
};

const COLOR_TRANSFORM_KEYS = new Map(
  Object.entries(COLOR_TRANSFORM_TAGS).map(([key, tag]) => [tag, key]),
);

/** Emit the transform children of a w14 color in the caller's key order. */
function stringifyColorTransforms(transforms: ColorTransformOptions): string {
  const parts: string[] = [];
  for (const [name, value] of Object.entries(transforms)) {
    const tag = COLOR_TRANSFORM_TAGS[name];
    if (tag === undefined || value === undefined || value === false) continue;
    if (value === true) {
      parts.push(`<${tag}/>`);
      continue;
    }
    const raw = PERCENT_TRANSFORMS.has(name as never)
      ? emitPercent(value as number)
      : ANGLE_TRANSFORMS.has(name as never)
        ? emitAngle(value as number)
        : value;
    parts.push(`<${tag} w14:val="${raw}"/>`);
  }
  return parts.join("");
}

/** The w14 element name + attribute value for a color choice. */
function colorChoice(color: TextEffectColorOptions): { tag: string; val: string } {
  return color.schemeColor !== undefined
    ? { tag: "w14:schemeClr", val: color.schemeColor }
    : { tag: "w14:srgbClr", val: color.value ?? "auto" };
}

function stringifyColor(color: TextEffectColorOptions): string {
  const { tag, val } = colorChoice(color);
  const inner = color.transforms ? stringifyColorTransforms(color.transforms) : "";
  return inner
    ? `<${tag} w14:val="${escapeXml(val)}">${inner}</${tag}>`
    : `<${tag} w14:val="${escapeXml(val)}"/>`;
}

const COLOR_TAGS = new Set(["w14:srgbClr", "w14:schemeClr"]);

/** Parse a w14 color element (`w14:srgbClr` / `w14:schemeClr`). */
function parseColor(el: Element | undefined): TextEffectColorOptions | undefined {
  if (!el?.name || !COLOR_TAGS.has(el.name)) return undefined;
  const val = attr(el, "w14:val");
  const color: TextEffectColorOptions =
    el.name === "w14:schemeClr" ? { schemeColor: val ?? "" } : { value: val ?? "" };
  const transforms: Record<string, unknown> = {};
  for (const child of el.elements ?? []) {
    const key = child.name === undefined ? undefined : COLOR_TRANSFORM_KEYS.get(child.name);
    if (key === undefined) continue;
    const raw = attr(child, "w14:val");
    if (raw === undefined) {
      if (BOOLEAN_TRANSFORMS.has(key as never)) transforms[key] = true;
      continue;
    }
    if (PERCENT_TRANSFORMS.has(key as never)) transforms[key] = parsePercentAttr(raw) ?? 0;
    else if (ANGLE_TRANSFORMS.has(key as never)) transforms[key] = parseAngle(Number(raw));
    else if (BOOLEAN_TRANSFORMS.has(key as never)) continue;
    else transforms[key] = Number(raw);
  }
  if (Object.keys(transforms).length > 0) {
    color.transforms = transforms as ColorTransformOptions;
  }
  return color;
}

/** First child whose name is in `tags`. */
function childOf(el: Element, ...tags: string[]): Element | undefined {
  for (const child of el.elements ?? []) {
    if (child.name !== undefined && tags.includes(child.name)) return child;
  }
  return undefined;
}

// ── Emit ──

/** `w14:stylisticSets` — one `w14:styleSet` per toggle. */
function stringifyStylisticSets(sets: NonNullable<TextEffectsOptions["stylisticSets"]>): string {
  const inner = sets
    .map(
      (set) =>
        `<w14:styleSet w14:id="${set.id}"${set.enabled === undefined ? "" : ` w14:val="${set.enabled ? 1 : 0}"`}/>`,
    )
    .join("");
  return `<w14:stylisticSets>${inner}</w14:stylisticSets>`;
}

/** `w14:props3d` — CT_Props3D. */
function stringifyProps3D(opts: TextProps3DOptions): string {
  const children: string[] = [];
  const bevel = (tag: string, bevel: BevelOptions | undefined): void => {
    if (!bevel) return;
    const attrs: string[] = [];
    if (bevel.w !== undefined) attrs.push(`w14:w="${bevel.w}"`);
    if (bevel.h !== undefined) attrs.push(`w14:h="${bevel.h}"`);
    if (bevel.prst !== undefined) attrs.push(`w14:prst="${bevel.prst}"`);
    children.push(attrs.length > 0 ? `<w14:${tag} ${attrs.join(" ")}/>` : `<w14:${tag}/>`);
  };
  bevel("bevelT", opts.bevelT);
  bevel("bevelB", opts.bevelB);
  if (opts.extrusionColor) {
    children.push(`<w14:extrusionClr>${stringifyColor(opts.extrusionColor)}</w14:extrusionClr>`);
  }
  if (opts.contourColor) {
    children.push(`<w14:contourClr>${stringifyColor(opts.contourColor)}</w14:contourClr>`);
  }
  const attrs: string[] = [];
  if (opts.extrusionH !== undefined) attrs.push(`w14:extrusionH="${opts.extrusionH}"`);
  if (opts.contourW !== undefined) attrs.push(`w14:contourW="${opts.contourW}"`);
  if (opts.prstMaterial !== undefined) attrs.push(`w14:prstMaterial="${opts.prstMaterial}"`);
  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  return children.length > 0
    ? `<w14:props3d${attrStr}>${children.join("")}</w14:props3d>`
    : `<w14:props3d${attrStr}/>`;
}

/** `w14:scene3d` — CT_Scene3D (camera + light rig only; no backdrop). */
function stringifyScene3D(opts: TextScene3DOptions): string {
  const rot = opts.lightRigRotation
    ? `<w14:rot w14:lat="${emitAngle(opts.lightRigRotation.lat)}" w14:lon="${emitAngle(opts.lightRigRotation.lon)}" w14:rev="${emitAngle(opts.lightRigRotation.rev)}"/>`
    : "";
  const lightRig = rot
    ? `<w14:lightRig w14:rig="${opts.lightRig}" w14:dir="${xsdLightRigDirection.to(opts.lightDirection)}">${rot}</w14:lightRig>`
    : `<w14:lightRig w14:rig="${opts.lightRig}" w14:dir="${xsdLightRigDirection.to(opts.lightDirection)}"/>`;
  return `<w14:scene3d><w14:camera w14:prst="${opts.camera}"/>${lightRig}</w14:scene3d>`;
}

/** `w14:gsLst` — the gradient stop list (2-10 `w14:gs` children). */
function stringifyGradientStops(stops: readonly TextGradientStopOptions[]): string {
  const inner = stops
    .map(
      (stop) =>
        `<w14:gs w14:pos="${emitPercent(stop.position)}">${stringifyColor(stop.color)}</w14:gs>`,
    )
    .join("");
  return `<w14:gsLst>${inner}</w14:gsLst>`;
}

/** `w14:gradFill` — CT_GradientFillProperties (stops + lin/path shade). */
function stringifyGradientFill(gradient: TextGradientOptions): string {
  let inner = "";
  if (gradient.stops && gradient.stops.length > 0) {
    inner += stringifyGradientStops(gradient.stops);
  }
  if (gradient.linear) {
    const attrs: string[] = [];
    if (gradient.linear.angle !== undefined) {
      attrs.push(`w14:ang="${emitAngle(gradient.linear.angle)}"`);
    }
    if (gradient.linear.scaled !== undefined) {
      attrs.push(`w14:scaled="${gradient.linear.scaled ? 1 : 0}"`);
    }
    inner += attrs.length > 0 ? `<w14:lin ${attrs.join(" ")}/>` : "<w14:lin/>";
  } else if (gradient.path) {
    const path = gradient.path;
    const pathAttr = path.path !== undefined ? ` w14:path="${path.path}"` : "";
    let rect = "";
    if (path.fillToRect) {
      const r = path.fillToRect;
      const attrs: string[] = [];
      if (r.left !== undefined) attrs.push(`w14:l="${emitPercent(r.left)}"`);
      if (r.top !== undefined) attrs.push(`w14:t="${emitPercent(r.top)}"`);
      if (r.right !== undefined) attrs.push(`w14:r="${emitPercent(r.right)}"`);
      if (r.bottom !== undefined) attrs.push(`w14:b="${emitPercent(r.bottom)}"`);
      rect = attrs.length > 0 ? `<w14:fillToRect ${attrs.join(" ")}/>` : "<w14:fillToRect/>";
    }
    inner += rect ? `<w14:path${pathAttr}>${rect}</w14:path>` : `<w14:path${pathAttr}/>`;
  }
  return `<w14:gradFill>${inner}</w14:gradFill>`;
}

/**
 * The fill payload shared by `w14:textFill` and `w14:textOutline`
 * (EG_FillProperties): noFill, solidFill, or gradFill.
 */
function stringifyTextFillInner(fill: TextFillOptions): string {
  if (fill.gradientXml) return fill.gradientXml;
  if (fill.gradient) return stringifyGradientFill(fill.gradient);
  if (fill.solid) return `<w14:solidFill>${stringifyColor(fill.solid)}</w14:solidFill>`;
  if (fill.none) return "<w14:noFill/>";
  return "";
}

/** `w14:textOutline` — CT_TextOutlineEffect. */
function stringifyTextOutline(outline: TextOutlineOptions): string {
  let inner = outline.fill ? stringifyTextFillInner(outline.fill) : "";
  if (outline.dash !== undefined) {
    inner += `<w14:prstDash w14:val="${outline.dash}"/>`;
  }
  if (outline.join === "round") inner += "<w14:round/>";
  else if (outline.join === "bevel") inner += "<w14:bevel/>";
  else if (outline.join) inner += `<w14:miter w14:lim="${emitPercent(outline.join.miterLimit)}"/>`;
  const attrs: string[] = [];
  if (outline.width !== undefined) attrs.push(`w14:w="${convertToEmu(outline.width)}"`);
  if (outline.cap !== undefined) attrs.push(`w14:cap="${xsdLineCap.to(outline.cap)}"`);
  if (outline.compound !== undefined) {
    attrs.push(`w14:cmpd="${xsdCompoundLine.to(outline.compound)}"`);
  }
  if (outline.alignment !== undefined) {
    attrs.push(`w14:algn="${xsdPenAlignment.to(outline.alignment)}"`);
  }
  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  return inner
    ? `<w14:textOutline${attrStr}>${inner}</w14:textOutline>`
    : `<w14:textOutline${attrStr}/>`;
}

/** Emit the typed w14 text-effect children in schema order. */
export function stringifyTextEffects(opts: TextEffectsOptions): string {
  let out = "";
  if (opts.ligatures !== undefined) out += `<w14:ligatures w14:val="${opts.ligatures}"/>`;
  if (opts.numForm !== undefined) out += `<w14:numForm w14:val="${opts.numForm}"/>`;
  if (opts.numSpacing !== undefined) out += `<w14:numSpacing w14:val="${opts.numSpacing}"/>`;
  if (opts.stylisticSets && opts.stylisticSets.length > 0) {
    out += stringifyStylisticSets(opts.stylisticSets);
  }
  if (opts.contextualAlternatives !== undefined) {
    out += opts.contextualAlternatives ? "<w14:cntxtAlts/>" : '<w14:cntxtAlts w14:val="0"/>';
  }
  if (opts.glow) {
    const rad =
      opts.glow.radius === undefined ? "" : ` w14:rad="${convertToEmu(opts.glow.radius)}"`;
    out += `<w14:glow${rad}>${stringifyColor(opts.glow.color)}</w14:glow>`;
  }
  if (opts.shadow) {
    const s = opts.shadow;
    const attrs: string[] = [];
    if (s.blurRadius !== undefined) attrs.push(`w14:blurRad="${convertToEmu(s.blurRadius)}"`);
    if (s.distance !== undefined) attrs.push(`w14:dist="${convertToEmu(s.distance)}"`);
    if (s.direction !== undefined) attrs.push(`w14:dir="${emitAngle(s.direction)}"`);
    if (s.scaleX !== undefined) attrs.push(`w14:sx="${emitPercent(s.scaleX)}"`);
    if (s.scaleY !== undefined) attrs.push(`w14:sy="${emitPercent(s.scaleY)}"`);
    if (s.skewX !== undefined) attrs.push(`w14:kx="${emitAngle(s.skewX)}"`);
    if (s.skewY !== undefined) attrs.push(`w14:ky="${emitAngle(s.skewY)}"`);
    if (s.alignment !== undefined) attrs.push(`w14:algn="${xsdRectAlignment.to(s.alignment)}"`);
    const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
    out += `<w14:shadow${attrStr}>${stringifyColor(s.color)}</w14:shadow>`;
  }
  if (opts.reflection) {
    const r = opts.reflection;
    const attrs: string[] = [];
    if (r.blurRadius !== undefined) attrs.push(`w14:blurRad="${convertToEmu(r.blurRadius)}"`);
    if (r.startAlpha !== undefined) attrs.push(`w14:stA="${emitPercent(r.startAlpha)}"`);
    if (r.startPosition !== undefined) attrs.push(`w14:stPos="${emitPercent(r.startPosition)}"`);
    if (r.endAlpha !== undefined) attrs.push(`w14:endA="${emitPercent(r.endAlpha)}"`);
    if (r.endPosition !== undefined) attrs.push(`w14:endPos="${emitPercent(r.endPosition)}"`);
    if (r.distance !== undefined) attrs.push(`w14:dist="${convertToEmu(r.distance)}"`);
    if (r.direction !== undefined) attrs.push(`w14:dir="${emitAngle(r.direction)}"`);
    if (r.fadeDirection !== undefined) attrs.push(`w14:fadeDir="${emitAngle(r.fadeDirection)}"`);
    if (r.scaleX !== undefined) attrs.push(`w14:sx="${emitPercent(r.scaleX)}"`);
    if (r.scaleY !== undefined) attrs.push(`w14:sy="${emitPercent(r.scaleY)}"`);
    if (r.skewX !== undefined) attrs.push(`w14:kx="${emitAngle(r.skewX)}"`);
    if (r.skewY !== undefined) attrs.push(`w14:ky="${emitAngle(r.skewY)}"`);
    if (r.alignment !== undefined) attrs.push(`w14:algn="${xsdRectAlignment.to(r.alignment)}"`);
    out += attrs.length > 0 ? `<w14:reflection ${attrs.join(" ")}/>` : "<w14:reflection/>";
  }
  if (opts.textOutline) out += stringifyTextOutline(opts.textOutline);
  if (opts.textFill) {
    const inner = stringifyTextFillInner(opts.textFill);
    if (inner) out += `<w14:textFill>${inner}</w14:textFill>`;
  }
  if (opts.scene3d) out += stringifyScene3D(opts.scene3d);
  if (opts.props3d) out += stringifyProps3D(opts.props3d);
  return out;
}

// ── Parse ──

/** Read a w14 or plain attribute as EMU (numbers are already EMU). */
const emuAttr = (el: Element, name: string): number | undefined => attrNum(el, `w14:${name}`);

/** Read an angle attribute (OOXML 1/60000 degree) as degrees. */
function angleAttr(el: Element, name: string): number | undefined {
  const raw = attrNum(el, `w14:${name}`);
  return raw === undefined ? undefined : parseAngle(raw);
}

/** Read a percentage attribute (OOXML 1/1000 percent) as integer percent. */
function percentAttr(el: Element, name: string): number | undefined {
  const raw = attr(el, `w14:${name}`);
  return raw === undefined ? undefined : parsePercentAttr(raw);
}

function parseBevel(el: Element | undefined): BevelOptions | undefined {
  if (!el) return undefined;
  const bevel: BevelOptions = {};
  const w = attrNum(el, "w14:w");
  const h = attrNum(el, "w14:h");
  const prst = attr(el, "w14:prst");
  if (w !== undefined) bevel.w = w;
  if (h !== undefined) bevel.h = h;
  if (prst !== undefined) bevel.prst = prst as BevelOptions["prst"];
  return bevel;
}

/** Read one percentage attribute off a w14 element as integer percent. */
function gradientPercent(el: Element, name: string): number | undefined {
  const raw = attr(el, `w14:${name}`);
  return raw === undefined ? undefined : parsePercentAttr(raw);
}

/** `w14:gradFill` → typed stops + lin/path shade; undefined when empty. */
function parseGradientFill(el: Element): TextGradientOptions | undefined {
  const gradient: TextGradientOptions = {};
  const gsLst = childOf(el, "w14:gsLst");
  if (gsLst) {
    const stops: TextGradientStopOptions[] = [];
    for (const child of gsLst.elements ?? []) {
      if (child.name !== "w14:gs") continue;
      const position = gradientPercent(child, "pos");
      const color = parseColor(childOf(child, ...COLOR_TAGS));
      if (position === undefined || !color) continue;
      stops.push({ position, color });
    }
    if (stops.length > 0) gradient.stops = stops;
  }
  const lin = childOf(el, "w14:lin");
  if (lin) {
    const linear: TextLinearShadeOptions = {};
    const angle = angleAttr(lin, "ang");
    if (angle !== undefined) linear.angle = angle;
    const scaled = attrBool(lin, "w14:scaled");
    if (scaled !== undefined) linear.scaled = scaled;
    gradient.linear = linear;
  } else {
    const path = childOf(el, "w14:path");
    if (path) {
      const shade: TextPathShadeOptions = {};
      const pathType = attr(path, "w14:path");
      if (pathType === "circle" || pathType === "rect" || pathType === "shape") {
        shade.path = pathType;
      }
      const rect = childOf(path, "w14:fillToRect");
      if (rect) {
        const fillToRect: NonNullable<TextPathShadeOptions["fillToRect"]> = {};
        const left = gradientPercent(rect, "l");
        const top = gradientPercent(rect, "t");
        const right = gradientPercent(rect, "r");
        const bottom = gradientPercent(rect, "b");
        if (left !== undefined) fillToRect.left = left;
        if (top !== undefined) fillToRect.top = top;
        if (right !== undefined) fillToRect.right = right;
        if (bottom !== undefined) fillToRect.bottom = bottom;
        shade.fillToRect = fillToRect;
      }
      gradient.path = shade;
    }
  }
  return gradient.stops || gradient.linear || gradient.path ? gradient : undefined;
}

/**
 * Parse an EG_FillProperties payload (`w14:noFill` / `w14:solidFill` /
 * `w14:gradFill`) — shared by textFill and textOutline. A gradient the typed
 * model cannot express degrades to verbatim `gradientXml`.
 */
function parseTextFill(el: Element): TextFillOptions | undefined {
  if (childOf(el, "w14:noFill")) return { none: true };
  const solidFill = childOf(el, "w14:solidFill");
  if (solidFill) {
    const color = parseColor(childOf(solidFill, ...COLOR_TAGS));
    return color ? { solid: color } : undefined;
  }
  const gradient = childOf(el, "w14:gradFill");
  if (gradient) {
    const typed = parseGradientFill(gradient);
    return typed ? { gradient: typed } : { gradientXml: stringifyElement(gradient) };
  }
  return undefined;
}

/**
 * Parse one w14 extension child into `effects`. Returns false when the element
 * is not a modeled text effect — the caller keeps it as verbatim `w14RawXml`.
 */
export function parseTextEffectChild(el: Element, effects: TextEffectsOptions): boolean {
  switch (el.name) {
    case "w14:ligatures": {
      const val = attr(el, "w14:val");
      if (val !== undefined) effects.ligatures = val as LigaturesValue;
      return true;
    }
    case "w14:numForm": {
      const val = attr(el, "w14:val");
      if (val !== undefined) effects.numForm = val as NumFormValue;
      return true;
    }
    case "w14:numSpacing": {
      const val = attr(el, "w14:val");
      if (val !== undefined) effects.numSpacing = val as NumSpacingValue;
      return true;
    }
    case "w14:stylisticSets": {
      const sets: NonNullable<TextEffectsOptions["stylisticSets"]> = [];
      for (const child of el.elements ?? []) {
        if (child.name !== "w14:styleSet") continue;
        const id = attrNum(child, "w14:id");
        if (id === undefined) continue;
        const val = attr(child, "w14:val");
        sets.push(val === undefined ? { id } : { id, enabled: val !== "0" && val !== "false" });
      }
      if (sets.length > 0) effects.stylisticSets = sets;
      return true;
    }
    case "w14:cntxtAlts": {
      const val = attr(el, "w14:val");
      effects.contextualAlternatives = val === undefined || (val !== "0" && val !== "false");
      return true;
    }
    case "w14:glow": {
      const color = parseColor(childOf(el, ...COLOR_TAGS));
      if (!color) return false;
      const glow: TextGlowOptions = { color };
      const radius = emuAttr(el, "rad");
      if (radius !== undefined) glow.radius = radius;
      effects.glow = glow;
      return true;
    }
    case "w14:shadow": {
      const color = parseColor(childOf(el, ...COLOR_TAGS));
      if (!color) return false;
      const shadow: TextShadowOptions = { color };
      const blurRadius = emuAttr(el, "blurRad");
      const distance = emuAttr(el, "dist");
      const direction = angleAttr(el, "dir");
      const scaleX = percentAttr(el, "sx");
      const scaleY = percentAttr(el, "sy");
      const skewX = angleAttr(el, "kx");
      const skewY = angleAttr(el, "ky");
      const alignment = attr(el, "w14:algn");
      if (blurRadius !== undefined) shadow.blurRadius = blurRadius;
      if (distance !== undefined) shadow.distance = distance;
      if (direction !== undefined) shadow.direction = direction;
      if (scaleX !== undefined) shadow.scaleX = scaleX;
      if (scaleY !== undefined) shadow.scaleY = scaleY;
      if (skewX !== undefined) shadow.skewX = skewX;
      if (skewY !== undefined) shadow.skewY = skewY;
      if (alignment !== undefined)
        shadow.alignment = xsdRectAlignment.from(alignment) as RectAlignment;
      effects.shadow = shadow;
      return true;
    }
    case "w14:reflection": {
      const reflection: TextReflectionOptions = {};
      const blurRadius = emuAttr(el, "blurRad");
      const startAlpha = percentAttr(el, "stA");
      const startPosition = percentAttr(el, "stPos");
      const endAlpha = percentAttr(el, "endA");
      const endPosition = percentAttr(el, "endPos");
      const distance = emuAttr(el, "dist");
      const direction = angleAttr(el, "dir");
      const fadeDirection = angleAttr(el, "fadeDir");
      const scaleX = percentAttr(el, "sx");
      const scaleY = percentAttr(el, "sy");
      const skewX = angleAttr(el, "kx");
      const skewY = angleAttr(el, "ky");
      const alignment = attr(el, "w14:algn");
      if (blurRadius !== undefined) reflection.blurRadius = blurRadius;
      if (startAlpha !== undefined) reflection.startAlpha = startAlpha;
      if (startPosition !== undefined) reflection.startPosition = startPosition;
      if (endAlpha !== undefined) reflection.endAlpha = endAlpha;
      if (endPosition !== undefined) reflection.endPosition = endPosition;
      if (distance !== undefined) reflection.distance = distance;
      if (direction !== undefined) reflection.direction = direction;
      if (fadeDirection !== undefined) reflection.fadeDirection = fadeDirection;
      if (scaleX !== undefined) reflection.scaleX = scaleX;
      if (scaleY !== undefined) reflection.scaleY = scaleY;
      if (skewX !== undefined) reflection.skewX = skewX;
      if (skewY !== undefined) reflection.skewY = skewY;
      if (alignment !== undefined) {
        reflection.alignment = xsdRectAlignment.from(alignment) as RectAlignment;
      }
      effects.reflection = reflection;
      return true;
    }
    case "w14:textFill": {
      const fill = parseTextFill(el);
      if (!fill) return false;
      effects.textFill = fill;
      return true;
    }
    case "w14:textOutline": {
      const outline: TextOutlineOptions = {};
      const width = emuAttr(el, "w");
      if (width !== undefined) outline.width = width;
      const cap = attr(el, "w14:cap");
      if (cap !== undefined) outline.cap = xsdLineCap.from(cap) as LineCap;
      const compound = attr(el, "w14:cmpd");
      if (compound !== undefined) outline.compound = xsdCompoundLine.from(compound) as CompoundLine;
      const alignment = attr(el, "w14:algn");
      if (alignment !== undefined) {
        outline.alignment = xsdPenAlignment.from(alignment) as PenAlignment;
      }
      const fill = parseTextFill(el);
      if (fill) outline.fill = fill;
      const dash = childOf(el, "w14:prstDash");
      if (dash) {
        const val = attr(dash, "w14:val");
        if (val !== undefined) outline.dash = val as PresetDash;
      }
      if (childOf(el, "w14:round")) outline.join = "round";
      else if (childOf(el, "w14:bevel")) outline.join = "bevel";
      else {
        const miter = childOf(el, "w14:miter");
        if (miter) outline.join = { miterLimit: percentAttr(miter, "lim") ?? 800 };
      }
      effects.textOutline = outline;
      return true;
    }
    case "w14:scene3d": {
      const camera = childOf(el, "w14:camera");
      const lightRig = childOf(el, "w14:lightRig");
      const prst = camera ? attr(camera, "w14:prst") : undefined;
      const rig = lightRig ? attr(lightRig, "w14:rig") : undefined;
      const dir = lightRig ? attr(lightRig, "w14:dir") : undefined;
      if (prst === undefined || rig === undefined || dir === undefined) return false;
      const scene: TextScene3DOptions = {
        camera: prst as CameraPreset,
        lightRig: rig as LightRigType,
        lightDirection: xsdLightRigDirection.from(dir) as LightRigDirection,
      };
      const rot = lightRig ? childOf(lightRig, "w14:rot") : undefined;
      if (rot) {
        const lat = attrNum(rot, "w14:lat");
        const lon = attrNum(rot, "w14:lon");
        const rev = attrNum(rot, "w14:rev");
        if (lat !== undefined && lon !== undefined && rev !== undefined) {
          scene.lightRigRotation = {
            lat: parseAngle(lat),
            lon: parseAngle(lon),
            rev: parseAngle(rev),
          };
        }
      }
      effects.scene3d = scene;
      return true;
    }
    case "w14:props3d": {
      const props: TextProps3DOptions = {};
      const bevelT = parseBevel(childOf(el, "w14:bevelT"));
      const bevelB = parseBevel(childOf(el, "w14:bevelB"));
      if (bevelT) props.bevelT = bevelT;
      if (bevelB) props.bevelB = bevelB;
      const extrusionClr = childOf(el, "w14:extrusionClr");
      const contourClr = childOf(el, "w14:contourClr");
      const extrusionColor = extrusionClr
        ? parseColor(childOf(extrusionClr, ...COLOR_TAGS))
        : undefined;
      const contourColor = contourClr ? parseColor(childOf(contourClr, ...COLOR_TAGS)) : undefined;
      if (extrusionColor) props.extrusionColor = extrusionColor;
      if (contourColor) props.contourColor = contourColor;
      const extrusionH = emuAttr(el, "extrusionH");
      const contourW = emuAttr(el, "contourW");
      const prstMaterial = attr(el, "w14:prstMaterial");
      if (extrusionH !== undefined) props.extrusionH = extrusionH;
      if (contourW !== undefined) props.contourW = contourW;
      if (prstMaterial !== undefined) props.prstMaterial = prstMaterial as PresetMaterial;
      effects.props3d = props;
      return true;
    }
    default:
      return false;
  }
}

/** True when `effects` carries no modeled member. */
export function isEmptyTextEffects(effects: TextEffectsOptions): boolean {
  return Object.keys(effects).length === 0;
}
