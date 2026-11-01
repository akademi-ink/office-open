/**
 * Word extension registry — the single source of truth for the Word extension
 * namespaces this engine declares, emits into, and passes through.
 *
 * Word puts every feature added after Office 2007 in an extension namespace
 * (`w14` for Word 2010 through `w16sdtfl` for Word 2024) and lists the ones a
 * consumer may skip in a root-level `mc:Ignorable` attribute (MCE, ECMA-376
 * Part 3). Two facts must stay in lockstep with the emitters:
 *
 *   1. prefix → namespace URI: {@link DocumentAttributeNamespaces} owns the
 *      URIs, this table adds the Word release and the feature it carries;
 *   2. the ignorable set: a root declares exactly the prefixes it also
 *      declares as `xmlns:*` and marks them ignorable, so Word can drop the
 *      ones it does not understand instead of rejecting the part.
 *
 * `mc:Ignorable` is derived from the prefix list, never hand-written: a part
 * that starts emitting a new extension namespace declares the prefix once and
 * passes its declaration list through {@link extensionIgnorableAttribute}.
 *
 * ## Coverage note (w14)
 *
 * Word 2010's 59 element names are split three ways:
 *   - **structured** — `w14:ligatures`, `w14:numForm`, `w14:numSpacing`,
 *     `w14:stylisticSets`, `w14:glow`, `w14:shadow`, `w14:reflection`,
 *     `w14:textFill`, `w14:textOutline`, `w14:props3d`, `w14:scene3d`,
 *     `w14:cntxtAlts` round-trip through the typed
 *     `RunStylePropertiesOptions.textEffects` model;
 *   - **preserved verbatim** — the remaining `w14:*` run-property children
 *     survive in `w14RawXml` (byte-faithful, not editable);
 *   - **modeled elsewhere** — `w14:checkbox` in `w:sdtPr`, `w14:paraId` /
 *     `w14:textId` on paragraphs, `w14:docId` in settings, and the conflict
 *     range markers have their own structured fields.
 *
 * Unknown Word namespaces (a future `word/2026/wordml`) are recognized by
 * {@link isWordExtensionNamespace} so a source that declares one keeps its
 * declaration and ignorable marking on passthrough instead of being stripped.
 *
 * @module
 */

import type { DocumentAttributeNamespace } from "./document-attributes";
import { DocumentAttributeNamespaces, documentNamespaceAttributes } from "./document-attributes";

/** One Word extension namespace and the feature it carries. */
export interface WordExtension {
  /** XML prefix (`w14`, `w16cex`, …) — the `mc:Ignorable` token. */
  prefix: DocumentAttributeNamespace;
  /** Word release that introduced the namespace. */
  word: string;
  /** Feature summary: what a consumer loses when the namespace is ignored. */
  feature: string;
  /** Whether Word marks the prefix `mc:Ignorable` on the roots it writes. */
  ignorable: boolean;
}

/**
 * Word extension namespaces, oldest first. Element inventories come from the
 * pinned {@link WORD_EXTENSION_SCHEMAS schema set} (`ooxml-schemas/microsoft`).
 */
export const WORD_EXTENSIONS: readonly WordExtension[] = [
  {
    prefix: "w14",
    word: "Word 2010",
    feature: "run text effects, sdt checkbox, conflict ranges, paragraph ids, docId",
    ignorable: true,
  },
  {
    prefix: "wp14",
    word: "Word 2010",
    feature: "drawing anchors: relative size, editId, anchorId",
    ignorable: true,
  },
  {
    prefix: "w15",
    word: "Word 2013",
    feature: "comment infrastructure (commentsEx, people), repeating sections",
    ignorable: true,
  },
  {
    prefix: "w16se",
    word: "Word 2015",
    feature: "Office 2016 symbol run (w16se:sym)",
    ignorable: true,
  },
  {
    prefix: "w16cid",
    word: "Word 2016",
    feature: "durable comment ids (commentsIds.xml)",
    ignorable: true,
  },
  {
    prefix: "w16",
    word: "Word 2018",
    feature: "sdt data binding extensions",
    ignorable: true,
  },
  {
    prefix: "w16cex",
    word: "Word 2018",
    feature: "extensible comments (commentsExtensible.xml)",
    ignorable: true,
  },
  {
    prefix: "w16sdtdh",
    word: "Word 2020",
    feature: "sdt content hash (w16sdtdh:storeItemChecksum)",
    ignorable: true,
  },
  {
    prefix: "w16du",
    word: "Word 2023",
    feature: "durable UTC timestamps (w16du:dateUtc)",
    ignorable: true,
  },
  {
    prefix: "w16sdtfl",
    word: "Word 2024",
    feature: "sdt formatting lock (w16sdtfl:formattingAllowed)",
    ignorable: true,
  },
];

/** Schema file backing each registry entry — the pinned spec sources. */
export const WORD_EXTENSION_SCHEMAS: Readonly<Record<string, string>> = {
  w14: "ooxml-schemas/microsoft/wml-2010.xsd",
  // wp14 lives in the DrawingML wordprocessingDrawing extension namespace,
  // which the pinned schema set does not ship separately.
  wp14: "ooxml-schemas/microsoft/wml-2010.xsd",
  w15: "ooxml-schemas/microsoft/wml-2012.xsd",
  w16se: "ooxml-schemas/microsoft/wml-symex-2015.xsd",
  w16cid: "ooxml-schemas/microsoft/wml-cid-2016.xsd",
  w16: "ooxml-schemas/microsoft/wml-2018.xsd",
  w16cex: "ooxml-schemas/microsoft/wml-cex-2018.xsd",
  w16sdtdh: "ooxml-schemas/microsoft/wml-sdtdatahash-2020.xsd",
  w16du: "ooxml-schemas/microsoft/wml-word16du-2023.xsd",
  w16sdtfl: "ooxml-schemas/microsoft/wml-sdtformatlock-2024.xsd",
};

const EXTENSIONS_BY_PREFIX = new Map<string, WordExtension>(
  WORD_EXTENSIONS.map((extension) => [extension.prefix, extension]),
);

/** Registry entry for a prefix, or undefined when it is not an extension. */
export const wordExtension = (prefix: string): WordExtension | undefined =>
  EXTENSIONS_BY_PREFIX.get(prefix);

/** True when a registered extension is safe for a consumer to skip. */
export const isIgnorableExtension = (prefix: string): boolean =>
  wordExtension(prefix)?.ignorable === true;

/**
 * Word extension namespace URIs follow
 * `http://schemas.microsoft.com/office/word/<year>/…`; matching the pattern
 * recognizes a namespace newer than the registry (Word 2026+) so its
 * declaration and ignorable marking survive a round-trip.
 */
const WORD_EXTENSION_NAMESPACE_RE = /^http:\/\/schemas\.microsoft\.com\/office\/word\/\d{4}\//;

/** True for any Word extension namespace, registered or newer. */
export const isWordExtensionNamespace = (namespace: string): boolean =>
  WORD_EXTENSION_NAMESPACE_RE.test(namespace);

/**
 * Space-separated ignorable prefixes for a root that declares `prefixes`,
 * derived from the registry in declaration order. Empty string when the root
 * carries no ignorable extension namespace — the value form for callers
 * building an attribute map.
 */
export const extensionIgnorablePrefixes = (prefixes: readonly string[]): string =>
  prefixes.filter(isIgnorableExtension).join(" ");

/**
 * `mc:Ignorable` attribute (leading space included) for a root that declares
 * `prefixes`. Empty string when there is nothing to ignore, so callers
 * concatenate it unconditionally after their `xmlns:*` declarations.
 */
export const extensionIgnorableAttribute = (prefixes: readonly string[]): string => {
  const ignorable = extensionIgnorablePrefixes(prefixes);
  return ignorable ? ` mc:Ignorable="${ignorable}"` : "";
};

/**
 * Root attributes for a part: the `xmlns:*` declarations for `prefixes`
 * followed by the registry-derived `mc:Ignorable` for the ignorable subset.
 * Every emitter builds its root header through here so a new extension
 * namespace is declared once and marked ignorable in the same step.
 */
export const rootNamespaceAttributes = (prefixes: readonly DocumentAttributeNamespace[]): string =>
  documentNamespaceAttributes(prefixes) + extensionIgnorableAttribute(prefixes);

/**
 * Prefixes declared on `word/styles.xml` — shared by the default-styles factory
 * template and the round-tripped styles root, which must agree or the same
 * document emits two different root headers.
 */
export const STYLES_PREFIXES = ["mc", "r", "w", "w14", "w15"] as const;

/** Namespace URIs for the registered extensions, in registry order. */
export const WORD_EXTENSION_NAMESPACES: readonly string[] = WORD_EXTENSIONS.map(
  (extension) => DocumentAttributeNamespaces[extension.prefix],
);
