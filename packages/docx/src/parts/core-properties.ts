/**
 * Core Properties module for WordprocessingML documents.
 *
 * Provides the DocumentOptions interface for document metadata.
 * XML generation is handled by the descriptor pipeline (corePropertiesDesc).
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * @module
 */
import type {
  ContentTypesInput,
  DataType,
  DigitalSignatureInfo,
  EncryptedContainerOptions,
} from "@office-open/core";
import type { BibliographyOptions } from "@parts/bibliography";
import type { CommentExtendedOptions } from "@parts/comments-extended";
import type { CommentExtensibleOptions } from "@parts/comments-extensible";
import type { CommentIdOptions } from "@parts/comments-ids";
import type { EmbeddedFontOptions } from "@parts/fonts/font-table";
import type { GlossaryDocumentOptions } from "@parts/glossary-document";
import type { CommentOptions } from "@parts/paragraph/run/comment-run";
import type { CommentPersonOptions } from "@parts/people";
import type { SettingsOptions } from "@parts/settings/settings";
import type { SectionOptions } from "@shared/section";

import type { AppPropertiesOptions } from "./app-properties";
import type { CustomPropertyOptions } from "./custom-properties";
import type { DocumentBackgroundOptions } from "./document";
import type { EndnoteOptions, EndnoteSeparator } from "./endnotes/descriptor";
import type { FootnoteOptions, FootnoteSeparator } from "./footnotes/descriptor";
import type { NumberingOptions } from "./numbering";
import type { StylesOptions } from "./styles";
import type { WebSettingsOptions } from "./web-settings";

/**
 * Options for configuring document properties.
 *
 * All settings.xml content lives under the single `settings` entry
 * (SettingsOptions), mirroring the OOXML part structure.
 *
 * @property sections - Document section configurations
 * @property title - Document title
 * @property subject - Document subject
 * @property creator - Document creator/author
 * @property keywords - Document keywords for searchability
 * @property description - Document description
 * @property lastModifiedBy - User who last modified the document
 * @property revision - Revision number
 * @property lastPrinted - Last printed timestamp (W3CDTF), round-tripped from cp:lastPrinted
 * @property styles - Document styles configuration
 * @property numbering - Numbering configuration
 * @property comments - Document comments (word/comments.xml)
 * @property people - Comment authors (word/people.xml)
 * @property commentsExtended - Comment metadata (word/commentsExtended.xml)
 * @property bibliography - Document bibliography sources
 * @property footnotes - Document footnotes
 * @property background - Document background settings
 * @property customProperties - Custom document properties
 * @property fonts - Font configurations
 * @property settings - Document settings (word/settings.xml)
 */
export interface DocumentOptions extends CorePropertiesOptions {
  sections: SectionOptions[];
  /**
   * Encrypted source (OLE2/CFB container). Round-trip only: the original
   * bytes are re-emitted unchanged, every other field stays empty (`sections:
   * []`); mixing real content is rejected. Do not hand-author.
   */
  encrypted?: EncryptedContainerOptions;
  styles?: StylesOptions;
  numbering?: NumberingOptions;
  comments?: CommentOptions[];
  /**
   * Comment authors (word/people.xml) — Word's display-name directory. Unlike
   * commentsExtended/commentsIds, a w15:person carries no paraId or durable
   * id, so author string equality is the only pairing Word's schema offers;
   * omit for documents whose comments carry no author registry.
   */
  people?: CommentPersonOptions[];
  /**
   * Comment metadata (word/commentsExtended.xml) — resolved state and reply
   * threading, keyed by the w14:paraId of each comment's first paragraph.
   */
  commentsExtended?: CommentExtendedOptions[];
  /**
   * Durable comment ids (word/commentsIds.xml, w16cid) — paraId → durableId.
   * The durable id is the stable key for a comment across edits; Word 2016+
   * writes it alongside comments.xml.
   */
  commentsIds?: CommentIdOptions[];
  /**
   * Extensible comment metadata (word/commentsExtensible.xml, w16cex) —
   * per-comment UTC time and extension list, keyed by the durable id from
   * `commentsIds`.
   */
  commentsExtensible?: CommentExtensibleOptions[];
  bibliography?: BibliographyOptions;
  /** User footnotes (word/footnotes.xml). `id` auto-assigns 1, 2, … when omitted. */
  footnotes?: FootnoteOptions[];
  /**
   * Separator footnotes — id + content round-tripped verbatim from the source
   * so the generated id stays consistent with settings.footnoteProperties,
   * references it. Omit for freshly generated documents (defaults apply).
   */
  footnoteSeparators?: {
    /** null = the parsed source carried no such note (emit nothing). */
    separator?: FootnoteSeparator | null;
    continuationSeparator?: FootnoteSeparator | null;
    /** w:type="continuationNotice" — round-tripped verbatim, no fresh default. */
    continuationNotice?: FootnoteSeparator;
  };
  /** User endnotes (word/endnotes.xml). `id` auto-assigns 1, 2, … when omitted. */
  endnotes?: EndnoteOptions[];
  /** Separator endnotes — round-tripped verbatim from the source. */
  endnoteSeparators?: {
    /** null = the parsed source carried no such note (emit nothing). */
    separator?: EndnoteSeparator | null;
    continuationSeparator?: EndnoteSeparator | null;
    /** w:type="continuationNotice" — round-tripped verbatim, no fresh default. */
    continuationNotice?: EndnoteSeparator;
  };
  background?: DocumentBackgroundOptions;
  customProperties?: CustomPropertyOptions[];
  fonts?: EmbeddedFontOptions[];
  /** Document conformance class (w:document/`@w:conformance`). */
  conformance?: "strict" | "transitional";
  /** Glossary document — building blocks (Quick Parts) */
  glossary?: GlossaryDocumentOptions;
  /** Document settings (word/settings.xml). */
  settings?: SettingsOptions;
  /** Web settings for browser rendering (word/webSettings.xml) */
  webSettings?: WebSettingsOptions;
  /** Content types from [Content_Types].xml (parse path only) */
  contentTypes?: ContentTypesInput;
  /**
   * Source parts generate() does not rebuild (word/theme/*, customXml/*,
   * unknown extension parts), carried verbatim with bytes and content-type
   * intact; compiler-rebuilt paths win over the passthrough copy. Round-trip
   * only — do not hand-author.
   */
  rawParts?: { path: string; data: DataType; contentType?: string }[];
  /**
   * Source .rels relationships from rebuilt parts pointing at rawParts (e.g.
   * document.xml → theme/customXml); re-emitted verbatim with a fresh rId.
   * Round-trip only — do not hand-author.
   */
  passthroughRelationships?: {
    source: string;
    relationshipType: string;
    target: string;
    rId: string;
    targetMode?: "External";
  }[];
  /**
   * Digital-signature parts detected at parse time. Round-trip only — do not
   * hand-author. The `_xmlsignatures/` bytes ride through `rawParts`
   * verbatim; regeneration changes signed content, so the saved package is no
   * longer validly signed (office-open never signs or verifies).
   */
  digitalSignatures?: DigitalSignatureInfo;
  /** Extended properties (docProps/app.xml) */
  appProperties?: AppPropertiesOptions;
}

// ── Descriptor ──

import type { CorePropertiesOptions } from "@office-open/core";
import { buildCorePropertiesXmlString, parseCorePropsElement } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";

export const corePropertiesDesc: CustomDescriptor<CorePropertiesOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildCorePropertiesXmlString(opts);
  },

  parse(el, _ctx) {
    return parseCorePropsElement(el);
  },
};
