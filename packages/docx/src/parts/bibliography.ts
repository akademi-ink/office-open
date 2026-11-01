/**
 * Bibliography module for WordprocessingML documents.
 *
 * Provides types for bibliography sources.
 * XML generation is handled by the descriptor pipeline (bibliographyDesc).
 *
 * Reference: ISO/IEC 29500-4, shared-bibliography.xsd, CT_Sources, CT_SourceType
 *
 * @module
 */

/**
 * A person in a bibliography name list (CT_PersonType).
 */
export interface PersonOptions {
  last?: string;
  first?: string;
  middle?: string;
}

/**
 * A corporate (organization) author entry (b:Corporate) — only valid in the
 * author and performer roles, whose XSD type is CT_NameOrCorporateType.
 */
export interface CorporateOptions {
  corporate: string;
}

/** One entry of a role list: a person or (author/performer roles) a corporation. */
export type AuthorEntry = PersonOptions | CorporateOptions;

/**
 * Author container (CT_AuthorType) — one array per named role; each entry
 * serializes to one `b:<Role>` element. All roles are optional; include only
 * the ones relevant to the source type.
 */
export interface AuthorOptions {
  /** b:Author role */
  authors?: AuthorEntry[];
  /** b:Artist role */
  artists?: PersonOptions[];
  /** b:BookAuthor role */
  bookAuthors?: PersonOptions[];
  /** b:Compiler role */
  compilers?: PersonOptions[];
  /** b:Composer role */
  composers?: PersonOptions[];
  /** b:Conductor role */
  conductors?: PersonOptions[];
  /** b:Counsel role */
  counsel?: PersonOptions[];
  /** b:Director role */
  directors?: PersonOptions[];
  /** b:Editor role */
  editors?: PersonOptions[];
  /** b:Interviewee role */
  interviewees?: PersonOptions[];
  /** b:Interviewer role */
  interviewers?: PersonOptions[];
  /** b:Inventor role */
  inventors?: PersonOptions[];
  /** b:Performer role */
  performers?: AuthorEntry[];
  /** b:ProducerName role */
  producers?: PersonOptions[];
  /** b:Translator role */
  translators?: PersonOptions[];
  /** b:Writer role */
  writers?: PersonOptions[];
}

/**
 * Options for a single bibliography source entry.
 *
 * Maps to CT_SourceType in the bibliography XSD schema (choice of 48 elements,
 * all optional). All fields are optional — include only the relevant ones for
 * each source type. `author` maps the structured CT_AuthorType role container.
 *
 * @property sourceType - Source type element (ST_SourceType: Book, JournalArticle, …)
 * @property type - Subcategory string (XSD Type element)
 * @property title - Title of the work
 * @property author - Authors by role (CT_AuthorType)
 * @property year - Publication year
 * @property month - Publication month
 * @property day - Publication day
 * @property bookTitle - Title of the book (for book sections, articles in collections)
 * @property journal - Journal name
 * @property volume - Volume number
 * @property issue - Issue number
 * @property pages - Page range
 * @property publisher - Publisher name
 * @property city - City of publication
 * @property url - URL for internet sources
 * @property edition - Edition number or description
 * @property institution - Institution (for theses, reports)
 */
export interface SourceTypeOptions {
  /** Source type element (b:SourceType, ST_SourceType). */
  sourceType?:
    | "ArticleInAPeriodical"
    | "Book"
    | "BookSection"
    | "JournalArticle"
    | "ConferenceProceedings"
    | "Report"
    | "SoundRecording"
    | "Performance"
    | "Art"
    | "DocumentFromInternetSite"
    | "InternetSite"
    | "Film"
    | "Interview"
    | "Patent"
    | "ElectronicSource"
    | "Case"
    | "Misc";
  /** Subcategory string (XSD Type element). */
  type?: string;
  title?: string;
  author?: AuthorOptions;
  year?: string;
  month?: string;
  day?: string;
  bookTitle?: string;
  journal?: string;
  volume?: string;
  issue?: string;
  pages?: string;
  publisher?: string;
  city?: string;
  url?: string;
  edition?: string;
  institution?: string;
  abbreviatedCaseNumber?: string;
  albumTitle?: string;
  broadcaster?: string;
  broadcastTitle?: string;
  caseNumber?: string;
  chapterNumber?: string;
  comments?: string;
  conferenceName?: string;
  countryRegion?: string;
  court?: string;
  dayAccessed?: string;
  department?: string;
  distributor?: string;
  guid?: string;
  internetSiteTitle?: string;
  lcid?: string;
  medium?: string;
  monthAccessed?: string;
  numberVolumes?: string;
  patentNumber?: string;
  periodicalTitle?: string;
  productionCompany?: string;
  publicationTitle?: string;
  recordingNumber?: string;
  refOrder?: string;
  reporter?: string;
  shortTitle?: string;
  standardNumber?: string;
  stateProvince?: string;
  station?: string;
  tag?: string;
  theater?: string;
  thesisType?: string;
  version?: string;
  yearAccessed?: string;
}

/**
 * Options for creating a bibliography container.
 *
 * @property sources - Array of bibliography source entries
 * @property styleName - Bibliography style name (e.g., "APA", "Chicago", "IEEE")
 */
export interface BibliographyOptions {
  sources: SourceTypeOptions[];
  /** Word's display name for the selected style (b:Sources/@StyleName). */
  styleName?: string;
  /** Word's XSLT style path (b:Sources/@SelectedStyle), e.g. "\\APA.XSL". */
  selectedStyle?: string;
  /**
   * Sources schema version (b:Sources/@Version), e.g. "6". Word writes it on
   * save; preserved only when the source part carries it.
   */
  version?: string;
  /** Word's style URI (b:Sources/@URI), preserved when present. */
  uri?: string;
}

// ── Descriptor ──

import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml, textOf } from "@office-open/xml";

/**
 * Namespaces a bibliography `<Sources>` part may bind. 2006 is the ECMA-376
 * namespace Word writes today; 2004 is the pre-ECMA dialect (Word 2003/2004
 * source lists) still present in converted documents.
 */
const BIBLIOGRAPHY_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
  "http://schemas.microsoft.com/office/word/2004/10/bibliography",
]);

interface QualifiedName {
  prefix: string;
  local: string;
}

/** Split `prefix:local`; the parser keeps names as written (canonicalized). */
function qualify(name: string | undefined): QualifiedName | undefined {
  if (name === undefined || name === "") return undefined;
  const colon = name.indexOf(":");
  return colon < 0
    ? { prefix: "", local: name }
    : { prefix: name.slice(0, colon), local: name.slice(colon + 1) };
}

/**
 * URI the element binds to its own prefix when it declares one, else
 * `undefined` (the binding is inherited from an ancestor — Word declares the
 * bibliography namespace once, on the part root).
 */
function declaredNamespace(el: XmlElement, name: QualifiedName): string | undefined {
  const value = el.attributes?.[name.prefix === "" ? "xmlns" : `xmlns:${name.prefix}`];
  return value === undefined ? undefined : String(value);
}

/**
 * Every prefix the part root binds to a bibliography namespace ("" is the
 * default namespace). Children inherit these bindings; a `b:` literal would
 * miss sources bound to a default or alternate prefix.
 */
function bibliographyPrefixes(root: XmlElement): ReadonlySet<string> {
  const prefixes = new Set<string>();
  for (const [key, value] of Object.entries(root.attributes ?? {})) {
    if (key !== "xmlns" && !key.startsWith("xmlns:")) continue;
    if (typeof value !== "string" || !BIBLIOGRAPHY_NAMESPACES.has(value)) continue;
    prefixes.add(key === "xmlns" ? "" : key.slice(6));
  }
  // Canonical `b:` root with no declaration to inspect (legacy/hand-built).
  if (prefixes.size === 0 && qualify(root.name)?.prefix === "b") prefixes.add("b");
  return prefixes;
}

/**
 * Match a direct child by local name in the part root's namespace scope. A
 * declaration on the child itself wins; otherwise it must use a prefix the
 * root bound to a bibliography namespace.
 */
function matchesBibliographyChild(
  el: XmlElement,
  local: string,
  scope: ReadonlySet<string>,
): boolean {
  const name = qualify(el.name);
  if (name?.local !== local) return false;
  const declared = declaredNamespace(el, name);
  if (declared !== undefined) return BIBLIOGRAPHY_NAMESPACES.has(declared);
  return scope.has(name.prefix);
}

function findBibliographyChild(
  parent: XmlElement | undefined,
  local: string,
  scope: ReadonlySet<string>,
): XmlElement | undefined {
  for (const child of parent?.elements ?? []) {
    if (matchesBibliographyChild(child, local, scope)) return child;
  }
  return undefined;
}

/**
 * True when `el` is the root of a bibliography part: local name `Sources` in
 * one of the bibliography namespaces. Prefix-agnostic, so a default-namespace
 * `<Sources xmlns="…bibliography">` is recognized alongside `b:Sources`.
 */
export function isBibliographySources(el: XmlElement): boolean {
  const name = qualify(el.name);
  if (name?.local !== "Sources") return false;
  const declared = declaredNamespace(el, name);
  if (declared !== undefined) return BIBLIOGRAPHY_NAMESPACES.has(declared);
  // Canonical/legacy shape with no declaration to inspect: the XML parser
  // normalizes 2006-namespace roots to `b:`.
  return name.prefix === "b";
}

const SOURCE_FIELDS: readonly (readonly [string, keyof SourceTypeOptions])[] = [
  ["SourceType", "sourceType"],
  ["Type", "type"],
  ["Title", "title"],
  ["Author", "author"],
  ["Year", "year"],
  ["Month", "month"],
  ["Day", "day"],
  ["BookTitle", "bookTitle"],
  ["JournalName", "journal"],
  ["Volume", "volume"],
  ["Issue", "issue"],
  ["Pages", "pages"],
  ["Publisher", "publisher"],
  ["City", "city"],
  ["URL", "url"],
  ["Edition", "edition"],
  ["Institution", "institution"],
  ["AbbreviatedCaseNumber", "abbreviatedCaseNumber"],
  ["AlbumTitle", "albumTitle"],
  ["Broadcaster", "broadcaster"],
  ["BroadcastTitle", "broadcastTitle"],
  ["CaseNumber", "caseNumber"],
  ["ChapterNumber", "chapterNumber"],
  ["Comments", "comments"],
  ["ConferenceName", "conferenceName"],
  ["CountryRegion", "countryRegion"],
  ["Court", "court"],
  ["DayAccessed", "dayAccessed"],
  ["Department", "department"],
  ["Distributor", "distributor"],
  ["Guid", "guid"],
  ["InternetSiteTitle", "internetSiteTitle"],
  ["LCID", "lcid"],
  ["Medium", "medium"],
  ["MonthAccessed", "monthAccessed"],
  ["NumberVolumes", "numberVolumes"],
  ["PatentNumber", "patentNumber"],
  ["PeriodicalTitle", "periodicalTitle"],
  ["ProductionCompany", "productionCompany"],
  ["PublicationTitle", "publicationTitle"],
  ["RecordingNumber", "recordingNumber"],
  ["RefOrder", "refOrder"],
  ["Reporter", "reporter"],
  ["ShortTitle", "shortTitle"],
  ["StandardNumber", "standardNumber"],
  ["StateProvince", "stateProvince"],
  ["Station", "station"],
  ["Tag", "tag"],
  ["Theater", "theater"],
  ["ThesisType", "thesisType"],

  ["Version", "version"],
  ["YearAccessed", "yearAccessed"],
] as const;

// Role element tag → AuthorOptions key. Author/Performer are CT_NameOrCorporateType;
// every other role is CT_NameType (person list only).
const AUTHOR_ROLES: readonly (readonly [string, keyof AuthorOptions])[] = [
  ["Author", "authors"],
  ["Artist", "artists"],
  ["BookAuthor", "bookAuthors"],
  ["Compiler", "compilers"],
  ["Composer", "composers"],
  ["Conductor", "conductors"],
  ["Counsel", "counsel"],
  ["Director", "directors"],
  ["Editor", "editors"],
  ["Interviewee", "interviewees"],
  ["Interviewer", "interviewers"],
  ["Inventor", "inventors"],
  ["Performer", "performers"],
  ["ProducerName", "producers"],
  ["Translator", "translators"],
  ["Writer", "writers"],
] as const;

const ROLE_LOCAL_TO_KEY = new Map<string, keyof AuthorOptions>(AUTHOR_ROLES);

/** Roles whose XSD type is CT_NameOrCorporateType (accept b:Corporate). */
const CORPORATE_ROLES = new Set(["Author", "Performer"]);

function stringifyAuthorEntry(entry: AuthorEntry): string {
  if ("corporate" in entry) {
    return `<b:Corporate>${escapeXml(entry.corporate)}</b:Corporate>`;
  }
  const parts: string[] = [];
  if (entry.last !== undefined) parts.push(`<b:Last>${escapeXml(entry.last)}</b:Last>`);
  if (entry.first !== undefined) parts.push(`<b:First>${escapeXml(entry.first)}</b:First>`);
  if (entry.middle !== undefined) parts.push(`<b:Middle>${escapeXml(entry.middle)}</b:Middle>`);
  return `<b:NameList><b:Person>${parts.join("")}</b:Person></b:NameList>`;
}

function stringifyAuthorElement(author: AuthorOptions): string {
  const roleParts: string[] = [];
  for (const [tagName, key] of AUTHOR_ROLES) {
    const entries = author[key];
    if (!entries) continue;
    for (const entry of entries) {
      roleParts.push(`<b:${tagName}>${stringifyAuthorEntry(entry)}</b:${tagName}>`);
    }
  }
  return roleParts.length > 0 ? `<b:Author>${roleParts.join("")}</b:Author>` : "";
}

function readPerson(el: XmlElement, scope: ReadonlySet<string>): PersonOptions {
  const person: PersonOptions = {};
  const last = findBibliographyChild(el, "Last", scope);
  if (last) {
    const val = textOf(last);
    if (val) person.last = val;
  }
  const first = findBibliographyChild(el, "First", scope);
  if (first) {
    const val = textOf(first);
    if (val) person.first = val;
  }
  const middle = findBibliographyChild(el, "Middle", scope);
  if (middle) {
    const val = textOf(middle);
    if (val) person.middle = val;
  }
  return person;
}

function parseAuthorElement(el: XmlElement, scope: ReadonlySet<string>): AuthorOptions | undefined {
  const author: { [K in keyof AuthorOptions]?: AuthorEntry[] } = {};
  for (const roleEl of el.elements ?? []) {
    const roleName = qualify(roleEl.name);
    if (!roleName || !matchesBibliographyChild(roleEl, roleName.local, scope)) continue;
    const key = ROLE_LOCAL_TO_KEY.get(roleName.local);
    if (!key) continue;

    const entries: AuthorEntry[] = [];
    const nameList = findBibliographyChild(roleEl, "NameList", scope);
    if (nameList) {
      for (const personEl of nameList.elements ?? []) {
        if (!matchesBibliographyChild(personEl, "Person", scope)) continue;
        entries.push(readPerson(personEl, scope));
      }
    } else if (CORPORATE_ROLES.has(roleName.local)) {
      const corporate = findBibliographyChild(roleEl, "Corporate", scope);
      const name = corporate ? textOf(corporate) : undefined;
      if (name) entries.push({ corporate: name });
    }
    if (entries.length === 0) continue;
    author[key] = [...(author[key] ?? []), ...entries];
  }
  return Object.keys(author).length > 0 ? (author as AuthorOptions) : undefined;
}

export const bibliographyDesc: CustomDescriptor<BibliographyOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const attrParts: string[] = [
      'xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"',
    ];
    if (opts.selectedStyle !== undefined) {
      attrParts.push(`SelectedStyle="${escapeXml(opts.selectedStyle)}"`);
    }
    if (opts.styleName !== undefined) {
      attrParts.push(`StyleName="${escapeXml(opts.styleName)}"`);
    }
    if (opts.version !== undefined) {
      attrParts.push(`Version="${escapeXml(opts.version)}"`);
    }
    if (opts.uri !== undefined) {
      attrParts.push(`URI="${escapeXml(opts.uri)}"`);
    }

    const parts: string[] = [`<b:Sources ${attrParts.join(" ")}>`];

    for (const source of opts.sources) {
      const sourceParts: string[] = [];
      for (const [tagName, key] of SOURCE_FIELDS) {
        if (key === "author") {
          if (source.author !== undefined) {
            sourceParts.push(stringifyAuthorElement(source.author));
          }
          continue;
        }
        const value = source[key];
        if (value !== undefined) {
          sourceParts.push(`<b:${tagName}>${escapeXml(value)}</b:${tagName}>`);
        }
      }
      parts.push(`<b:Source>${sourceParts.join("")}</b:Source>`);
    }

    parts.push("</b:Sources>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const opts: Partial<BibliographyOptions> = {};
    const scope = bibliographyPrefixes(el);

    // Style attributes (SelectedStyle is what Word actually uses to pick the XSLT)
    const selectedStyle = el.attributes?.["SelectedStyle"];
    if (selectedStyle) opts.selectedStyle = selectedStyle as string;
    const styleName = el.attributes?.["StyleName"];
    if (styleName) opts.styleName = styleName as string;
    // nativeTypeAttributes coerces a numeric version ("6") to a number; the
    // Options field is a string so the value round-trips verbatim.
    const version = el.attributes?.["Version"];
    if (version !== undefined && version !== "") opts.version = String(version);
    const uri = el.attributes?.["URI"];
    if (uri) opts.uri = uri as string;

    // Parse Source children by local name in the bibliography namespace —
    // the source may bind it to a default prefix or none of Word's `b:`.
    const sources: SourceTypeOptions[] = [];
    for (const child of el.elements ?? []) {
      if (!matchesBibliographyChild(child, "Source", scope)) continue;
      const source: SourceTypeOptions = {};
      for (const [tagName, key] of SOURCE_FIELDS) {
        const xmlChild = findBibliographyChild(child, tagName, scope);
        if (!xmlChild) continue;
        if (key === "author") {
          const parsed = parseAuthorElement(xmlChild, scope);
          if (parsed) source.author = parsed;
          continue;
        }
        const val = textOf(xmlChild);
        if (!val) continue;
        if (key === "sourceType") {
          source.sourceType = val as SourceTypeOptions["sourceType"];
        } else {
          source[key] = val;
        }
      }
      sources.push(source);
    }
    opts.sources = sources;

    return opts as BibliographyOptions;
  },
};
