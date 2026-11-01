/**
 * OPC (Open Packaging Convention) consistency validator.
 *
 * Pure function operating on an already-unzipped part map. Detects the three
 * classes of cross-part breakage that single-part XSD validation cannot see:
 *
 *   - relationship integrity  — dangling targets, undeclared r:id references,
 *                                duplicate rIds (O3/O4/O7)
 *   - content-type integrity  — parts with no Override/Default, stale
 *                                Overrides pointing at absent parts (O5/O6)
 *   - structural integrity    — missing `always` parts, undeclared orphans
 *                                (O1/O2)
 *
 * O3/O4/O5/O6/O7 are derived directly from the ZIP + `.rels` + `[Content_Types].xml`
 * and need no registry. O1/O2 consult {@link PackagePartRegistry}.
 *
 * Reference: ECMA-376 Part 2 (OPC).
 *
 * @module
 */
import { parse } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { PackagePartRegistry } from "./part-registry";

export type OpcSeverity = "error" | "warn";

export type OpcCode = "O1" | "O2" | "O3" | "O4" | "O5" | "O6" | "O7";

export interface OpcIssue {
  code: OpcCode;
  severity: OpcSeverity;
  /** Path of the offending part, `.rels`, or Override `PartName`. */
  part: string;
  message: string;
}

interface RelationshipEntry {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

interface ParsedRels {
  relsPath: string;
  entries: readonly RelationshipEntry[];
}

interface ContentTypesDoc {
  defaults: ReadonlyMap<string, string>;
  overrides: ReadonlyMap<string, string>;
}

// ── XML parsing helpers ─────────────────────────────────────────────────────

function rootElement(xml: string): Element | undefined {
  try {
    const wrapper = parse(xml) as Element;
    return wrapper.elements?.find((e) => e.type === "element");
  } catch {
    return undefined;
  }
}

function parseContentTypes(entries: ReadonlyMap<string, string>): ContentTypesDoc {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  const xml = entries.get("[Content_Types].xml");
  const root = xml ? rootElement(xml) : undefined;
  if (!root) return { defaults, overrides };
  for (const child of root.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "Default") {
      const ext = String(child.attributes?.Extension ?? "").toLowerCase();
      const ct = String(child.attributes?.ContentType ?? "");
      if (ext && ct) defaults.set(ext, ct);
    } else if (child.name === "Override") {
      const partName = normalizePartName(String(child.attributes?.PartName ?? ""));
      const ct = String(child.attributes?.ContentType ?? "");
      if (partName && ct) overrides.set(partName, ct);
    }
  }
  return { defaults, overrides };
}

function parseAllRels(entries: ReadonlyMap<string, string>): ParsedRels[] {
  const out: ParsedRels[] = [];
  for (const [relsPath, xml] of entries) {
    if (!relsPath.endsWith(".rels")) continue;
    const root = rootElement(xml);
    if (!root) continue;
    const list: RelationshipEntry[] = [];
    for (const child of root.elements ?? []) {
      if (child.type !== "element" || child.name !== "Relationship") continue;
      list.push({
        id: String(child.attributes?.Id ?? ""),
        type: String(child.attributes?.Type ?? ""),
        target: String(child.attributes?.Target ?? ""),
        targetMode: child.attributes?.TargetMode ? String(child.attributes?.TargetMode) : undefined,
      });
    }
    out.push({ relsPath, entries: list });
  }
  return out;
}

/** Strip the leading slash OPC uses on `PartName` to align with ZIP paths. */
function normalizePartName(partName: string): string {
  return partName.startsWith("/") ? partName.slice(1) : partName;
}

/**
 * Resolve a relationship `Target` (relative to its owning `.rels`) against the
 * ZIP root. Handles the `..` segments slides/masters use to reference siblings.
 */
function resolveRelTarget(relsPath: string, target: string): string {
  const slash = relsPath.lastIndexOf("/");
  const relsDir = slash === -1 ? "" : relsPath.slice(0, slash);
  // `<partDir>/_rels` → `<partDir>`; root-level `_rels` → ""
  const base =
    relsDir === "_rels" ? "" : relsDir.endsWith("/_rels") ? relsDir.slice(0, -6) : relsDir;
  const segments = (base ? base.split("/") : []).concat(target.split("/"));
  const resolved: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") resolved.pop();
    else resolved.push(seg);
  }
  return resolved.join("/");
}

/** Compile a registry path template (with `${i}`) into an anchored matcher. */
function pathMatcher(template: string): RegExp {
  const marked = template.replace(/\$\{i\}/g, "\u0000");
  const escaped = marked.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = escaped.replaceAll("\u0000", "(\\d+)");
  return new RegExp(`^${pattern}$`);
}

function extensionOf(path: string): string {
  const dot = path.lastIndexOf(".");
  return dot === -1 ? "" : path.slice(dot + 1).toLowerCase();
}

/**
 * Reference attributes in the officeDocument relationships namespace — every
 * one resolves against the owning part's `.rels`. Escaped text can never match
 * (a literal `"` in text is `&quot;`), so hits are always real attributes.
 */
const REL_REF_ATTR_RE = /\br:(?:id|embed|link|dm|lo|qs|cs|pict)="([^"]*)"/g;

/** `word/_rels/document.xml.rels` → `word/document.xml`; package rels → undefined. */
function partOfRels(relsPath: string): string | undefined {
  const slash = relsPath.lastIndexOf("/");
  const dir = slash === -1 ? "" : relsPath.slice(0, slash);
  const name = relsPath.slice(slash + 1).replace(/\.rels$/, "");
  if (dir === "_rels") return name === "" ? undefined : name;
  if (dir.endsWith("/_rels")) return `${dir.slice(0, -6)}/${name}`;
  return undefined;
}

// ── Validator ───────────────────────────────────────────────────────────────

/**
 * Validate OPC consistency of an unzipped package.
 *
 * @param entries  ZIP path → decoded XML/binary text. Binary parts (media,
 *                 fonts) are never parsed here — only their presence matters.
 * @param registry Declarative part expectations for the package format.
 * @returns Issues sorted by code then part. Empty array = consistent.
 */
export function validateOpcConsistency(
  entries: ReadonlyMap<string, string>,
  registry: PackagePartRegistry,
): OpcIssue[] {
  const issues: OpcIssue[] = [];
  const contentTypes = parseContentTypes(entries);
  const relsFiles = parseAllRels(entries);

  // O2 — missing `always` part. Conditional/repeated absences are legitimate.
  for (const def of registry.parts) {
    if (def.presence.kind !== "always" || def.path.includes("${i}")) continue;
    if (!entries.has(def.path)) {
      issues.push({
        code: "O2",
        severity: "error",
        part: def.path,
        message: `required ${registry.format} part is absent from the package`,
      });
    }
  }

  // O1 — orphan part: present but not declared and not whitelisted.
  const matchers = registry.parts.map((def) => pathMatcher(def.path));
  for (const path of entries.keys()) {
    if (path === "[Content_Types].xml") continue;
    if (path.endsWith(".rels")) continue;
    if (registry.orphanWhitelist.some((prefix) => path.startsWith(prefix))) continue;
    if (matchers.some((re) => re.test(path))) continue;
    issues.push({
      code: "O1",
      severity: "warn",
      part: path,
      message: `part not declared in the ${registry.format} registry (possible orphan)`,
    });
  }

  // O3 — dangling internal relationship target.
  for (const rf of relsFiles) {
    for (const rel of rf.entries) {
      if (rel.targetMode === "External") continue;
      if (/^https?:\/\//i.test(rel.target)) continue;
      const resolved = resolveRelTarget(rf.relsPath, rel.target);
      if (!resolved || entries.has(resolved)) continue;
      issues.push({
        code: "O3",
        severity: "error",
        part: rf.relsPath,
        message: `relationship ${rel.id || "(no Id)"} (${rel.type}) Target "${rel.target}" resolves to absent part "${resolved}"`,
      });
    }
  }

  // O4 — relationship reference integrity: every r:* reference attribute in a
  // part must be declared by that part's own .rels. An undeclared reference —
  // a hand-written rId, or an unreplaced `{key}` placeholder — makes Office
  // report the package as corrupt. External-mode relationships declare their
  // id like any other, so they satisfy references too.
  const relIdsByPart = new Map<string, Set<string>>();
  for (const rf of relsFiles) {
    const part = partOfRels(rf.relsPath);
    if (part === undefined) continue;
    const ids = relIdsByPart.get(part) ?? new Set<string>();
    for (const r of rf.entries) ids.add(r.id);
    relIdsByPart.set(part, ids);
  }
  for (const [path, xml] of entries) {
    if (path === "[Content_Types].xml" || path.endsWith(".rels") || !path.endsWith(".xml")) {
      continue;
    }
    const declared = relIdsByPart.get(path);
    const reported = new Set<string>();
    REL_REF_ATTR_RE.lastIndex = 0;
    for (let m = REL_REF_ATTR_RE.exec(xml); m; m = REL_REF_ATTR_RE.exec(xml)) {
      const ref = m[1]!;
      if (declared?.has(ref) || reported.has(ref)) continue;
      reported.add(ref);
      issues.push({
        code: "O4",
        severity: "error",
        part: path,
        message: `r:id/r:embed reference "${ref}" is not declared by the part's .rels`,
      });
    }
  }

  // O5 — part present but no Override covers it and no Default covers its
  // extension. Word rejects such parts as unreadable content.
  for (const path of entries.keys()) {
    if (path === "[Content_Types].xml" || path.endsWith(".rels")) continue;
    if (contentTypes.overrides.has(path)) continue;
    const ext = extensionOf(path);
    if (ext && contentTypes.defaults.has(ext)) continue;
    issues.push({
      code: "O5",
      severity: "error",
      part: path,
      message: `part has no [Content_Types] Override and no Default covers extension "${ext || "(none)"}"`,
    });
  }

  // O6 — stale Override: [Content_Types] declares a part that is not in the ZIP.
  for (const partName of contentTypes.overrides.keys()) {
    if (entries.has(partName)) continue;
    issues.push({
      code: "O6",
      severity: "error",
      part: `/${partName}`,
      message: `[Content_Types] Override references absent part "${partName}"`,
    });
  }

  // O7 — duplicate relationship Id within a single .rels part.
  for (const rf of relsFiles) {
    const counts = new Map<string, number>();
    for (const rel of rf.entries) {
      if (!rel.id) continue;
      counts.set(rel.id, (counts.get(rel.id) ?? 0) + 1);
    }
    for (const [id, count] of counts) {
      if (count <= 1) continue;
      issues.push({
        code: "O7",
        severity: "warn",
        part: rf.relsPath,
        message: `duplicate relationship Id "${id}" appears ${count} times`,
      });
    }
  }

  return issues.sort((a, b) =>
    a.code === b.code ? a.part.localeCompare(b.part) : a.code.localeCompare(b.code),
  );
}

export function summarizeOpcIssues(issues: readonly OpcIssue[]): {
  errors: number;
  warnings: number;
} {
  let errors = 0;
  let warnings = 0;
  for (const i of issues) {
    if (i.severity === "error") errors++;
    else warnings++;
  }
  return { errors, warnings };
}
