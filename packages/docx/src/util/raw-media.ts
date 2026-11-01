/**
 * Raw-XML media registration shared by every verbatim passthrough path.
 *
 * A `rawXml` child (block-level unknown section children, run-level unknown
 * markup, track-change fallback markup, unrecognized drawings, VML fallbacks,
 * the document background) carries `{fileName}` placeholders that
 * {@link replaceRelsWithPlaceholders} produced from the source `r:id` /
 * `r:embed` / `r:link` refs, plus the media bytes they referenced. Emitting the
 * XML alone would leave the placeholders unresolved (or, worse, keep the source
 * rIds), so every stringifier routes through here: the bytes are registered
 * with the package media store and the placeholder is remapped when dedup
 * reuses an earlier file name.
 *
 * Works on a local copy of the XML — the caller's options object is shared,
 * serializable data, and writing this generate run's numbering back into it
 * would leak into the next run's input.
 *
 * @module
 */
import { toUint8Array } from "@office-open/core";
import type { BackgroundRawMediaOptions } from "@parts/document/document-background/document-background";
import type { Media } from "@shared/media";
import type { MediaData } from "@shared/media/data";

/**
 * Register a raw XML child's referenced media and remap its `{fileName}`
 * placeholders. Returns the XML unchanged when it carries no media.
 */
export function registerRawMedia(
  xml: string,
  rawMedia: readonly BackgroundRawMediaOptions[] | undefined,
  media: Media<MediaData>,
): string {
  if (!rawMedia || rawMedia.length === 0) return xml;
  let out = xml;
  for (const m of rawMedia) {
    const data = toUint8Array(m.data);
    const entry = media.addMedia(
      data,
      m.type,
      (fileName) =>
        ({
          type: m.type,
          data,
          fileName,
          transformation: { emus: { x: 0, y: 0 }, pixels: { x: 0, y: 0 } },
        }) as MediaData,
      m.fileName,
    );
    // Dedup may reuse an earlier file name; remap the placeholder so the
    // compiler resolves it to the shared media relationship.
    if (entry.fileName !== m.fileName) {
      out = out.split(`{${m.fileName}}`).join(`{${entry.fileName}}`);
    }
  }
  return out;
}
