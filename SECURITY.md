# Security Policy

## Supported Versions

Use this section to tell people about which versions of your project are
currently being supported with security updates.

| Version | Supported          |
| ------- | ------------------ |
| 0.1.x   | :white_check_mark: |

## Reporting a Vulnerability

We encourage responsible disclosure of security vulnerabilities. If you believe you have found a security vulnerability in this project, please report it via the [Security Tab](https://github.com/DemoMacro/office-open/security/advisories)

Please include the following information in your report:

- A description of the vulnerability
- Steps to reproduce the vulnerability
- Impact of the vulnerability

We will investigate all reported vulnerabilities and take appropriate action.

We appreciate your help in keeping this project secure.

## Signed, encrypted, and macro-enabled packages

office-open is a **preservation** layer for the content-protection features of
OOXML packages. It detects them, carries their parts through parse → generate,
and refuses or warns precisely. It is **not** a signing, signature-verification,
password-decryption, or macro-runtime layer.

### Support matrix

| Feature                                                                                    | Detection                                                                                                                                                                | Byte-exact preservation                                                                                                                                                                                     | Edit / regenerate                                                                                                                                                                                                                     | Decrypt / execute / verify  |
| ------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------- |
| ECMA-376 **agile** encryption (4.4, AES-CBC/SHA-512)                                       | `detectEncryptedContainer` → `flavor: "agile"`, `ooxmlPackage: true`; `parseDocument` returns `sections: []`, `encrypted.data`/`encrypted.status`                        | Yes — the whole CFB is re-emitted unchanged by `generateDocument*` (encrypted passthrough is exclusive of real content)                                                                                     | Unsupported in office-open; the original bytes round-trip. `@akademi-ink/docx` decrypts with a password (WebCrypto) and re-encrypts on save                                                                                           | No decrypt, no execute      |
| **Standard/CryptoAPI** encryption (3.2 / 4.2)                                              | `detectEncryptedContainer` → `flavor: "standard"`, message says “not supported here”                                                                                     | Yes — verbatim CFB passthrough                                                                                                                                                                              | Unsupported (no decrypt path)                                                                                                                                                                                                         | No                          |
| **Other CFB** containers (legacy `.doc`, arbitrary OLE)                                    | `detectEncryptedContainer` → `ooxmlPackage: false`, `flavor: "unknown"`; `parseArchive` throws `EncryptedPackageError` with that status instead of a generic ZIP error   | `parseDocument` still treats any CFB as encrypted passthrough (the editor routes `.doc` through its bounded MS-DOC reader)                                                                                  | Unsupported through the OOXML path                                                                                                                                                                                                    | No                          |
| **Digital signatures** (`_xmlsignatures/`)                                                 | `detectDigitalSignatures` → `origin`, `signatures`, `parts`, and package-root/origin `relationships`; surfaced on `DocumentOptions.digitalSignatures`                    | Yes — every `_xmlsignatures/` part (including `.sigs`, `.xml`, and companion `.rels`) is carried verbatim; the obsolete-namespace normalization that other passthrough XML receives is deliberately skipped | The package regenerates, so signatures no longer validate; `compileDocument` warns once via `digitalSignatureInvalidationWarning` / `warnDigitalSignatureInvalidation`                                                                | No signing, no verification |
| **Macros** (`.docm` / `.dotm`, `vbaProject.bin`, `vbaData.xml`, `vbaProjectSignature.bin`) | `isMacroEnabledPackage` (main-document content type) / `isMacroPartPath` (part basename); the editor also probes the ZIP directory (`_xmlsignatures/`, `vbaProject.bin`) | Yes on a macro-enabled save — VBA parts are passthrough bytes, never parsed or executed                                                                                                                     | A macro-free output (fresh compile, or a variant flip to `docx`/`dotx`) drops the VBA parts and any part-specific declarations (a source `Default`/`Override` may remain as an unused declaration); a macro-enabled output keeps them | No execution, no scanning   |

### Encryption matrix notes

- office-open never decrypts. `parseDocument` reports
  `encrypted: { data, status }`; `generateDocument*` re-emits `data` verbatim and
  rejects mixing it with real content (`assertEncryptedExclusive`).
- Flavor detection is intentionally best-effort and file-agnostic: it reads the
  CFB stream names (`EncryptionInfo`, `EncryptedPackage`, stored as UTF-16LE
  directory entries) and the EncryptionInfo version header (agile `4.4`,
  Standard `3.2`/`4.2`). A malformed or non-contiguous envelope reports
  `unknown`, never “decryptable”.
- `parseArchive` rejects a CFB container with `EncryptedPackageError` carrying
  that status — never the generic “not a zip” error. `@akademi-ink/docx` maps the
  same states to `EncryptedDocumentError` (prompt for a password) and
  `EncryptionUnsupportedError` (Standard/CryptoAPI; no retry can succeed).
- The tests pin both synthetic wrappers (`packages/core/src/opc/encryption.spec.ts`)
  and the docx parse status (`packages/docx/src/parse/encrypted-container.spec.ts`).

### Digital-signature policy

- Signature parts and their relationships are detected and preserved; the
  package is **not** verified and office-open cannot tell whether the source
  signature was valid.
- **Any regeneration invalidates the signatures.** Even when the generated
  document bytes happen to match, `[Content_Types].xml` and `_rels/.rels` are
  rebuilt and part serialization can differ, so the signed digests no longer
  match. `compileDocument` emits a one-time console warning when the parsed
  options carry `digitalSignatures`; callers that surface UI should show
  `digitalSignatureInvalidationWarning(info)` before saving.
- Preserved signature parts are exact bytes: `_xmlsignatures/` is exempt from
  the obsolete-namespace normalization applied to other passthrough XML. This
  keeps the source package re-signable/diagnosable, not valid.

### Macro policy

- Macros are never executed, compiled, or inspected. VBA parts are opaque
  bytes.
- A macro-enabled output (`docm`/`dotm` main content type) preserves
  `vbaProject.bin`, `vbaData.xml`, and `vbaProjectSignature.bin` byte-for-byte,
  along with the source content-type declarations.
- A macro-free output (`.docx`/`.dotx`, including a Save As from a `.docm`)
  strips those parts; a macro-free package must not carry an orphaned VBA
  project. The macro-enabled signal is the main-document content type
  (`…macroEnabled.main+xml`); office-open has no separate variant parameter.
- The editor warns on open for `.docm`/`.dotm` and whenever a VBA part is
  present; the parts are preserved for the source-faithful save.

### Residual risks

1. **No signature verification.** A preserved-but-invalid signature can still
   be displayed as “signed” by other tools; only the source package is known
   intact, not its trust chain.
2. **Digital signatures do not survive any edit.** The warning is advisory;
   callers must not present a regenerated file as signed.
3. **CFB flavor detection is heuristic.** office-open does not parse the
   compound-file directory; unusual/segmented envelopes can report `unknown`.
   Treat `flavor` as a status hint, not a security decision.
4. **Legacy `.doc` shares the CFB path.** `isEncryptedContainer` matches any
   OLE2/CFB file; the precise status (`ooxmlPackage: false`) names it, but
   `parseDocument` still takes the encrypted-passthrough branch. The editor must
   route `.doc` to its bounded reader before the OOXML path.
5. **Macro detection is name-based.** Only the VBA project basenames are
   stripped from a macro-free save; ActiveX controls, OLE embeddings, and other
   executable passthrough parts are preserved as opaque bytes.
6. **No antivirus / content scanning.** Preserved bytes are not sanitized;
   downstream hosts own trust, sandboxing, and policy.
7. **Encrypted packages are opaque.** No content validation, repair, or
   re-encryption happens in office-open; the bytes round-trip only. The traits
   inside an encrypted package (signatures, macros) cannot be read without
   decrypting, so a source-probe reports only `encrypted` for it.
8. **ZIP-level differences are expected.** Rebuilt XML parts, content types,
   and relationship ids can differ from the source; only the explicitly
   preserved passthrough parts are guaranteed byte-identical.
