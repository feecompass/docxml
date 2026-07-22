# `@feecompass/docxml` — section formatting P0 + P1

**Status:** v1.1 — implementation spec (2026-07-22)  
**Package version target:** `1.2.0` (minor bump from `1.1.2`)  
**Audience:** Engineers working on `@feecompass/docxml` and Verosetta DOCX export  
**Consumer feature spec:** [Verosetta Fibery #141 — DOCX export section formatting](https://github.com/feecompass/docxml/blob/main/docs/specs/verosetta-consumer-notes.md) (see also `adk-sr/docs/feature-specs/docx-export-section-formatting-spec.md`)

---

## 1. Purpose

Verosetta patent DOCX export (Fibery feature #141) needs OOXML that Word applies correctly for:

- **FIG** — per-section page numbering restarted at 1 (`w:pgNumType`) plus header fields (`PAGE` / `SECTIONPAGES`)
- **LINENB** — per-section line numbering every 5th line (`w:lnNumType`)
- **OFF** — exports without Word compatibility-mode settings that break image positioning

This spec defines **library-level** changes only. It does **not** cover Verosetta org settings, `DocxmlVisitor` wiring, or custom document properties — those APIs already exist in docxml v1.1.2.

### Local dev loop (Verosetta monorepo)

During library development, Verosetta uses a pnpm workspace override:

```yaml
# adk-sr/pnpm-workspace.yaml
overrides:
    '@feecompass/docxml': 'file:///Users/pbuszka/projects/docxml/build/npm'
```

After each docxml change, rebuild the npm artifact and refresh in `adk-sr`:

```bash
# from docxml repo root
deno run -A tasks/build-npm.ts 1.2.0
```

Then reinstall/refresh in `adk-sr`. Remove the override when publishing `1.2.0` to the registry and bump the catalog pin.

---

## 2. Scope summary

| Priority | Area                            | OOXML element                  | New API surface                   |
| -------- | ------------------------------- | ------------------------------ | --------------------------------- |
| **P0**   | Section page numbering          | `w:pgNumType`                  | `SectionProperties.pageNumbering` |
| **P0**   | Section line numbering          | `w:lnNumType`                  | `SectionProperties.lineNumbering` |
| **P1**   | Document compatibility settings | `w:compat` / `w:compatSetting` | `SettingsI.compatibility`         |

### Explicitly out of scope (this release)

| Item                                                 | Rationale                                                                                                                                             |
| ---------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| **`w:docGrid`**                                      | Present incidentally in Artur `_LINENB` / `_FIG` reference samples; **not** a Verosetta product requirement                                           |
| **High-level `createPageNumberHeader()` helper**     | Verosetta composes headers with existing `Paragraph`, `FieldRange*`, `headers.add()` APIs                                                             |
| **`docProps/custom.xml` custom properties**          | Already supported via `docx.customProperties.add()`                                                                                                   |
| **`w:headerReference` / header parts**               | Already supported via `SectionProperties.headers` + `document.headers.add()`                                                                          |
| **DOTX / attached template**                         | Deferred in Verosetta v1                                                                                                                              |
| **Other `w:compat` children** (`w:spaceForUL`, etc.) | **v1:** not parsed or emitted; round-trip may drop them. Verosetta v1 only needs omit-on-write (`compatibility: null`). Opaque preservation deferred. |
| **Full `ST_NumberFormat` enum**                      | API exposes the subset Verosetta needs; other `w:fmt` values may appear as plain strings on read until a wider union is added deliberately            |

---

## 3. Reference OOXML (Verosetta samples)

Unzipped reference trees live in the Verosetta repo at `.scratches/feature/{fig,linenb,off}/ooxml/`.

### 3.1 Page numbering (`fig/`)

On the figures section `w:sectPr` (4th top-level break in the sample document):

```xml
<w:pgNumType w:start="1"/>
```

Header (`header4.xml`) uses complex fields (docxml already supports this):

```xml
<w:p>
  <w:pPr><w:jc w:val="right"/></w:pPr>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText>PAGE   \* MERGEFORMAT</w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>…</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:t>/</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText>SECTIONPAGES</w:instrText></w:r>
  …
</w:p>
```

**Lib responsibility:** emit `w:pgNumType` only. Header composition is consumer responsibility.

### 3.2 Line numbering (`linenb/`)

On the description section `w:sectPr` (2nd top-level break):

```xml
<w:lnNumType w:countBy="5"/>
```

The same section also has `w:docGrid w:linePitch="272"` — **ignore for implementation** (out of scope).

### 3.3 Compatibility mode (`off/`)

Artur's `_OFF` sample **still contains** a `w:compat` block with several `w:compatSetting` entries, including one with `w:name="compatibilityMode"` and `w:val="15"`. That is **not** a separate OOXML element — it is a `w:compatSetting` row (see §5.6 fixture). The sample is therefore a weak literal target for “compatibility off” as a product label.

**Observed behaviour that matches Verosetta intent:** `Docx.fromNothing()` today emits `word/settings.xml` **without** any `w:compat` element. P1 formalises and tests that default, and adds round-trip support for `w:compatSetting` when reading existing documents.

**OFF product semantics (resolved):** Verosetta **OFF** means **omit `w:compat` on export** (`compatibility: null`). UAT with Artur confirms image positioning; no requirement to match the `_OFF` sample's incidental compat block literally.

---

## 4. P0 — `SectionProperties` extensions

**Files:**

- `lib/properties/src/section-properties.ts` — types, `sectionPropertiesFromNode`, `sectionPropertiesToNode`
- `lib/properties/test/section-properties.test.ts` — XML ↔ object round-trip tests (existing `createXmlRoundRobinTest` pattern)

### 4.1 TypeScript API

Add to `SectionProperties`:

```typescript
/**
 * Maps to w:pgNumType (CT_PageNumber).
 * @see docs/ecma-376-5th/markup-reference/wml.xsd — CT_PageNumber
 */
pageNumbering?: null | {
  /** w:start — 1-based page number at section start. Verosetta uses 1 for FIG. */
  start?: number;
  /**
   * w:fmt — page number format.
   * Omit on write when decimal (schema default). On read, omit property when attribute absent.
   */
  format?:
    | 'decimal'
    | 'upperRoman'
    | 'lowerRoman'
    | 'upperLetter'
    | 'lowerLetter'
    | 'ordinal'
    | 'cardinalText';
  /** w:chapStyle — style whose heading level drives chapter numbering. */
  chapterStyle?: number;
  /** w:chapSep — separator between chapter number and page number. */
  chapterSeparator?: 'hyphen' | 'period' | 'colon' | 'emDash' | 'enDash';
};

/**
 * Maps to w:lnNumType (CT_LineNumber).
 * @see docs/ecma-376-5th/markup-reference/wml.xsd — CT_LineNumber
 */
lineNumbering?: null | {
  /** w:countBy — number interval between displayed line numbers. Verosetta uses 5. */
  countBy?: number;
  /** w:start — first line number (default 1 in OOXML). */
  start?: number;
  /** w:distance — margin between line number and text (twips). */
  distance?: Length;
  /** w:restart — when numbering restarts within the section. */
  restart?: 'newPage' | 'newSection' | 'continuous';
};
```

**Semantics:**

| Input                             | `toNode()` behaviour                                                                    |
| --------------------------------- | --------------------------------------------------------------------------------------- |
| Property **absent** or **`null`** | Do not emit the OOXML element                                                           |
| Property **`{}`**                 | Emit empty element (`<w:pgNumType/>` or `<w:lnNumType/>`; Word applies schema defaults) |
| Nested field **`undefined`**      | Omit that attribute on write                                                            |

**Read normalisation (`format`):** When `w:fmt` is absent, do **not** set `format: 'decimal'` — leave the property undefined. When `w:fmt` is present, map to the typed union; values outside the union round-trip as the raw string (same pattern as other loosely-typed OOXML enums elsewhere in the codebase).

### 4.2 `sectionPropertiesFromNode` (XPath map extension)

Extend the existing `evaluateXPathToMap` in `sectionPropertiesFromNode` (use `${QNS.w}` in implementation; illustrative XPath below):

```xpath
"pageNumbering": map {
  "start": ./w:pgNumType/@w:start/number(),
  "format": ./w:pgNumType/@w:fmt/string(),
  "chapterStyle": ./w:pgNumType/@w:chapStyle/number(),
  "chapterSeparator": ./w:pgNumType/@w:chapSep/string()
},
"lineNumbering": map {
  "countBy": ./w:lnNumType/@w:countBy/number(),
  "start": ./w:lnNumType/@w:start/number(),
  "distance": docxml:length(./w:lnNumType/@w:distance, 'twip'),
  "restart": ./w:lnNumType/@w:restart/string()
}
```

**Post-map normalisation** (TypeScript, after `evaluateXPathToMap`):

```typescript
function dropEmptySectionProp<T extends Record<string, unknown>>(
	obj: T | undefined,
): T | undefined {
	if (!obj) return undefined;
	const hasValue = Object.values(obj).some(
		(v) => v !== undefined && v !== null && v !== '',
	);
	return hasValue ? obj : undefined;
}

// After map:
props.pageNumbering = dropEmptySectionProp(props.pageNumbering);
props.lineNumbering = dropEmptySectionProp(props.lineNumbering);
if (props.pageNumbering === undefined) delete props.pageNumbering;
if (props.lineNumbering === undefined) delete props.lineNumbering;
```

Rules:

- Element **absent** in XML → parent key absent on the returned object.
- Element **present but empty** (`<w:pgNumType/>`) → parent key present as **`{}`**.
- `distance` uses existing `Length` type (`twip()` on read).

### 4.3 `sectionPropertiesToNode` (XQuery emission)

Insert **after** `w:pgMar` / **before** `w:titlePg`. Per ECMA `EG_SectPrContents`, emit **`w:lnNumType` before `w:pgNumType`** (Word tolerates other orders; this matches the schema):

```xquery
if (exists($lineNumbering)) then element w:lnNumType {
  if (exists($lineNumbering('countBy')))
    then attribute w:countBy { $lineNumbering('countBy') } else (),
  if (exists($lineNumbering('start')))
    then attribute w:start { $lineNumbering('start') } else (),
  if (exists($lineNumbering('distance')))
    then attribute w:distance { round($lineNumbering('distance')('twip')) } else (),
  if (exists($lineNumbering('restart')))
    then attribute w:restart { $lineNumbering('restart') } else ()
} else (),
if (exists($pageNumbering)) then element w:pgNumType {
  if (exists($pageNumbering('start')))
    then attribute w:start { $pageNumbering('start') } else (),
  if (exists($pageNumbering('format')) and $pageNumbering('format') != 'decimal')
    then attribute w:fmt { $pageNumbering('format') } else (),
  if (exists($pageNumbering('chapterStyle')))
    then attribute w:chapStyle { $pageNumbering('chapterStyle') } else (),
  if (exists($pageNumbering('chapterSeparator')))
    then attribute w:chapSep { $pageNumbering('chapterSeparator') } else ()
} else (),
```

Bindings:

```typescript
pageNumbering: data.pageNumbering ?? null,
lineNumbering: data.lineNumbering ?? null,
```

`?? null` maps `undefined` to `null` (omit element). An explicit `{}` binding is preserved so `exists($pageNumbering)` emits an empty element.

**`change` (track changes) nesting:** When `data.change` is set, recurse via `sectionPropertiesToNode(data.change)` as today — new props must flow through nested `sectPrChange` the same way as `pageMargin`, etc.

### 4.4 JSX / imperative usage (consumer example)

`SectionProps` is an alias of `SectionProperties`, so declarative usage works without further API work:

```tsx
<Section pageNumbering={{ start: 1 }} headers={headerRelId}>
  …
</Section>

<Section lineNumbering={{ countBy: 5 }}>
  …
</Section>
```

Verosetta `DocxmlVisitor` will call `paragraph.setSectionProperties({ … })` at section breaks. Equivalent imperative usage:

```typescript
paragraph.setSectionProperties({
	pageNumbering: { start: 1 },
	headers: headerRelId,
});

paragraph.setSectionProperties({
	lineNumbering: { countBy: 5 },
});
```

Header creation (unchanged docxml API):

```tsx
const header = docx.document.headers.add(
	'word/header-figures.xml',
	<Paragraph alignment="right">
		<Text>
			<FieldRangeStart />
			<FieldRangeInstruction>PAGE \* MERGEFORMAT</FieldRangeInstruction>
			<FieldRangeSeparator />
			<FieldRangeEnd />
		</Text>
		<Text>/</Text>
		<Text>
			<FieldRangeStart />
			<FieldRangeInstruction>SECTIONPAGES</FieldRangeInstruction>
			<FieldRangeSeparator />
			<FieldRangeEnd />
		</Text>
	</Paragraph>,
);
```

### 4.5 P0 tests

Add round-trip cases to `section-properties.test.ts`:

| Test name                         | Input / action                                               | Expected                                      |
| --------------------------------- | ------------------------------------------------------------ | --------------------------------------------- |
| `pageNumbering start only`        | `<w:pgNumType w:start="1"/>`                                 | `{ pageNumbering: { start: 1 } }`             |
| `pageNumbering empty element`     | `<w:pgNumType/>`                                             | `{ pageNumbering: {} }`                       |
| `pageNumbering full attributes`   | all four `w:pgNumType` attributes                            | matching object; `distance` N/A               |
| `pageNumbering format omitted`    | `<w:pgNumType w:start="1"/>` (no `w:fmt`)                    | `format` property absent (not `'decimal'`)    |
| `lineNumbering countBy only`      | `<w:lnNumType w:countBy="5"/>`                               | `{ lineNumbering: { countBy: 5 } }`           |
| `lineNumbering full`              | all four attributes                                          | matching object with `distance` as `Length`   |
| `combined`                        | both elements in one `sectPr`                                | both props present                            |
| `clear pageNumbering`             | `reverseTest({ pageNumbering: null }, sectPr w/o pgNumType)` | no `w:pgNumType` in output                    |
| `empty pageNumbering object`      | `reverseTest({ pageNumbering: {} }, '<w:pgNumType/>')`       | empty element emitted                         |
| `sectPrChange with pageNumbering` | `w:sectPrChange` wrapping `w:pgNumType`                      | `change.pageNumbering` populated; round-trips |

Also verify **object → XML → object** round-trip for Verosetta-minimal shapes:

```typescript
{
	pageNumbering: {
		start: 1;
	}
}
{
	lineNumbering: {
		countBy: 5;
	}
}
```

**Acceptance:** emitted XML matches reference attribute sets from §3.1 and §3.2 (no `w:docGrid`).

---

## 5. P1 — `SettingsI.compatibility`

**Files:**

- `lib/files/src/SettingsXml.ts` — `SettingsI`, `DEFAULT_SETTINGS`, `settingsMeta`, `toNode`, `fromArchive`
- `lib/files/test/SettingsXml.test.ts`

### 5.1 TypeScript API

```typescript
export type CompatSetting = {
	/** w:compatSetting/@w:name */
	name: string;
	/** w:compatSetting/@w:uri */
	uri: string;
	/** w:compatSetting/@w:val */
	val: string;
};

export type CompatibilitySettings = {
	/**
	 * Ordered list of w:compatSetting children.
	 * Empty array is equivalent to null (omit w:compat).
	 */
	settings: CompatSetting[];
};

// Extend SettingsI:
compatibility: CompatibilitySettings | null;
```

**Default:** `compatibility: null` in `DEFAULT_SETTINGS`.

**Semantics:**

| Value               | `toNode()` behaviour                                     |
| ------------------- | -------------------------------------------------------- |
| `null`              | Omit `w:compat` entirely (**Verosetta export path**)     |
| `{ settings: [] }`  | Treat as `null` (omit)                                   |
| `{ settings: […] }` | Emit `<w:compat>` with one `<w:compatSetting>` per entry |

`Docx.withSettings({ compatibility: null })` is a no-op on fresh documents (already default). Explicit `set('compatibility', null)` clears any parsed settings before save.

**Round-trip policy (v1):** Only `w:compatSetting` children are modelled. Other `w:compat` children (`w:spaceForUL`, `w:useFELayout`, etc.) are **ignored on read and not re-emitted on write**. A document whose `w:compat` contains only those children parses as `compatibility: null` and loses them on save — acceptable for Verosetta v1.

**Verosetta consumer note:** No `withSettings` call is required for OFF behaviour on `fromNothing()` exports — absence of `w:compat` is the target. P1 exists so behaviour is tested, documented, and round-trip-safe for `compatSetting` rows.

### 5.2 `settingsMeta` registration

Register for `set` / `get` / `entries` discovery:

```typescript
{
  docxmlName: 'compatibility',
  ooxmlLocalName: 'compat',
  ooxmlType: SettingType.Formatting, // marker only — not handled by generic switch
}
```

Because `w:compat` is a container (not a simple on/off), implement **explicit branches** in `toNode()` and `fromArchive()` — the same pattern as `defaultTabStop` (dedicated XPath block in `fromArchive`, inline XQuery in `toNode()`). Do **not** rely on the generic on/off/length `settingsMeta` switch for serialisation.

### 5.3 `toNode()` emission

After existing settings elements, conditionally emit:

```xml
<w:compat>
  <w:compatSetting w:name="…" w:uri="…" w:val="…"/>
  …
</w:compat>
```

Omit the entire `w:compat` block when `compatibility` is `null` or `settings` is empty. Do **not** emit `w:spaceForUL` or other non-`compatSetting` children.

### 5.4 `fromArchive()` parsing

Add a dedicated block in `fromArchive()` (alongside the existing `defaultTabStop` XPath), not only the partial top-level settings map:

```xpath
"compatibility": map {
  "settings": array {
    ./w:compat/w:compatSetting/map {
      "name": @w:name/string(),
      "uri": @w:uri/string(),
      "val": @w:val/string()
    }
  }
}
```

Post-process:

- `w:compat` absent → `compatibility` stays `null` (default).
- `w:compat` present but zero `compatSetting` children → `compatibility: null`.
- One or more `compatSetting` children → `compatibility: { settings: […] }` preserving document order.

### 5.5 P1 tests

| Test                         | Assertion                                                                                                   |
| ---------------------------- | ----------------------------------------------------------------------------------------------------------- |
| Default settings             | `get('compatibility') === null`; `Docx.fromNothing().toArchive()` settings XML has no `w:compat`            |
| Set compat settings          | Round-trip fixture from §5.6                                                                                |
| Clear compat                 | `set('compatibility', null)` after parse → no `w:compat` on save                                            |
| `withSettings` integration   | `Docx.fromNothing().withSettings({ compatibility: { settings: […] } })` produces expected XML               |
| `cloneAsEmptyTemplate`       | Template cloned with `compatibility` set retains settings via `entries()`                                   |
| Non-`compatSetting` children | Load XML with `<w:compat><w:spaceForUL/></w:compat>` only → `compatibility === null`; save omits `w:compat` |

Optional fixture test: load `_OFF` `settings.xml` via `Docx.fromArchive`, assert parsed `compatibility.settings` includes `name === 'compatibilityMode'`, then re-save with `compatibility: null` and assert output omits `w:compat`.

### 5.6 Test fixture (`compatSetting` rows)

Use these values in unit tests (subset of typical Word / `_OFF` sample rows). Add as inline XML in `SettingsXml.test.ts` or under `assets/` if a file fixture is preferred:

```xml
<w:compat>
  <w:compatSetting w:name="compatibilityMode"
    w:uri="http://schemas.microsoft.com/office/word"
    w:val="15"/>
  <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification"
    w:uri="http://schemas.microsoft.com/office/word"
    w:val="1"/>
  <w:compatSetting w:name="enableOpenTypeFeatures"
    w:uri="http://schemas.microsoft.com/office/word"
    w:val="1"/>
</w:compat>
```

Expected parsed object:

```typescript
{
  compatibility: {
    settings: [
      { name: 'compatibilityMode', uri: 'http://schemas.microsoft.com/office/word', val: '15' },
      { name: 'overrideTableStyleFontSizeAndJustification', uri: 'http://schemas.microsoft.com/office/word', val: '1' },
      { name: 'enableOpenTypeFeatures', uri: 'http://schemas.microsoft.com/office/word', val: '1' },
    ],
  },
}
```

Verify against `adk-sr/.scratches/feature/off/ooxml/word/settings.xml` when available; adjust values if the sample differs.

---

## 6. Exports and versioning

### 6.1 `mod.ts` exports

Add explicit re-exports (none of these are public today):

```typescript
export type { SectionProperties } from './lib/properties/src/section-properties.ts';
export type {
	SettingsI,
	CompatSetting,
	CompatibilitySettings,
} from './lib/files/src/SettingsXml.ts';
```

`SectionProps` remains an alias of `SectionProperties` on the `Section` component. Consumers may use either name.

### 6.2 Version bump and build

The npm package version is set at build time, not by hand-editing `build/npm/package.json` (generated by dnt):

```bash
deno run -A tasks/build-npm.ts 1.2.0
```

### 6.3 Changelog

Create `CHANGELOG.md` at repo root (file does not exist yet) with a `1.2.0` section covering P0 and P1.

### 6.4 Publish

Publish to npm registry; update Verosetta `pnpm-workspace.yaml` catalog `'@feecompass/docxml': '=1.2.0'` and **remove** the `file://` override.

**Backward compatibility:** Additive only. Existing callers unaffected. Default document settings unchanged.

---

## 7. Implementation checklist

### P0

- [ ] Add `pageNumbering` / `lineNumbering` to `SectionProperties` type + JSDoc
- [ ] Extend `sectionPropertiesFromNode` XPath map + `dropEmptySectionProp` normalisation
- [ ] Extend `sectionPropertiesToNode` XQuery + bindings (`lnNumType` before `pgNumType`; omit `w:fmt` when decimal)
- [ ] Round-trip tests per §4.5 (including clear, empty `{}`, `sectPrChange`)
- [ ] Manual smoke: build npm package, Verosetta override, export DOCX with FIG/LINENB sections

### P1

- [ ] Add `CompatSetting`, `CompatibilitySettings`, extend `SettingsI`
- [ ] Implement dedicated `toNode` / `fromArchive` branches for `w:compat`
- [ ] Register in `settingsMeta`; wire `set` / `get` / `entries` / `withSettings`
- [ ] Unit tests per §5.5 (including `fromNothing` integration and non-`compatSetting` drop)
- [ ] Confirm `fromNothing()` still omits `w:compat` without explicit call

### Release

- [ ] `deno run -A tasks/build-npm.ts 1.2.0`
- [ ] Add `CHANGELOG.md` entry for `1.2.0`
- [ ] Export types from `mod.ts` per §6.1
- [ ] Publish; Verosetta: bump catalog, drop override, integration tests in `backend-srv`

---

## 8. Consumer mapping (Verosetta — informational)

Not implemented in docxml; documented for traceability.

| Verosetta variant | docxml APIs used                                                                              |
| ----------------- | --------------------------------------------------------------------------------------------- |
| **VAR**           | `docx.customProperties.add()` (existing)                                                      |
| **OFF**           | Default `SettingsI.compatibility === null` (P1) — omit `w:compat` on export                   |
| **FIG**           | `pageNumbering: { start: 1 }` (P0) + `headers.add()` + `SectionProperties.headers` (existing) |
| **LINENB**        | `lineNumbering: { countBy: 5 }` (P0)                                                          |

---

## 9. Open questions

| #    | Question                                                                           | Resolution                                                                                                                                                 |
| ---- | ---------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| OQ-1 | `_OFF` sample contains `compatibilityMode` — contradicts “compatibility off” label | **Resolved:** OFF = omit `w:compat` on export. `compatibilityMode` is a `compatSetting` row, not a separate element. UAT with Artur for image positioning. |
| OQ-2 | Should P1 support `w:useFELayout` under `w:compat`?                                | Defer — not needed for Verosetta v1; v1 may drop on round-trip (§5.1)                                                                                      |
| OQ-3 | `w:pgNumType` / `w:lnNumType` order inside `w:sectPr`                              | **Resolved:** emit `lnNumType` then `pgNumType` after `pgMar`, before `titlePg` (§4.3)                                                                     |

---

## 10. Version history

| Version | Date       | Changes                                                                                                |
| ------- | ---------- | ------------------------------------------------------------------------------------------------------ |
| 1.0     | 2026-07-22 | Initial P0+P1 spec; `w:docGrid` explicitly excluded                                                    |
| 1.1     | 2026-07-22 | Round-trip policy, normalisation, test fixtures, exports/build instructions, open-question resolutions |
