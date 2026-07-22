import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../classes/src/Archive.ts';
import { Docx } from '../../Docx.ts';
import { FileLocation, RelationshipType } from '../../enums.ts';
import { serialize } from '../../utilities/src/dom.ts';
import { pt } from '../../utilities/src/length.ts';
import { evaluateXPathToBoolean } from '../../utilities/src/xquery.ts';
import { ContentTypesXml } from '../src/ContentTypesXml.ts';
import { type CompatibilitySettings, SettingsXml } from '../src/SettingsXml.ts';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const COMPAT_FIXTURE: CompatibilitySettings = {
	settings: [
		{
			name: 'compatibilityMode',
			uri: 'http://schemas.microsoft.com/office/word',
			val: '15',
		},
		{
			name: 'overrideTableStyleFontSizeAndJustification',
			uri: 'http://schemas.microsoft.com/office/word',
			val: '1',
		},
		{
			name: 'enableOpenTypeFeatures',
			uri: 'http://schemas.microsoft.com/office/word',
			val: '1',
		},
	],
};

const COMPAT_SETTINGS_XML = `<w:settings xmlns:w="${WORD_NS}">
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
</w:settings>`;

async function settingsFromXml(xml: string): Promise<SettingsXml> {
	const archive = new Archive();
	archive.addTextFile(FileLocation.settings, xml);
	const contentTypes = await ContentTypesXml.fromArchive(
		archive,
		FileLocation.contentTypes,
	).catch(() => new ContentTypesXml(FileLocation.contentTypes));
	return SettingsXml.fromArchive(
		archive,
		contentTypes,
		FileLocation.settings,
	);
}

describe('SettingsXml', () => {
	it('evenAndOddHeaders', () => {
		const settings = new SettingsXml('test');
		expect(settings.get('evenAndOddHeaders')).toBe(false);
		settings.set('evenAndOddHeaders', true);
		expect(settings.get('evenAndOddHeaders')).toBe(true);
	});
	it('isTrackChangesEnabled', () => {
		const settings = new SettingsXml('test');
		expect(settings.get('isTrackChangesEnabled')).toBe(false);
		settings.set('isTrackChangesEnabled', true);
		expect(settings.get('isTrackChangesEnabled')).toBe(true);
	});
	it('attachedTemplate', () => {
		const settings = new SettingsXml('test');
		expect(settings.get('attachedTemplate')).toBe(null);
		settings.set('attachedTemplate', 'foobar');
		expect(settings.get('attachedTemplate')).toBe('foobar');
		const meta = settings.relationships.meta.find(
			(meta) => meta.type === RelationshipType.attachedTemplate,
		);
		expect(meta).toBeTruthy();
		expect(settings.relationships.getTarget(meta?.id as string)).toBe(
			'foobar',
		);
	});
	it('defaultTabStop', () => {
		const settings = new SettingsXml('test');
		expect(settings.get('defaultTabStop')).toBe(null);
		settings.set('defaultTabStop', pt(50));
		expect(settings.get('defaultTabStop')).toEqual(pt(50));
	});
	it('footnoteProperties', () => {
		const settings = new SettingsXml('test');
		expect(settings.get('footnoteProperties')).toBe(null);
		settings.set('footnoteProperties', {
			restart: 'continuous',
			numberingFormat: 'lowerRoman',
			position: 'beneathText',
		});
		expect(settings.get('footnoteProperties')).toEqual({
			restart: 'continuous',
			numberingFormat: 'lowerRoman',
			position: 'beneathText',
		});
	});

	describe('compatibility', () => {
		it('defaults to null', () => {
			const settings = new SettingsXml('test');
			expect(settings.get('compatibility')).toBe(null);
		});

		it('normalizes empty settings array to null on set', () => {
			const settings = new SettingsXml('test');
			settings.set('compatibility', { settings: [] });
			expect(settings.get('compatibility')).toBe(null);
		});

		it('Docx.fromNothing() omits w:compat in settings XML', async () => {
			const docx = Docx.fromNothing();
			const settingsNode = await docx.document.settings.$$$toNode();
			expect(
				evaluateXPathToBoolean(
					`boolean(/*/*[local-name() = 'compat'])`,
					settingsNode.documentElement,
				),
			).toBe(false);
		});

		it('round-trips compatSetting rows from archive', async () => {
			const settings = await settingsFromXml(COMPAT_SETTINGS_XML);
			expect(settings.get('compatibility')).toEqual(COMPAT_FIXTURE);

			const serialized = serialize(await settings.$$$toNode());
			expect(serialized).toContain('w:compat');
			expect(serialized).toContain('w:name="compatibilityMode"');
			expect(serialized).toContain('w:val="15"');

			const roundTripped = await settingsFromXml(serialized);
			expect(roundTripped.get('compatibility')).toEqual(COMPAT_FIXTURE);
		});

		it('clears compatibility on set(null)', async () => {
			const settings = await settingsFromXml(COMPAT_SETTINGS_XML);
			settings.set('compatibility', null);
			const serialized = serialize(await settings.$$$toNode());
			expect(serialized).not.toContain('w:compat');
		});

		it('withSettings integration produces expected XML', async () => {
			const docx = Docx.fromNothing().withSettings({
				compatibility: COMPAT_FIXTURE,
			});
			const serialized = serialize(
				await docx.document.settings.$$$toNode(),
			);
			expect(serialized).toContain('w:compatSetting');
			expect(serialized).toContain('w:name="enableOpenTypeFeatures"');
			expect(docx.document.settings.get('compatibility')).toEqual(
				COMPAT_FIXTURE,
			);
		});

		it('cloneAsEmptyTemplate retains compatibility settings', () => {
			const original = Docx.fromNothing().withSettings({
				compatibility: COMPAT_FIXTURE,
			});
			const clone = original.cloneAsEmptyTemplate();
			expect(clone.document.settings.get('compatibility')).toEqual(
				COMPAT_FIXTURE,
			);
		});

		it('ignores non-compatSetting children on read and omits w:compat on save', async () => {
			const xml = `<w:settings xmlns:w="${WORD_NS}">
				<w:compat>
					<w:spaceForUL/>
				</w:compat>
			</w:settings>`;
			const settings = await settingsFromXml(xml);
			expect(settings.get('compatibility')).toBe(null);
			const serialized = serialize(await settings.$$$toNode());
			expect(serialized).not.toContain('w:compat');
		});
	});
});
