import { assertSnapshot } from 'jsr:@std/testing/snapshot';
import { expect } from 'std/expect';
import { beforeAll, describe, it } from 'std/testing/bdd';

import { crypto } from 'jsr:@std/crypto';
import { encodeHex } from 'jsr:@std/encoding/hex';
import { Section } from './components/document/src/Section.ts';
import { Docx } from './Docx.ts';
import { FileLocation } from './enums.ts';
import { RelationshipMeta } from './files/src/RelationshipsXml.ts';
import { file } from './utilities/src/tests.ts';

async function normalizeXmlIds(xml: string, relationships: RelationshipMeta[]) {
	if (!relationships.length) {
		return xml;
	}
	const idMap = new Map<string, string>();
	for (let index = 0; index < relationships.length; index++) {
		const originalId = relationships[index]?.id;
		if (!originalId) {
			continue;
		}
		const md5 = await crypto.subtle.digest(
			'SHA-256',
			new TextEncoder().encode(String(index))
		);
		const md5Hex = encodeHex(md5);
		const stableId = `r-${index}-${md5Hex}`;
		idMap.set(originalId, stableId);
	}
	let normalized = xml;
	for (const [oldId, newId] of idMap.entries()) {
		normalized = normalized.split(oldId).join(newId);
	}
	return normalized;
}

describe('Docx', () => {
	it.skip('read and writes back inline image', async (t) => {
		const docx = await Docx.fromArchive(
			file('../assets/debug_min_inline.docx')
		);
    const archive = await docx.toArchive();
		const xml = await archive.readText('word/document.xml');
		const relationships = docx.document.relationships;
		const normalizedXml = await normalizeXmlIds(xml, relationships.meta);

		await assertSnapshot(t, normalizedXml);
		// await archive.toFile('./inline-output.docx');
	});

	it.skip('read and writes back anchor image', async (t) => {
		const docx = await Docx.fromArchive(
			file('../assets/debug_min_anchor.docx')
		);

		const archive = await docx.toArchive();
		const xml = await archive.readText('word/document.xml');
		const relationships = docx.document.relationships;
		const normalizedXml = await normalizeXmlIds(xml, relationships.meta);

		await assertSnapshot(t, normalizedXml);
		// await archive.toFile('./anchor-output.docx');
	});

	it('read and writes back wpg image', async (t) => {
		const docx = await Docx.fromArchive(
			file('../assets/wpg_image_nomc.docx')
		);

		const archive = await docx.toArchive();
		const xml = await archive.readText('word/document.xml');
		const relationships = docx.document.relationships;
		const normalizedXml = await normalizeXmlIds(xml, relationships.meta);

		await assertSnapshot(t, normalizedXml);
		// await archive.toFile('./output.docx');
	});

});
