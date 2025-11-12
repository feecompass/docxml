import { assertSnapshot } from 'jsr:@std/testing/snapshot';
import { expect } from 'std/expect';
import { beforeAll, describe, it } from 'std/testing/bdd';

import { crypto } from 'jsr:@std/crypto';
import { encodeHex } from 'jsr:@std/encoding/hex';
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
	let bundle: Docx;
	beforeAll(async () => {
		bundle = await Docx.fromArchive(file('../assets/inlineImage.docx'));
	});

	it('read and writes back inline image', async (t) => {
		const docx = await Docx.fromArchive(file('../assets/inlineImage.docx'));
		const archive = await docx.toArchive();
		const xml = await archive.readText('word/document.xml');
		const relationships = docx.document.relationships;
		const normalizedXml = await normalizeXmlIds(xml, relationships.meta);

		await assertSnapshot(t, normalizedXml);
		await archive.toFile('./inlineImage-output.docx');
	});
});
