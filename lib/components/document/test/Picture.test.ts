import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import { BinaryFile } from '../../../classes/src/BinaryFile.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { FileMime, RelationshipType } from '../../../enums.ts';
import { RelationshipsXml } from '../../../files/src/RelationshipsXml.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { emu } from '../../../utilities/src/length.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { Picture } from '../src/Picture.ts';

describe('Picture', () => {
	it('deserializes from pic:pic using relationships', () => {
		const archive = new Archive({
			// PNG header bytes for mime recognition
			'word/media/image.bin': new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
		});
		const relationships = new RelationshipsXml(
			'word/_rels/document.xml.rels',
			[
				{
					id: 'rIdMain',
					type: RelationshipType.image,
					target: 'word/media/image.bin',
					isExternal: false,
					isBinary: true,
				},
			]
		);
		const context: ComponentContext = { archive, relationships };

		const dom = create(`
			<pic:pic xmlns:pic="${NamespaceUri.pic}" xmlns:a="${NamespaceUri.a}" xmlns:r="${NamespaceUri.r}">
				<pic:nvPicPr><pic:cNvPr id="0" name="img" descr="desc"/><pic:cNvPicPr/></pic:nvPicPr>
				<pic:blipFill>
					<a:blip r:embed="rIdMain" cstate="print"/>
					<a:stretch><a:fillRect/></a:stretch>
				</pic:blipFill>
				<pic:spPr>
					<a:xfrm><a:off x="0" y="0"/><a:ext cx="111" cy="222"/></a:xfrm>
					<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
				</pic:spPr>
			</pic:pic>
		`);

		const picture = Picture.fromNode(dom, context)!;
		expect(picture.props.title).toBe('img');
		expect(picture.props.alt).toBe('desc');
		expect(Math.round(picture.props.width.emu)).toBe(111);
		expect(Math.round(picture.props.height.emu)).toBe(222);
		expect(picture.meta.location).toBe('word/media/image.bin');
	});

	it('serializes after ensureRelationship', async () => {
		const picture = new Picture({
			// JPEG header bytes for mime recognition
			data: Promise.resolve(new Uint8Array([0xff, 0xd8, 0xff, 0xe0])),
			width: emu(1000),
			height: emu(2000),
			title: 'T',
			alt: 'A',
		});
		const relationships = new RelationshipsXml(
			'word/_rels/document.xml.rels'
		);
		await picture.ensureRelationship(relationships);
		const xml = serialize(picture.toNode([]) as Node);
		expect(xml).toContain('<pic');
		expect(xml).toContain('cNvPr');
		expect(xml).toContain('blip');
		expect(xml).toContain('ext'); // xfrm ext with cx/cy
	});

	it('serializes with SVG extension when provided', async () => {
		const picture = new Picture({
			// GIF header bytes for mime recognition
			data: Promise.resolve(new Uint8Array([0x47, 0x49, 0x46, 0x38])),
			dataExtensions: { svg: Promise.resolve('<svg/>') },
			width: emu(10),
			height: emu(20),
			title: 'S',
			alt: 'V',
		});
		const relationships = new RelationshipsXml(
			'word/_rels/document.xml.rels'
		);
		await picture.ensureRelationship(relationships);
		const xml = serialize(picture.toNode([]) as Node);
		// Expect extLst and svgBlip presence
		expect(xml).toContain('extLst');
		expect(xml).toContain('svgBlip');
	});
});
