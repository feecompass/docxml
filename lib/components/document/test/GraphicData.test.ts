import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { RelationshipsXml } from '../../../files/src/RelationshipsXml.ts';
import { emu } from '../../../utilities/src/length.ts';
import { Picture } from '../src/Picture.ts';
import { GraphicData } from '../src/GraphicData.ts';

const context: ComponentContext = {
	archive: new Archive({
		// PNG header bytes for mime recognition
		'word/media/image.bin': new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
	}),
	relationships: {
		location: 'word/_rels/document.xml.rels',
		meta: [
			{
				id: 'rIdImg',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
				target: 'word/media/image.bin',
				isExternal: false,
				isBinary: true,
			},
		],
		getTarget(id: string) {
			if (id === 'rIdImg') return 'word/media/image.bin';
			throw new Error('unknown id');
		},
	} as unknown as ComponentContext['relationships'],
};

describe('GraphicData', () => {
	it('deserializes with child pic:pic', () => {
		const dom = create(`
			<a:graphic xmlns:a="${NamespaceUri.a}" xmlns:pic="${NamespaceUri.pic}" xmlns:r="${NamespaceUri.r}">
				<a:graphicData uri="${NamespaceUri.pic}">
					<pic:pic>
						<pic:nvPicPr><pic:cNvPr id="0" name="n" descr="d"/><pic:cNvPicPr/></pic:nvPicPr>
						<pic:blipFill><a:blip r:embed="rIdImg" cstate="print"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
						<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="20"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		`);
		const gd = GraphicData.fromNode(dom, context)!;
		expect(gd.children).toHaveLength(1);
	});

	it('serializes to a:graphic/a:graphicData (programmatic tree)', async () => {
		const pic = new Picture({
			// PNG header bytes for mime recognition
			data: Promise.resolve(
				new Uint8Array([0x89, 0x50, 0x4e, 0x47])
			),
			width: emu(10),
			height: emu(20),
			title: 'n',
			alt: 'd',
		});
		const gd = new GraphicData({}, pic);
		const rels = new RelationshipsXml('word/_rels/document.xml.rels');
		await pic.ensureRelationship(rels);
		const xml = serialize(await gd.toNode([]));
		expect(xml).toContain('<graphic');
		expect(xml).toContain('<graphicData');
	});
});


