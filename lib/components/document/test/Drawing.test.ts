import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { RelationshipsXml } from '../../../files/src/RelationshipsXml.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { emu } from '../../../utilities/src/length.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { Drawing } from '../src/Drawing.ts';
import { GraphicData } from '../src/GraphicData.ts';
import { Inline } from '../src/Inline.ts';
import { Picture } from '../src/Picture.ts';

const baseContext: ComponentContext = {
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

describe('Drawing', () => {
	it('deserializes from w:drawing with wp:inline', () => {
		const dom = create(`
			<w:drawing xmlns:w="${NamespaceUri.w}" xmlns:wp="${NamespaceUri.wp}" xmlns:a="${NamespaceUri.a}" xmlns:pic="${NamespaceUri.pic}" xmlns:r="${NamespaceUri.r}">
				<wp:inline>
					<wp:extent cx="1000" cy="2000"/>
					<wp:docPr id="1" name="title" descr="alt"/>
					<wp:cNvGraphicFramePr>
						<a:graphicFrameLocks noChangeAspect="1"/>
					</wp:cNvGraphicFramePr>
					<a:graphic>
						<a:graphicData uri="${NamespaceUri.pic}">
							<pic:pic>
								<pic:nvPicPr>
									<pic:cNvPr id="0" name="title" descr="alt"/>
									<pic:cNvPicPr/>
								</pic:nvPicPr>
								<pic:blipFill>
									<a:blip r:embed="rIdImg" cstate="print"/>
									<a:stretch><a:fillRect/></a:stretch>
								</pic:blipFill>
								<pic:spPr>
									<a:xfrm>
										<a:off x="0" y="0"/>
										<a:ext cx="1000" cy="2000"/>
									</a:xfrm>
									<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
								</pic:spPr>
							</pic:pic>
						</a:graphicData>
					</a:graphic>
				</wp:inline>
			</w:drawing>
		`);
		const drawing = Drawing.fromNode(dom, baseContext);
		expect(drawing).toBeTruthy();
		expect(drawing?.children).toHaveLength(1);
	});

	it('throws on wp:anchor', () => {
		const dom = create(`
			<w:drawing xmlns:w="${NamespaceUri.w}" xmlns:wp="${NamespaceUri.wp}">
				<wp:anchor />
			</w:drawing>
		`);
		expect(() => Drawing.fromNode(dom, baseContext)).toThrow();
	});

	it('serializes to w:drawing (programmatic tree)', async () => {
		const pic = new Picture({
			// PNG header bytes for mime recognition
			data: Promise.resolve(new Uint8Array([0x89, 0x50, 0x4e, 0x47])),
			width: emu(1000),
			height: emu(2000),
			title: 'title',
			alt: 'alt',
		});
		const gd = new GraphicData({}, pic);
		const inline = new Inline(
			{ width: emu(1000), height: emu(2000), title: 'title', alt: 'alt' },
			gd
		);
		const drawing = new Drawing({}, inline);

		const rels = new RelationshipsXml('word/_rels/document.xml.rels');
		await pic.ensureRelationship(rels);

		const xml = serialize(await drawing.toNode([]));
		expect(xml).toContain('<drawing');
	});
});
