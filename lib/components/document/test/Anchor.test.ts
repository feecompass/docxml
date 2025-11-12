import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { RelationshipType } from '../../../enums.ts';
import { RelationshipsXml } from '../../../files/src/RelationshipsXml.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { emu } from '../../../utilities/src/length.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { Anchor } from '../src/Anchor.ts';
import { GraphicData } from '../src/GraphicData.ts';
import { Picture } from '../src/Picture.ts';

describe('Anchor', () => {
	it('deserializes from wp:anchor', () => {
		const archive = new Archive({
			'word/media/image.bin': new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
		});
		const relationships = new RelationshipsXml(
			'word/_rels/document.xml.rels',
			[
				{
					id: 'rId8',
					type: RelationshipType.image,
					target: 'word/media/image.bin',
					isExternal: false,
					isBinary: true,
				},
			]
		);
		const context: ComponentContext = {
			archive,
			relationships,
		};

		const dom = create(`
			<wp:anchor xmlns:wp="${NamespaceUri.wp}" xmlns:a="${NamespaceUri.a}" xmlns:pic="${NamespaceUri.pic}" xmlns:r="${NamespaceUri.r}"
				allowOverlap="1" behindDoc="0" distB="0" distT="0" distL="114300" distR="114300"
				hidden="0" layoutInCell="1" locked="0" relativeHeight="0" simplePos="0">
				<wp:simplePos x="0" y="0" />
				<wp:positionH relativeFrom="column"><wp:posOffset>1832759</wp:posOffset></wp:positionH>
				<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
				<wp:extent cx="2095200" cy="6271200" />
				<wp:effectExtent b="0" l="0" r="0" t="0" />
				<wp:wrapSquare wrapText="bothSides" distB="0" distT="0" distL="114300" distR="114300" />
				<wp:docPr id="1517777417" name="image1.png" descr="Alt" />
				<a:graphic>
					<a:graphicData uri="${NamespaceUri.pic}">
						<pic:pic>
							<pic:nvPicPr>
								<pic:cNvPr id="0" name="image1.png" descr="Alt" />
								<pic:cNvPicPr />
							</pic:nvPicPr>
							<pic:blipFill>
								<a:blip r:embed="rId8" />
								<a:stretch><a:fillRect /></a:stretch>
							</pic:blipFill>
							<pic:spPr>
								<a:xfrm><a:off x="0" y="0" /><a:ext cx="2095200" cy="6271200" /></a:xfrm>
								<a:prstGeom prst="rect" />
							</pic:spPr>
						</pic:pic>
					</a:graphicData>
				</a:graphic>
			</wp:anchor>
		`);
		const anchor = Anchor.fromNode(dom, context)!;
		expect(anchor.props.title).toBe('image1.png');
		expect(anchor.props.wrap.type).toBe('square');
	});

	it('serializes to wp:anchor (programmatic tree)', async () => {
		const pic = new Picture({
			data: Promise.resolve(new Uint8Array([0x89, 0x50, 0x4e, 0x47])),
			width: emu(2095200),
			height: emu(6271200),
			title: 'image1.png',
			alt: 'Alt',
		});
		const gd = new GraphicData({}, pic);
		const anchor = new Anchor(
			{
				title: 'image1.png',
				alt: 'Alt',
				width: emu(2095200),
				height: emu(6271200),
				positionH: { relativeFrom: 'column', offset: emu(1832759) },
				positionV: { relativeFrom: 'paragraph', offset: emu(0) },
				wrap: {
					type: 'square',
					wrapText: 'bothSides',
					dist: {
						left: emu(114300),
						right: emu(114300),
						top: emu(0),
						bottom: emu(0),
					},
				},
			},
			gd
		);
		const rels = new RelationshipsXml('word/_rels/document.xml.rels');
		await pic.ensureRelationship(rels);
		const xml = serialize(await anchor.toNode([]));
		expect(xml).toContain('<anchor');
		expect(xml).toContain('wrapSquare');
	});
});
