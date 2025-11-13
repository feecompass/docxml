import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { RelationshipType } from '../../../enums.ts';
import { RelationshipsXml } from '../../../files/src/RelationshipsXml.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { emu } from '../../../utilities/src/length.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { GraphicData } from '../src/GraphicData.ts';
import { Group } from '../src/Group.ts';
import { Paragraph } from '../src/Paragraph.ts';
import { Picture } from '../src/Picture.ts';
import { TextBox } from '../src/TextBox.ts';

describe('Group', () => {
	it('deserializes from wpg:wgp with Picture', () => {
		const archive = new Archive({
			'word/media/image.bin': new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
		});
		const relationships = new RelationshipsXml(
			'word/_rels/document.xml.rels',
			[
				{
					id: 'rId5',
					type: RelationshipType.image,
					target: 'word/media/image.bin',
					isExternal: false,
					isBinary: true,
				},
			]
		);
		const context: ComponentContext = { archive, relationships };
		const dom = create(`
			<a:graphic xmlns:a="${NamespaceUri.a}" xmlns:wpg="${NamespaceUri.wpg}" xmlns:wps="${NamespaceUri.wps}" xmlns:pic="${NamespaceUri.pic}" xmlns:r="${NamespaceUri.r}">
				<a:graphicData uri="${NamespaceUri.wpg}">
					<wpg:wgp>
						<wpg:cNvGrpSpPr><a:grpSpLocks noChangeAspect="1"/></wpg:cNvGrpSpPr>
						<wpg:grpSpPr>
							<a:xfrm>
								<a:off x="0" y="0"/>
								<a:ext cx="1504950" cy="485775"/>
								<a:chOff x="0" y="0"/>
								<a:chExt cx="0" cy="0"/>
							</a:xfrm>
						</wpg:grpSpPr>
						<pic:pic>
							<pic:nvPicPr>
								<pic:cNvPr id="1" name="img"/>
								<pic:cNvPicPr/>
							</pic:nvPicPr>
							<pic:blipFill>
								<a:blip r:embed="rId5"/>
								<a:stretch><a:fillRect/></a:stretch>
							</pic:blipFill>
							<pic:spPr>
								<a:xfrm><a:off x="0" y="0"/><a:ext cx="1504950" cy="485775"/></a:xfrm>
								<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
							</pic:spPr>
						</pic:pic>
					</wpg:wgp>
				</a:graphicData>
			</a:graphic>
		`);
		const gd = GraphicData.fromNode(dom, context)!;
		expect(gd.children[0]).toBeInstanceOf(Group);
	});

	it('serializes Group with Picture', async () => {
		const pic = new Picture({
			data: Promise.resolve(new Uint8Array([0x89, 0x50, 0x4e, 0x47])),
			width: emu(1504950),
			height: emu(485775),
			title: 'img',
			alt: 'desc',
		});
		const group = new Group(
			{
				width: emu(1504950),
				height: emu(485775),
			},
			pic
		);
		const gd = new GraphicData({}, group);
		const rels = new RelationshipsXml('word/_rels/document.xml.rels');
		await pic.ensureRelationship(rels);
		const xml = serialize(await gd.toNode([]));
		expect(xml).toContain('<graphicData');
		expect(xml).toContain('wgp');
		expect(xml).toContain('pic');
	});

	it('serializes Group with TextBox containing a Paragraph', async () => {
		const tb = new TextBox(
			{
				title: 'tb',
				width: emu(619125),
				height: emu(142875),
				offsetX: emu(152400),
				offsetY: emu(0),
			},
			new Paragraph({})
		);
		const group = new Group(
			{
				width: emu(1504950),
				height: emu(485775),
			},
			tb
		);
		const gd = new GraphicData({}, group);
		const xml = serialize(await gd.toNode([]));
		expect(xml).toContain('<graphicData');
		expect(xml).toContain('wgp');
		expect(xml).toContain('wsp');
		expect(xml).toContain('txbxContent');
	});
});
