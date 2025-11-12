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
import { Inline } from '../src/Inline.ts';
import { Picture } from '../src/Picture.ts';

// const context: ComponentContext = {
// 	archive: new Archive({
// 		// PNG header bytes for mime recognition
// 		'word/media/image.bin': new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
// 	})
//   const relationships = new RelationshipsXml(
//     'word/_rels/document.xml.rels',
//     [
//       {
//         id: 'rId8',
//         type: RelationshipType.image,
//         target: 'word/media/image.bin',
//         isExternal: false,
//         isBinary: true,
//       },
//     ]
//   );
//   const context: ComponentContext = { archive, relationships };

// 	relationships: {
// 		location: 'word/_rels/document.xml.rels',
// 		meta: [
// 			{
// 				id: 'rIdImg',
// 				type: RelationshipType.image,
// 				target: 'word/media/image.bin',
// 				isExternal: false,
// 				isBinary: true,
// 			},
// 		],
// 		getTarget(id: string) {
// 			if (id === 'rIdImg') return 'word/media/image.bin';
// 			throw new Error('unknown id');
// 		},
// 	} satisfies ComponentContext['relationships'],
// };

describe('Inline', () => {
	it('deserializes props and child GraphicData', () => {
		const archive = new Archive({
			// PNG header bytes for mime recognition
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
			<wp:inline xmlns:wp="${NamespaceUri.wp}" xmlns:a="${NamespaceUri.a}" xmlns:pic="${NamespaceUri.pic}" xmlns:r="${NamespaceUri.r}">
				<wp:extent cx="1234" cy="5678"/>
				<wp:docPr id="1" name="my-title" descr="my-alt"/>
				<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
				<a:graphic>
					<a:graphicData uri="${NamespaceUri.pic}">
						<pic:pic>
							<pic:nvPicPr><pic:cNvPr id="0" name="my-title" descr="my-alt"/><pic:cNvPicPr/></pic:nvPicPr>
							<pic:blipFill><a:blip r:embed="rId8" cstate="print"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
							<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1234" cy="5678"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
						</pic:pic>
					</a:graphicData>
				</a:graphic>
			</wp:inline>
		`);
		const inline = Inline.fromNode(dom, context)!;
		expect(inline.props.title).toBe('my-title');
		expect(inline.props.alt).toBe('my-alt');
		expect(Math.round(inline.props.width.emu)).toBe(1234);
		expect(Math.round(inline.props.height.emu)).toBe(5678);
		expect(inline.children).toHaveLength(1);
	});

	it('serializes to wp:inline (programmatic tree)', async () => {
		const pic = new Picture({
			// PNG header bytes for mime recognition
			data: Promise.resolve(new Uint8Array([0x89, 0x50, 0x4e, 0x47])),
			width: emu(2222),
			height: emu(3333),
			title: 't',
			alt: 'd',
		});
		const gd = new GraphicData({}, pic);
		const inline = new Inline(
			{ width: emu(2222), height: emu(3333), title: 't', alt: 'd' },
			gd
		);
		const rels = new RelationshipsXml('word/_rels/document.xml.rels');
		await pic.ensureRelationship(rels);
		const xml = serialize(await inline.toNode([]));
		expect(xml).toContain('<inline');
		expect(xml).toContain('extent');
	});
});
