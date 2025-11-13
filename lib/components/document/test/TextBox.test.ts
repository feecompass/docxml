import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { create, serialize } from '../../../utilities/src/dom.ts';
import { emu } from '../../../utilities/src/length.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { Paragraph } from '../src/Paragraph.ts';
import { TextBox } from '../src/TextBox.ts';

describe('TextBox', () => {
	it('deserializes solidFill and bodyPr when present', () => {
		const dom = create(`
			<wps:wsp xmlns:wps="${NamespaceUri.wps}" xmlns:a="${NamespaceUri.a}" xmlns:w="${NamespaceUri.w}">
				<wps:cNvPr id="1" name="tb"/>
				<wps:cNvSpPr txBox="1"/>
				<wps:spPr bwMode="auto">
					<a:xfrm><a:off x="1828800" y="819150"/><a:ext cx="790575" cy="123825"/></a:xfrm>
					<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
					<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
				</wps:spPr>
				<wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>
				<wps:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"/>
			</wps:wsp>
		`);
		const tb = TextBox.fromNode(dom, {
			archive: undefined as any,
			relationships: null,
		} as any)!;
		expect(tb.props.fill?.type).toBe('solid');
		expect(tb.props.fill?.color).toBe('FFFFFF');
		expect(tb.props.bodyPr?.anchor).toBe('ctr');
		expect(tb.props.bodyPr?.lIns?.emu).toBe(0);
		expect(tb.props.bodyPr?.rIns?.emu).toBe(0);
		expect(tb.props.bodyPr?.tIns?.emu).toBe(0);
		expect(tb.props.bodyPr?.bIns?.emu).toBe(0);
	});

	it('serializes optional solidFill and bodyPr only when provided', async () => {
		const tb = new TextBox(
			{
				title: 'tb',
				width: emu(790575),
				height: emu(123825),
				offsetX: emu(1828800),
				offsetY: emu(819150),
				fill: { type: 'solid', color: 'FFFFFF' },
				bodyPr: {
					lIns: emu(1),
					rIns: emu(0),
					tIns: emu(0),
					bIns: emu(0),
					anchor: 'ctr',
				},
			},
			new Paragraph({})
		);
		const xml = serialize(await tb.toNode([]));
		expect(xml).toContain('solidFill');
		expect(xml).toContain('srgbClr');
		expect(xml).toContain('bodyPr');
		expect(xml).toContain('anchor="ctr"');
		expect(xml).toContain('lIns="1"');
		expect(xml).toContain('rIns="0"');
		expect(xml).toContain('tIns="0"');
		expect(xml).toContain('bIns="0"');
	});
});
