import { Paragraph } from './Paragraph.ts';

import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import {
	createChildComponentsFromNodes,
	registerComponent,
} from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { createUniqueNumericIdentifier } from '../../../utilities/src/identifiers.ts';
import { type Length } from '../../../utilities/src/length.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import {
	evaluateXPathToFirstNode,
	evaluateXPathToNumber,
	evaluateXPathToString,
} from '../../../utilities/src/xquery.ts';

export type TextBoxChild = Paragraph;

export type TextBoxProps = {
	title?: null | string;
	width: Length;
	height: Length;
	offsetX?: Length;
	offsetY?: Length;
	fill?: {
		type: 'solid';
		color: string;
	};
	bodyPr?: {
		lIns?: Length;
		rIns?: Length;
		tIns?: Length;
		bIns?: Length;
		anchor?: 'top' | 'ctr' | 'bot' | 'just' | 'dist';
	};
};

export class TextBox extends Component<TextBoxProps, TextBoxChild> {
	public static override readonly children: string[] = ['Paragraph'];
	public static override readonly mixed: boolean = false;

	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		const { title, width, height, offsetX, offsetY, fill, bodyPr } =
			this.props;
		return create(
			`
				element ${QNS.wps}wsp {
					element ${QNS.wps}cNvPr {
						attribute id { $identifier },
						attribute name { $name }
					},
					element ${QNS.wps}cNvSpPr {
						attribute txBox { "1" }
					},
					element ${QNS.wps}spPr {
						attribute bwMode { "auto" },
						element ${QNS.a}xfrm {
							element ${QNS.a}off {
								attribute x { $offX },
								attribute y { $offY }
							},
							element ${QNS.a}ext {
								attribute cx { $extX },
								attribute cy { $extY }
							}
						},
						element ${QNS.a}prstGeom {
							attribute prst { "rect" },
							element ${QNS.a}avLst {}
						},
						if ($fillColor) then element ${QNS.a}solidFill {
							element ${QNS.a}srgbClr {
								attribute val { $fillColor }
							}
						} else ()
					},
					element ${QNS.wps}txbx {
						element ${QNS.w}txbxContent {
							$children
						}
					},
					element ${QNS.wps}bodyPr {
						if (exists($lIns)) then attribute lIns { $lIns } else (),
						if (exists($tIns)) then attribute tIns { $tIns } else (),
						if (exists($rIns)) then attribute rIns { $rIns } else (),
						if (exists($bIns)) then attribute bIns { $bIns } else (),
						if (exists($anchor)) then attribute anchor { $anchor } else ()
					}
				}
			`,
			{
				identifier: createUniqueNumericIdentifier(),
				name: title ?? '',
				offX: Math.round(offsetX?.emu ?? 0),
				offY: Math.round(offsetY?.emu ?? 0),
				extX: Math.round(width.emu),
				extY: Math.round(height.emu),
				fillColor: fill?.type === 'solid' ? fill.color : null,
				lIns: bodyPr?.lIns?.emu ?? null,
				rIns: bodyPr?.rIns?.emu ?? null,
				tIns: bodyPr?.tIns?.emu ?? null,
				bIns: bodyPr?.bIns?.emu ?? null,
				anchor: bodyPr?.anchor ?? null,
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	public static override matchesNode(node: Node): boolean {
		return node.nodeName === 'wps:wsp';
	}

	public static override fromNode(
		node: Node,
		context: ComponentContext
	): TextBox {
		const title = evaluateXPathToString(
			`./${QNS.wps}cNvPr/@name/string()`,
			node
		);
		const xfrm = evaluateXPathToFirstNode(
			`./${QNS.wps}spPr/${QNS.a}xfrm`,
			node
		);
		if (!xfrm) {
			throw new Error('Failed to load TextBox. Missing a:xfrm.');
		}
		const width = {
			emu: evaluateXPathToNumber(`./${QNS.a}ext/@cx/number()`, xfrm),
		} as Length;
		const height = {
			emu: evaluateXPathToNumber(`./${QNS.a}ext/@cy/number()`, xfrm),
		} as Length;
		const offsetX = {
			emu: evaluateXPathToNumber(`./${QNS.a}off/@x/number()`, xfrm),
		} as Length;
		const offsetY = {
			emu: evaluateXPathToNumber(`./${QNS.a}off/@y/number()`, xfrm),
		} as Length;

		// Optional fill
		const fillColor = evaluateXPathToString(
			`./${QNS.wps}spPr/${QNS.a}solidFill/${QNS.a}srgbClr/@val/string()`,
			node
		);
		const fill =
			fillColor !== ''
				? ({
						type: 'solid',
						color: fillColor,
				  } as const)
				: undefined;

		// Optional bodyPr (only provided attributes)
		const lIns = evaluateXPathToNumber(
			`./${QNS.wps}bodyPr/@lIns/number()`,
			node
		);
		const rIns = evaluateXPathToNumber(
			`./${QNS.wps}bodyPr/@rIns/number()`,
			node
		);
		const tIns = evaluateXPathToNumber(
			`./${QNS.wps}bodyPr/@tIns/number()`,
			node
		);
		const bIns = evaluateXPathToNumber(
			`./${QNS.wps}bodyPr/@bIns/number()`,
			node
		);
		const anchor = evaluateXPathToString(
			`./${QNS.wps}bodyPr/@anchor/string()`,
			node
		);
		const hasAnyBodyPr =
			!Number.isNaN(lIns) ||
			!Number.isNaN(rIns) ||
			!Number.isNaN(tIns) ||
			!Number.isNaN(bIns) ||
			anchor !== '';
		const bodyPr = hasAnyBodyPr
			? {
					...(Number.isNaN(lIns)
						? {}
						: { lIns: { emu: lIns } as Length }),
					...(Number.isNaN(rIns)
						? {}
						: { rIns: { emu: rIns } as Length }),
					...(Number.isNaN(tIns)
						? {}
						: { tIns: { emu: tIns } as Length }),
					...(Number.isNaN(bIns)
						? {}
						: { bIns: { emu: bIns } as Length }),
					...(anchor === ''
						? {}
						: {
								anchor: anchor as
									| 'top'
									| 'ctr'
									| 'bot'
									| 'just'
									| 'dist',
						  }),
			  }
			: undefined;

		const contentNode = evaluateXPathToFirstNode(
			`./${QNS.wps}txbx/${QNS.w}txbxContent`,
			node
		);
		if (!contentNode) {
			throw new Error('Failed to load TextBox. Missing txbxContent.');
		}
		const childrenNodes = Array.from(
			contentNode.childNodes
		) as unknown as Node[];
		const children = createChildComponentsFromNodes<TextBoxChild>(
			this.children,
			childrenNodes,
			context
		);

		return new TextBox(
			{
				title,
				width,
				height,
				offsetX,
				offsetY,
				fill,
				bodyPr,
			},
			...children
		);
	}
}

registerComponent(TextBox);
