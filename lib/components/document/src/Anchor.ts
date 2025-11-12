import { GraphicData } from './GraphicData.ts';

import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { createUniqueNumericIdentifier } from '../../../utilities/src/identifiers.ts';
import { type Length } from '../../../utilities/src/length.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import {
	evaluateXPathToFirstNode,
	evaluateXPathToMap,
	evaluateXPathToNumber,
	evaluateXPathToString,
} from '../../../utilities/src/xquery.ts';

export type AnchorChild = GraphicData;

export type AnchorWrapSquare = {
	type: 'square';
	wrapText: 'bothSides' | 'left' | 'right' | 'largest';
	dist: {
		left: Length;
		right: Length;
		top: Length;
		bottom: Length;
	};
};

export type AnchorPosition = {
	relativeFrom: 'column' | 'page' | 'margin' | 'paragraph';
	offset: Length;
};

export type AnchorProps = {
	width: Length;
	height: Length;
	title?: null | string;
	alt?: null | string;
	positionH: AnchorPosition;
	positionV: AnchorPosition;
	simplePos?: { x: number; y: number };
	wrap: AnchorWrapSquare;
	// Optional advanced flags, defaulting to schema-safe values
	allowOverlap?: boolean;
	behindDoc?: boolean;
	hidden?: boolean;
	layoutInCell?: boolean;
	locked?: boolean;
	relativeHeight?: number;
};

export class Anchor extends Component<AnchorProps, AnchorChild> {
	public static override readonly children: string[] = ['GraphicData'];
	public static override readonly mixed: boolean = false;

	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		const {
			title,
			alt,
			width,
			height,
			positionH,
			positionV,
			simplePos,
			wrap,
			allowOverlap = true,
			behindDoc = false,
			hidden = false,
			layoutInCell = true,
			locked = false,
			relativeHeight = 0,
		} = this.props;

		return create(
			`
				element ${QNS.wp}anchor {
					attribute allowOverlap { $allowOverlap },
					attribute behindDoc { $behindDoc },
					attribute distB { $distB },
					attribute distT { $distT },
					attribute distL { $distL },
					attribute distR { $distR },
					attribute hidden { $hidden },
					attribute layoutInCell { $layoutInCell },
					attribute locked { $locked },
					attribute relativeHeight { $relativeHeight },
					attribute simplePos { $simplePosAttr },

					element ${QNS.wp}simplePos {
						attribute x { $simplePosX },
						attribute y { $simplePosY }
					},
					element ${QNS.wp}positionH {
						attribute relativeFrom { $posHRelativeFrom },
						element ${QNS.wp}posOffset { $posHOffset }
					},
					element ${QNS.wp}positionV {
						attribute relativeFrom { $posVRelativeFrom },
						element ${QNS.wp}posOffset { $posVOffset }
					},
					element ${QNS.wp}extent {
						attribute cx { $width },
						attribute cy { $height }
					},
					element ${QNS.wp}effectExtent {
						attribute b { "0" },
						attribute l { "0" },
						attribute r { "0" },
						attribute t { "0" }
					},
					element ${QNS.wp}wrapSquare {
						attribute wrapText { $wrapText },
						attribute distB { $distB },
						attribute distT { $distT },
						attribute distL { $distL },
						attribute distR { $distR }
					},
					element ${QNS.wp}docPr {
						attribute id { $identifier },
						attribute name { $name },
						attribute descr { $desc }
					},
					$children
				}
			`,
			{
				identifier: createUniqueNumericIdentifier(),
				name: title || '',
				desc: alt || '',
				width: Math.round(width.emu),
				height: Math.round(height.emu),
				posHRelativeFrom: positionH.relativeFrom,
				posHOffset: Math.round(positionH.offset.emu),
				posVRelativeFrom: positionV.relativeFrom,
				posVOffset: Math.round(positionV.offset.emu),
				wrapText: wrap.wrapText,
				distL: Math.round(wrap.dist.left.emu),
				distR: Math.round(wrap.dist.right.emu),
				distT: Math.round(wrap.dist.top.emu),
				distB: Math.round(wrap.dist.bottom.emu),
				allowOverlap: allowOverlap ? '1' : '0',
				behindDoc: behindDoc ? '1' : '0',
				hidden: hidden ? '1' : '0',
				layoutInCell: layoutInCell ? '1' : '0',
				locked: locked ? '1' : '0',
				relativeHeight: `${relativeHeight}`,
				simplePosAttr: simplePos ? '1' : '0',
				simplePosX: `${simplePos?.x || 0}`,
				simplePosY: `${simplePos?.y || 0}`,
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'wp:anchor';
	}

	static override fromNode(node: Node, context: ComponentContext): Anchor {
		const {
			wrapText,
			distL,
			distR,
			distT,
			distB,
			width,
			height,
			title,
			alt,
			posHRelativeFrom,
			posHOffset,
			posVRelativeFrom,
			posVOffset,
			simplePosX,
			simplePosY,
		} = evaluateXPathToMap<{
			wrapText: string;
			distL: number;
			distR: number;
			distT: number;
			distB: number;
			width: number;
			height: number;
			title: string;
			alt: string;
			posHRelativeFrom: string;
			posHOffset: number;
			posVRelativeFrom: string;
			posVOffset: number;
			simplePosX: number;
			simplePosY: number;
		}>(
			`
				map {
					"wrapText": string(.//${QNS.wp}wrapSquare/@wrapText),
					"distL": number(.//${QNS.wp}wrapSquare/@distL),
					"distR": number(.//${QNS.wp}wrapSquare/@distR),
					"distT": number(.//${QNS.wp}wrapSquare/@distT),
					"distB": number(.//${QNS.wp}wrapSquare/@distB),
					"width": number(.//${QNS.wp}extent/@cx),
					"height": number(.//${QNS.wp}extent/@cy),
					"title": string(.//${QNS.wp}docPr/@name),
					"alt": string(.//${QNS.wp}docPr/@descr),
					"posHRelativeFrom": string(.//${QNS.wp}positionH/@relativeFrom),
					"posHOffset": number(.//${QNS.wp}positionH/${QNS.wp}posOffset),
					"posVRelativeFrom": string(.//${QNS.wp}positionV/@relativeFrom),
					"posVOffset": number(.//${QNS.wp}positionV/${QNS.wp}posOffset),
					"simplePosX": number(.//${QNS.wp}simplePos/@x),
					"simplePosY": number(.//${QNS.wp}simplePos/@y)
				}
			`,
			node
		);

		const graphicNode = evaluateXPathToFirstNode(`./${QNS.a}graphic`, node);
		if (!graphicNode) {
			throw new Error('Failed to load Anchor. Missing a:graphic child.');
		}
		const graphicData = GraphicData.fromNode(graphicNode, context);

		return new Anchor(
			{
				title,
				alt,
				width: { emu: width } as Length,
				height: { emu: height } as Length,
				positionH: {
					relativeFrom:
						posHRelativeFrom as AnchorPosition['relativeFrom'],
					offset: { emu: posHOffset } as Length,
				},
				positionV: {
					relativeFrom:
						posVRelativeFrom as AnchorPosition['relativeFrom'],
					offset: { emu: posVOffset } as Length,
				},
				simplePos: { x: simplePosX || 0, y: simplePosY || 0 },
				wrap: {
					type: 'square',
					wrapText: wrapText as AnchorWrapSquare['wrapText'],
					dist: {
						left: { emu: distL } as Length,
						right: { emu: distR } as Length,
						top: { emu: distT } as Length,
						bottom: { emu: distB } as Length,
					},
				},
			},
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			graphicData!
		);
	}
}

registerComponent(Anchor);
