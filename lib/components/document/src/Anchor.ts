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

export type AnchorWrap =
	| {
			type: 'none';
	  }
	| {
			type: 'square';
			wrapText: 'bothSides' | 'left' | 'right' | 'largest';
			dist: {
				left: Length;
				right: Length;
				top: Length;
				bottom: Length;
			};
	  }
	| {
			type: 'tight';
			wrapText: 'bothSides' | 'left' | 'right' | 'largest';
			dist: {
				left: Length;
				right: Length;
				top: Length;
				bottom: Length;
			};
	  }
	| {
			type: 'through';
			wrapText: 'bothSides' | 'left' | 'right' | 'largest';
			dist: {
				left: Length;
				right: Length;
				top: Length;
				bottom: Length;
			};
	  }
	| {
			type: 'topAndBottom';
			dist: {
				top: Length;
				bottom: Length;
			};
	  };

export type AnchorPosition =
	| {
			mode: 'offset';
			relativeFrom: 'column' | 'page' | 'margin' | 'paragraph';
			offset: Length;
	  }
	| {
			mode: 'align';
			relativeFrom: 'column' | 'page' | 'margin' | 'paragraph';
			align:
				| 'left'
				| 'center'
				| 'right'
				| 'inside'
				| 'outside'
				| 'top'
				| 'bottom';
	  };

export type AnchorProps = {
	width: Length;
	height: Length;
	title?: null | string;
	alt?: null | string;
	positionH: AnchorPosition;
	positionV: AnchorPosition;
	simplePos?: { x: number; y: number };
	wrap: AnchorWrap;
	/**
	 * Floating margins on the anchor element itself (distinct from wrap distances).
	 * If provided, only the given sides are emitted as dist* attributes on wp:anchor.
	 */
	floating?: {
		marginLeft?: Length;
		marginRight?: Length;
		marginTop?: Length;
		marginBottom?: Length;
	};
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
			floating,
			relativeHeight = 0,
			allowOverlap = true,
			behindDoc = false,
			hidden = false,
			layoutInCell = true,
			locked = false,
		} = this.props;

		return create(
			`
				element ${QNS.wp}anchor {
					attribute allowOverlap { $allowOverlap },
					attribute behindDoc { $behindDoc },
					if (exists($marginB)) then attribute distB { $marginB } else (),
					if (exists($marginT)) then attribute distT { $marginT } else (),
					if (exists($marginL)) then attribute distL { $marginL } else (),
					if (exists($marginR)) then attribute distR { $marginR } else (),
					attribute hidden { $hidden },
					attribute layoutInCell { $layoutInCell },
					attribute locked { $locked },
          if (exists($relativeHeight)) then attribute relativeHeight { $relativeHeight } else (),
					attribute simplePos { $simplePosAttr },

					element ${QNS.wp}simplePos {
						attribute x { $simplePosX },
						attribute y { $simplePosY }
					},
					if (exists($hasPosH) and $hasPosH) then element ${QNS.wp}positionH {
						attribute relativeFrom { $posHRelativeFrom },
						if (exists($posHAlign)) then element ${QNS.wp}align { $posHAlign } else (),
						if (exists($posHOffset)) then element ${QNS.wp}posOffset { $posHOffset } else ()
					} else (),
					if (exists($hasPosV) and $hasPosV) then element ${QNS.wp}positionV {
						attribute relativeFrom { $posVRelativeFrom },
						if (exists($posVAlign)) then element ${QNS.wp}align { $posVAlign } else (),
						if (exists($posVOffset)) then element ${QNS.wp}posOffset { $posVOffset } else ()
					} else (),
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
					if ($isWrapNone) then element ${QNS.wp}wrapNone {} else (),
					if ($isWrapSquare) then element ${QNS.wp}wrapSquare {
						if (exists($wrapText)) then attribute wrapText { $wrapText } else (),
						if (exists($distB)) then attribute distB { $distB } else (),
						if (exists($distT)) then attribute distT { $distT } else (),
						if (exists($distL)) then attribute distL { $distL } else (),
						if (exists($distR)) then attribute distR { $distR } else ()
					} else (),
					if ($isWrapTight) then element ${QNS.wp}wrapTight {
						if (exists($wrapText)) then attribute wrapText { $wrapText } else (),
						if (exists($distB)) then attribute distB { $distB } else (),
						if (exists($distT)) then attribute distT { $distT } else (),
						if (exists($distL)) then attribute distL { $distL } else (),
						if (exists($distR)) then attribute distR { $distR } else ()
					} else (),
					if ($isWrapThrough) then element ${QNS.wp}wrapThrough {
						if (exists($wrapText)) then attribute wrapText { $wrapText } else (),
						if (exists($distB)) then attribute distB { $distB } else (),
						if (exists($distT)) then attribute distT { $distT } else (),
						if (exists($distL)) then attribute distL { $distL } else (),
						if (exists($distR)) then attribute distR { $distR } else ()
					} else (),
					if ($isWrapTopAndBottom) then element ${QNS.wp}wrapTopAndBottom {
						if (exists($distB)) then attribute distB { $distB } else (),
						if (exists($distT)) then attribute distT { $distT } else ()
					} else (),
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
				hasPosH: !!positionH,
				hasPosV: !!positionV,
				posHRelativeFrom: positionH?.relativeFrom ?? null,
				posVRelativeFrom: positionV?.relativeFrom ?? null,
				posHAlign:
					positionH && positionH.mode === 'align'
						? (positionH.align as string)
						: null,
				posVAlign:
					positionV && positionV.mode === 'align'
						? (positionV.align as string)
						: null,
				posHOffset:
					positionH && positionH.mode === 'offset'
						? Math.round(positionH.offset.emu)
						: null,
				posVOffset:
					positionV && positionV.mode === 'offset'
						? Math.round(positionV.offset.emu)
						: null,
				marginL: floating?.marginLeft?.emu ?? null,
				marginR: floating?.marginRight?.emu ?? null,
				marginT: floating?.marginTop?.emu ?? null,
				marginB: floating?.marginBottom?.emu ?? null,
				isWrapNone: wrap?.type === 'none' || null,
				isWrapSquare: wrap?.type === 'square' || null,
				isWrapTight: wrap?.type === 'tight' || null,
				isWrapThrough: wrap?.type === 'through' || null,
				isWrapTopAndBottom: wrap?.type === 'topAndBottom' || null,
				wrapText:
					wrap &&
					(wrap.type === 'square' ||
						wrap.type === 'tight' ||
						wrap.type === 'through')
						? wrap.wrapText
						: null,
				distL:
					wrap &&
					(wrap.type === 'square' ||
						wrap.type === 'tight' ||
						wrap.type === 'through')
						? Math.round(wrap.dist.left.emu)
						: null,
				distR:
					wrap &&
					(wrap.type === 'square' ||
						wrap.type === 'tight' ||
						wrap.type === 'through')
						? Math.round(wrap.dist.right.emu)
						: null,
				distT:
					wrap &&
					(wrap.type === 'square' ||
						wrap.type === 'tight' ||
						wrap.type === 'through')
						? Math.round(wrap.dist.top.emu)
						: wrap && wrap.type === 'topAndBottom'
						? Math.round(wrap.dist.top.emu)
						: null,
				distB:
					wrap &&
					(wrap.type === 'square' ||
						wrap.type === 'tight' ||
						wrap.type === 'through')
						? Math.round(wrap.dist.bottom.emu)
						: wrap && wrap.type === 'topAndBottom'
						? Math.round(wrap.dist.bottom.emu)
						: null,
				allowOverlap: allowOverlap ? '1' : '0',
				behindDoc: behindDoc ? '1' : '0',
				hidden: hidden ? '1' : '0',
				layoutInCell: layoutInCell ? '1' : '0',
				locked: locked ? '1' : '0',
				relativeHeight: relativeHeight ?? 0,
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
			width,
			height,
			title,
			alt,
			posHRelativeFrom,
			posHAlign,
			posHOffset,
			posVRelativeFrom,
			posVAlign,
			posVOffset,
			simplePosX,
			simplePosY,
		} = evaluateXPathToMap<{
			width: number;
			height: number;
			title: string;
			alt: string;
			posHRelativeFrom: string;
			posHAlign?: string;
			posHOffset?: number;
			posVRelativeFrom: string;
			posVAlign?: string;
			posVOffset?: number;
			simplePosX: number;
			simplePosY: number;
		}>(
			`
				map {
					"width": number(.//${QNS.wp}extent/@cx),
					"height": number(.//${QNS.wp}extent/@cy),
					"title": string(.//${QNS.wp}docPr/@name),
					"alt": string(.//${QNS.wp}docPr/@descr),
					"posHRelativeFrom": string(.//${QNS.wp}positionH/@relativeFrom),
					"posHAlign": string(.//${QNS.wp}positionH/${QNS.wp}align),
					"posHOffset": number(.//${QNS.wp}positionH/${QNS.wp}posOffset),
					"posVRelativeFrom": string(.//${QNS.wp}positionV/@relativeFrom),
					"posVAlign": string(.//${QNS.wp}positionV/${QNS.wp}align),
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

		// Floating margins on anchor element
		const mL = evaluateXPathToNumber(`@distL/number()`, node);
		const mR = evaluateXPathToNumber(`@distR/number()`, node);
		const mT = evaluateXPathToNumber(`@distT/number()`, node);
		const mB = evaluateXPathToNumber(`@distB/number()`, node);
		const hasFloating =
			!Number.isNaN(mL) ||
			!Number.isNaN(mR) ||
			!Number.isNaN(mT) ||
			!Number.isNaN(mB);
		const floating = hasFloating
			? {
					...(Number.isNaN(mL)
						? {}
						: { marginLeft: { emu: mL } as Length }),
					...(Number.isNaN(mR)
						? {}
						: { marginRight: { emu: mR } as Length }),
					...(Number.isNaN(mT)
						? {}
						: { marginTop: { emu: mT } as Length }),
					...(Number.isNaN(mB)
						? {}
						: { marginBottom: { emu: mB } as Length }),
			  }
			: undefined;

		// Determine wrap
		let wrap: AnchorWrap | undefined;
		const wrapNone = evaluateXPathToFirstNode(`./${QNS.wp}wrapNone`, node);
		const wrapSquare = evaluateXPathToFirstNode(
			`./${QNS.wp}wrapSquare`,
			node
		) as Element | null;
		const wrapTight = evaluateXPathToFirstNode(
			`./${QNS.wp}wrapTight`,
			node
		) as Element | null;
		const wrapThrough = evaluateXPathToFirstNode(
			`./${QNS.wp}wrapThrough`,
			node
		) as Element | null;
		const wrapTopBottom = evaluateXPathToFirstNode(
			`./${QNS.wp}wrapTopAndBottom`,
			node
		) as Element | null;
		if (wrapNone) {
			wrap = { type: 'none' };
		} else if (wrapSquare) {
			const wt = wrapSquare.getAttribute('wrapText') || 'bothSides';
			const toNum = (a: string | null) => (a ? Number(a) : NaN);
			const l = toNum(wrapSquare.getAttribute('distL'));
			const r = toNum(wrapSquare.getAttribute('distR'));
			const t = toNum(wrapSquare.getAttribute('distT'));
			const b = toNum(wrapSquare.getAttribute('distB'));
			wrap = {
				type: 'square',
				wrapText: wt as AnchorWrap extends infer T
					? T extends { type: 'square'; wrapText: infer W }
						? W
						: never
					: never,
				dist: {
					left: { emu: Number.isNaN(l) ? 0 : l } as Length,
					right: { emu: Number.isNaN(r) ? 0 : r } as Length,
					top: { emu: Number.isNaN(t) ? 0 : t } as Length,
					bottom: { emu: Number.isNaN(b) ? 0 : b } as Length,
				},
			};
		} else if (wrapTight) {
			const wt = wrapTight.getAttribute('wrapText') || 'bothSides';
			const toNum = (a: string | null) => (a ? Number(a) : NaN);
			wrap = {
				type: 'tight',
				wrapText: wt as any,
				dist: {
					left: {
						emu: Number.isNaN(
							toNum(wrapTight.getAttribute('distL'))
						)
							? 0
							: toNum(wrapTight.getAttribute('distL')),
					} as Length,
					right: {
						emu: Number.isNaN(
							toNum(wrapTight.getAttribute('distR'))
						)
							? 0
							: toNum(wrapTight.getAttribute('distR')),
					} as Length,
					top: {
						emu: Number.isNaN(
							toNum(wrapTight.getAttribute('distT'))
						)
							? 0
							: toNum(wrapTight.getAttribute('distT')),
					} as Length,
					bottom: {
						emu: Number.isNaN(
							toNum(wrapTight.getAttribute('distB'))
						)
							? 0
							: toNum(wrapTight.getAttribute('distB')),
					} as Length,
				},
			};
		} else if (wrapThrough) {
			const wt = wrapThrough.getAttribute('wrapText') || 'bothSides';
			const toNum = (a: string | null) => (a ? Number(a) : NaN);
			wrap = {
				type: 'through',
				wrapText: wt as any,
				dist: {
					left: {
						emu: Number.isNaN(
							toNum(wrapThrough.getAttribute('distL'))
						)
							? 0
							: toNum(wrapThrough.getAttribute('distL')),
					} as Length,
					right: {
						emu: Number.isNaN(
							toNum(wrapThrough.getAttribute('distR'))
						)
							? 0
							: toNum(wrapThrough.getAttribute('distR')),
					} as Length,
					top: {
						emu: Number.isNaN(
							toNum(wrapThrough.getAttribute('distT'))
						)
							? 0
							: toNum(wrapThrough.getAttribute('distT')),
					} as Length,
					bottom: {
						emu: Number.isNaN(
							toNum(wrapThrough.getAttribute('distB'))
						)
							? 0
							: toNum(wrapThrough.getAttribute('distB')),
					} as Length,
				},
			};
		} else if (wrapTopBottom) {
			const toNum = (a: string | null) => (a ? Number(a) : NaN);
			const t = toNum(wrapTopBottom.getAttribute('distT'));
			const b = toNum(wrapTopBottom.getAttribute('distB'));
			wrap = {
				type: 'topAndBottom',
				dist: {
					top: { emu: Number.isNaN(t) ? 0 : t } as Length,
					bottom: { emu: Number.isNaN(b) ? 0 : b } as Length,
				},
			};
		}

		// Positions
		let positionH: AnchorPosition | undefined;
		let positionV: AnchorPosition | undefined;
		if (posHAlign && !Number.isNaN(posHOffset as number)) {
			throw new Error(
				'Invalid Anchor: positionH cannot contain both align and posOffset.'
			);
		}
		if (posVAlign && !Number.isNaN(posVOffset as number)) {
			throw new Error(
				'Invalid Anchor: positionV cannot contain both align and posOffset.'
			);
		}
		if (posHAlign) {
			positionH = {
				mode: 'align',
				relativeFrom: posHRelativeFrom as AnchorPosition extends infer T
					? T extends { relativeFrom: infer R }
						? R
						: never
					: never,
				align: posHAlign as any,
			};
		} else if (!Number.isNaN(posHOffset as number)) {
			positionH = {
				mode: 'offset',
				relativeFrom: posHRelativeFrom as any,
				offset: { emu: posHOffset as number } as Length,
			};
		}
		if (posVAlign) {
			positionV = {
				mode: 'align',
				relativeFrom: posVRelativeFrom as any,
				align: posVAlign as any,
			};
		} else if (!Number.isNaN(posVOffset as number)) {
			positionV = {
				mode: 'offset',
				relativeFrom: posVRelativeFrom as any,
				offset: { emu: posVOffset as number } as Length,
			};
		}

		return new Anchor(
			{
				title,
				alt,
				width: { emu: width } as Length,
				height: { emu: height } as Length,
				positionH,
				positionV,
				simplePos: { x: simplePosX || 0, y: simplePosY || 0 },
				floating,
				wrap,
			},
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			graphicData!
		);
	}
}

registerComponent(Anchor);
