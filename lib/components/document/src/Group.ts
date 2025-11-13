import { Picture } from './Picture.ts';
import { TextBox } from './TextBox.ts';

import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { type Length } from '../../../utilities/src/length.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import { evaluateXPathToFirstNode, evaluateXPathToNumber } from '../../../utilities/src/xquery.ts';

export type GroupChild = Picture | TextBox;

export type GroupProps = {
	width: Length;
	height: Length;
	offsetX?: Length;
	offsetY?: Length;
	childOffsetX?: Length;
	childOffsetY?: Length;
	childExtX?: Length;
	childExtY?: Length;
};

export class Group extends Component<GroupProps, GroupChild> {
	public static override readonly children: string[] = ['Picture', 'TextBox'];
	public static override readonly mixed: boolean = false;

	public override async toNode(
		ancestry: ComponentAncestor[]
	): Promise<Node> {
		const {
			width,
			height,
			offsetX,
			offsetY,
			childOffsetX,
			childOffsetY,
			childExtX,
			childExtY,
		} = this.props;
		return create(
			`
				element ${QNS.wpg}wgp {
					element ${QNS.wpg}cNvGrpSpPr {
						element ${QNS.a}grpSpLocks {
							attribute noChangeAspect { "1" }
						}
					},
					element ${QNS.wpg}grpSpPr {
						element ${QNS.a}xfrm {
							element ${QNS.a}off {
								attribute x { $offX },
								attribute y { $offY }
							},
							element ${QNS.a}ext {
								attribute cx { $extX },
								attribute cy { $extY }
							},
							element ${QNS.a}chOff {
								attribute x { $chOffX },
								attribute y { $chOffY }
							},
							element ${QNS.a}chExt {
								attribute cx { $chExtX },
								attribute cy { $chExtY }
							}
						}
					},
					$children
				}
			`,
			{
				offX: Math.round((offsetX?.emu ?? 0)),
				offY: Math.round((offsetY?.emu ?? 0)),
				extX: Math.round(width.emu),
				extY: Math.round(height.emu),
				chOffX: Math.round((childOffsetX?.emu ?? 0)),
				chOffY: Math.round((childOffsetY?.emu ?? 0)),
				chExtX: Math.round((childExtX?.emu ?? 0)),
				chExtY: Math.round((childExtY?.emu ?? 0)),
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	public static override matchesNode(node: Node): boolean {
		return node.nodeName === 'wpg:wgp';
	}

	public static override fromNode(
		node: Node,
		context: ComponentContext
	): Group {
		const xfrm = evaluateXPathToFirstNode(
			`./${QNS.wpg}grpSpPr/${QNS.a}xfrm`,
			node
		);
		if (!xfrm) {
			throw new Error('Failed to load Group. Missing a:xfrm.');
		}
		const width = { emu: evaluateXPathToNumber(`./${QNS.a}ext/@cx/number()`, xfrm) } as Length;
		const height = { emu: evaluateXPathToNumber(`./${QNS.a}ext/@cy/number()`, xfrm) } as Length;
		const offsetX = { emu: evaluateXPathToNumber(`./${QNS.a}off/@x/number()`, xfrm) } as Length;
		const offsetY = { emu: evaluateXPathToNumber(`./${QNS.a}off/@y/number()`, xfrm) } as Length;
		const childOffsetX = { emu: evaluateXPathToNumber(`./${QNS.a}chOff/@x/number()`, xfrm) } as Length;
		const childOffsetY = { emu: evaluateXPathToNumber(`./${QNS.a}chOff/@y/number()`, xfrm) } as Length;
		const childExtX = { emu: evaluateXPathToNumber(`./${QNS.a}chExt/@cx/number()`, xfrm) } as Length;
		const childExtY = { emu: evaluateXPathToNumber(`./${QNS.a}chExt/@cy/number()`, xfrm) } as Length;

		// Children: Picture and/or TextBox
		const children: Array<Picture | TextBox> = [];
		const picNode = evaluateXPathToFirstNode(`./${QNS.pic}pic`, node);
		if (picNode) {
			const pic = Picture.fromNode(picNode, context);
			if (pic) children.push(pic);
		}
		const tbNode = evaluateXPathToFirstNode(`./${QNS.wps}wsp`, node);
		if (tbNode) {
			const tb = TextBox.fromNode(tbNode, context);
			if (tb) children.push(tb);
		}

		return new Group(
			{
				width,
				height,
				offsetX,
				offsetY,
				childOffsetX,
				childOffsetY,
				childExtX,
				childExtY,
			},
			...children
		);
	}
}

registerComponent(Group);


