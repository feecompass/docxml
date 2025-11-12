import { Anchor } from './Anchor.ts';
import { Inline } from './Inline.ts';

import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import { evaluateXPathToFirstNode } from '../../../utilities/src/xquery.ts';
import type { Anchor as AnchorType } from './Anchor.ts';
import type { Inline as InlineType } from './Inline.ts';

export type DrawingChild = InlineType | AnchorType;

export type DrawingProps = Record<string, never>;

export class Drawing extends Component<DrawingProps, DrawingChild> {
	public static override readonly children: string[] = ['Inline', 'Anchor'];
	public static override readonly mixed: boolean = false;

	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		return create(
			`
				element ${QNS.w}drawing {
					$children
				}
			`,
			{
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:drawing';
	}

	static override fromNode(node: Node, context: ComponentContext): Drawing {
		const inlineNode = evaluateXPathToFirstNode(
			`./(${QNS.wp}inline)`,
			node
		);
		if (inlineNode) {
			const inline = Inline.fromNode(inlineNode, context) as InlineType;
			return new Drawing({}, inline);
		}
		const anchorNode = evaluateXPathToFirstNode(
			`./(${QNS.wp}anchor)`,
			node
		);
		if (anchorNode) {
			const anchor = Anchor.fromNode(anchorNode, context) as AnchorType;
			return new Drawing({}, anchor);
		}
		throw new Error(
			'Failed to load Drawing. Missing either wp:inline or wp:anchor child.'
		);
	}
}

registerComponent(Drawing);
