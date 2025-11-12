import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import {
	createUniqueNumericIdentifier,
} from '../../../utilities/src/identifiers.ts';
import { type Length, emu } from '../../../utilities/src/length.ts';
import {
	evaluateXPathToFirstNode,
	evaluateXPathToNumber,
	evaluateXPathToString,
} from '../../../utilities/src/xquery.ts';
import { GraphicData } from './GraphicData.ts';

export type InlineChild = GraphicData;

export type InlineProps = {
	width: Length;
	height: Length;
	title?: null | string;
	alt?: null | string;
};

export class Inline extends Component<InlineProps, InlineChild> {
	public static override readonly children: string[] = ['GraphicData'];
	public static override readonly mixed: boolean = false;

	public override async toNode(
		ancestry: ComponentAncestor[]
	): Promise<Node> {
		return create(
			`
				element ${QNS.wp}inline {
					element ${QNS.wp}extent {
						attribute cx { $width },
						attribute cy { $height }
					},
					element ${QNS.wp}docPr {
						attribute id { $identifier },
						attribute name { $name },
						attribute descr { $desc }
					},
					element ${QNS.wp}cNvGraphicFramePr {
						element ${QNS.a}graphicFrameLocks {
							attribute noChangeAspect { "1" }
						}
					},
					$children
				}
			`,
			{
				identifier: createUniqueNumericIdentifier(),
				width: Math.round(this.props.width.emu),
				height: Math.round(this.props.height.emu),
				name: this.props.title || '',
				desc: this.props.alt || '',
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'wp:inline';
	}

	static override fromNode(node: Node, context: ComponentContext): Inline {
		const title = evaluateXPathToString(
			`./${QNS.wp}docPr/@name/string()`,
			node
		);
		const alt = evaluateXPathToString(
			`./${QNS.wp}docPr/@descr/string()`,
			node
		);
		const width = emu(
			evaluateXPathToNumber(`./${QNS.wp}extent/@cx/number()`, node)
		);
		const height = emu(
			evaluateXPathToNumber(`./${QNS.wp}extent/@cy/number()`, node)
		);

		const graphicNode = evaluateXPathToFirstNode(
			`./${QNS.a}graphic`,
			node
		);
		if (!graphicNode) {
			throw new Error(
				'Failed to load Inline. Missing a:graphic child.'
			);
		}

		const graphicData = GraphicData.fromNode(
			graphicNode,
			context
		) as GraphicData;

		return new Inline(
			{
				title,
				alt,
				width,
				height,
			},
			graphicData
		);
	}
}

registerComponent(Inline);


