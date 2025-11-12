import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { NamespaceUri, QNS } from '../../../utilities/src/namespaces.ts';
import { evaluateXPathToFirstNode } from '../../../utilities/src/xquery.ts';
import { Picture } from './Picture.ts';

export type GraphicDataChild = Picture;

export type GraphicDataProps = Record<string, never>;

export class GraphicData extends Component<
	GraphicDataProps,
	GraphicDataChild
> {
	public static override readonly children: string[] = ['Picture'];
	public static override readonly mixed: boolean = false;

	public override async toNode(
		ancestry: ComponentAncestor[]
	): Promise<Node> {
		return create(
			`
				element ${QNS.a}graphic {
					element ${QNS.a}graphicData {
						attribute uri { "${NamespaceUri.pic}" },
						$children
					}
				}
			`,
			{
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'a:graphic';
	}

	static override fromNode(
		node: Node,
		context: ComponentContext
	): GraphicData {
		const graphicDataNode = evaluateXPathToFirstNode(
			`./${QNS.a}graphicData`,
			node
		);
		if (!graphicDataNode) {
			throw new Error(
				'Failed to load GraphicData. Missing a:graphicData child.'
			);
		}
		const picNode = evaluateXPathToFirstNode(
			`./${QNS.pic}pic`,
			graphicDataNode
		);
		if (!picNode) {
			throw new Error(
				'Failed to load GraphicData. Missing pic:pic child.'
			);
		}

		const picture = Picture.fromNode(picNode, context) as Picture;

		return new GraphicData({}, picture);
	}
}

registerComponent(GraphicData);


