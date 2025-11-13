import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { NamespaceUri, QNS } from '../../../utilities/src/namespaces.ts';
import { evaluateXPathToFirstNode } from '../../../utilities/src/xquery.ts';
import { Group } from './Group.ts';
import { Picture } from './Picture.ts';

export type GraphicDataChild = Picture | Group;

export type GraphicDataProps = Record<string, never>;

export class GraphicData extends Component<GraphicDataProps, GraphicDataChild> {
	public static override readonly children: string[] = ['Picture', 'Group'];
	public static override readonly mixed: boolean = false;

	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		const firstChild = this.children[0];
		const uri =
			firstChild &&
			typeof firstChild !== 'string' &&
			firstChild instanceof Picture
				? NamespaceUri.pic
				: NamespaceUri.wpg;
		return create(
			`
				element ${QNS.a}graphic {
					element ${QNS.a}graphicData {
						attribute uri { "${uri}" },
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

		const picOrWpgNode = evaluateXPathToFirstNode(
			`./${QNS.pic}pic | ./${QNS.wpg}wgp`,
			graphicDataNode
		);
		if (!picOrWpgNode) {
			throw new Error(
				'Failed to load GraphicData. Support only for pic:pic or wpg:wgp children.'
			);
		}

		console.log('graphic data child name', picOrWpgNode.nodeName);

		if (picOrWpgNode.nodeName === `pic:pic` || picOrWpgNode.nodeName === 'pic') {
			const picture = Picture.fromNode(picOrWpgNode, context);
			return new GraphicData({}, picture);
		} else if (picOrWpgNode.nodeName === 'wpg:wgp' || picOrWpgNode.nodeName === 'wgp') {
			const group = Group.fromNode(picOrWpgNode, context);
			return new GraphicData({}, group);
		} else {
			// Shoould not reach here, but just in case
			throw new Error(
				'Failed to load GraphicData. Support only for pic:pic or wpg:wgp children.'
			);
		}
	}
}

registerComponent(GraphicData);
