import type { AnyNode, DocRoot } from '../model/doc-tree';

export type WalkCallback = (node: AnyNode, parent: AnyNode | null, depth: number) => void | 'stop';

export function walkNode(
  node: AnyNode,
  cb: WalkCallback,
  parent: AnyNode | null = null,
  depth = 0,
): boolean {
  const result = cb(node, parent, depth);
  if (result === 'stop') return false;

  if ('kids' in node && Array.isArray((node as any).kids)) {
    for (const kid of (node as any).kids) {
      if (!walkNode(kid as AnyNode, cb, node, depth + 1)) return false;
    }
  }
  return true;
}

export class TreeWalker {
  walk(root: DocRoot, cb: WalkCallback): void {
    walkNode(root, cb);
  }

  findAll<T extends AnyNode>(root: DocRoot, predicate: (n: AnyNode) => n is T): T[] {
    const results: T[] = [];
    walkNode(root, (n) => { if (predicate(n)) results.push(n); });
    return results;
  }

  extractText(root: DocRoot): string {
    const parts: string[] = [];
    walkNode(root, (n) => {
      if (n.tag === 'txt') parts.push(n.content);
      if (n.tag === 'br') parts.push('\n');
      if (n.tag === 'pb') parts.push('\n\n');
    });
    return parts.join('');
  }
}
