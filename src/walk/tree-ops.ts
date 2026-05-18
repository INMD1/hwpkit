import type { DocRoot, GridNode } from '../model/doc-tree';
import { walkNode } from './TreeWalker';

export function countNodes(root: DocRoot): Record<string, number> {
  const counts: Record<string, number> = {};
  walkNode(root, (n) => { counts[n.tag] = (counts[n.tag] ?? 0) + 1; });
  return counts;
}

export function validateRoot(root: DocRoot): string[] {
  const errors: string[] = [];
  if (root.tag !== 'root') errors.push('Root node must have tag "root"');
  if (!Array.isArray(root.kids)) errors.push('Root.kids must be an array');
  if (root.kids.length === 0) errors.push('Document has no sheets');

  walkNode(root, (n) => {
    if (n.tag === 'cell' && n.kids.length === 0) {
      errors.push('CellNode must have at least one ParaNode child');
    }
    if (n.tag === 'grid' && (n as GridNode).kids.length === 0) {
      errors.push('GridNode must have at least one RowNode');
    }
  });

  return errors;
}
