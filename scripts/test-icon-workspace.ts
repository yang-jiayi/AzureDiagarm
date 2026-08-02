import assert from 'node:assert/strict';
import {
  createIconCollection,
  deleteIconCollection,
  normalizeIconWorkspace,
  recordRecentIcon,
  toggleFavoriteIcon,
  toggleIconInCollection,
} from '../src/services/iconWorkspaceService';

const normalized = normalizeIconWorkspace({
  favoriteIds: ['storage/account', 'storage/account', '', 42],
  recentIds: Array.from({ length: 30 }, (_, index) => `icon-${index}`),
  collections: [
    {
      id: 'data',
      name: ' Data services ',
      iconIds: ['sql/database', 'sql/database', 'storage/account'],
      createdAt: '2026-01-01T00:00:00.000Z',
      updatedAt: '2026-01-01T00:00:00.000Z',
    },
    { id: 'invalid', name: '   ', iconIds: [] },
  ],
});

assert.deepEqual(normalized.favoriteIds, ['storage/account']);
assert.equal(normalized.recentIds.length, 24);
assert.equal(normalized.collections.length, 1);
assert.equal(normalized.collections[0].name, 'Data services');
assert.deepEqual(normalized.collections[0].iconIds, ['sql/database', 'storage/account']);

const favorited = toggleFavoriteIcon(normalized, 'sql/database');
assert.deepEqual(favorited.favoriteIds, ['sql/database', 'storage/account']);
assert.deepEqual(toggleFavoriteIcon(favorited, 'sql/database').favoriteIds, ['storage/account']);

let recentState = normalized;
recentState = recordRecentIcon(recentState, 'storage/account');
recentState = recordRecentIcon(recentState, 'sql/database');
recentState = recordRecentIcon(recentState, 'storage/account');
assert.equal(recentState.recentIds[0], 'storage/account');
assert.equal(recentState.recentIds[1], 'sql/database');
assert.equal(new Set(recentState.recentIds).size, recentState.recentIds.length);

const created = createIconCollection(normalized, 'Platform');
assert.ok(created.collection);
assert.equal(created.collection?.name, 'Platform');
assert.equal(createIconCollection(created.state, 'platform').collection, null);

const withIcon = toggleIconInCollection(
  created.state,
  created.collection!.id,
  'compute/virtual-machine',
);
assert.deepEqual(
  withIcon.collections.find((collection) => collection.id === created.collection!.id)?.iconIds,
  ['compute/virtual-machine'],
);
assert.equal(
  deleteIconCollection(withIcon, created.collection!.id).collections.some(
    (collection) => collection.id === created.collection!.id,
  ),
  false,
);

console.log('[test:icon-workspace] persistence normalization and workspace actions verified');
