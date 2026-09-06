import { discardDraft, readDraft, writeDraft, DraftConflictError, DRAFT_DATABASE, DRAFT_STORE } from '../../src/services/draftStorage';
import { createSnapshot, getVersion, getAllVersions, deleteVersion } from '../../src/services/versionStorageService';
import type { EditorDocument } from '../../src/services/editorHistory';
import { createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { flushSync } from 'react-dom';
import { useDraftAutosave } from '../../src/hooks/useDraftAutosave';

function check(value: unknown, message: string): asserts value {
  if (!value) throw new Error(message);
}

export async function runWorkspaceStorageFixtures() {
  const document: EditorDocument = {
    nodes: [{ id: 'storage-web', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Storage test' } }],
    edges: [], titleBlockData: { architectureName: 'Draft', author: '', date: '', version: '1' },
    architecturePrompt: 'Start', originalPrompt: 'Start', workflow: [],
    settings: { pricingMode: 'payg', stylePreset: 'detailed', edgeStyle: 'orthogonal' },
  };
  const key = crypto.randomUUID();
  check(await readDraft(key) === null, 'The fixture must start with a new key.');
  const first = await writeDraft(key, document, null);
  check(first.revision === 1 && (await readDraft(key))?.document.nodes.length === 1, 'Committed draft is readable.');
  const outcomes = await Promise.allSettled([
    writeDraft(key, { ...document, architecturePrompt: 'Tab A' }, 1),
    writeDraft(key, { ...document, architecturePrompt: 'Tab B' }, 1),
  ]);
  check(outcomes.filter(result => result.status === 'fulfilled').length === 1, 'Only one concurrent tab may update a revision.');
  check(outcomes.some(result => result.status === 'rejected' && result.reason instanceof DraftConflictError), 'Concurrent changes must not silently overwrite a draft.');
  const saved = await readDraft(key);
  check(saved?.revision === 2, 'The winning revision must be 2.');

  const originalPut = IDBObjectStore.prototype.put;
  IDBObjectStore.prototype.put = function (...args: Parameters<IDBObjectStore['put']>) {
    const request = originalPut.apply(this, args);
    if (this.transaction.db.name === DRAFT_DATABASE && this.name === DRAFT_STORE) {
      request.addEventListener('success', () => this.transaction.abort());
    }
    return request;
  };
  try {
    let failed = false;
    try { await writeDraft(key, document, 2); } catch { failed = true; }
    check(failed, 'Request success followed by transaction abort must reject autosave.');
    check((await readDraft(key))?.revision === 2, 'Aborted draft transactions must not overwrite committed content.');
  } finally {
    IDBObjectStore.prototype.put = originalPut;
  }
  const snapshotOptions = {
    settings: { ...document.settings, pricingRegion: 'eastus2' },
    titleBlockData: { ...document.titleBlockData },
    reviewHistory: [],
    validationSourceFingerprint: 'fixture-source',
  };
  const snapshotPending = createSnapshot(document.nodes, document.edges, 'Workspace fixture', snapshotOptions);
  snapshotOptions.settings.pricingRegion = 'japaneast';
  snapshotOptions.titleBlockData.architectureName = 'Changed during snapshot save';
  const version = await snapshotPending;
  const persistedVersion = await getVersion(version.versionId);
  check(persistedVersion?.nodes[0].id === 'storage-web', 'Version snapshots remain compatible.');
  check(persistedVersion.settings?.pricingRegion === 'eastus2'
    && persistedVersion.titleBlockData?.architectureName === 'Draft'
    && persistedVersion.validationSourceFingerprint === 'fixture-source',
  'Snapshot metadata must be an isolated, durable copy, including settings and review provenance.');
  const versionCount = (await getAllVersions()).length;
  IDBObjectStore.prototype.put = function (...args: Parameters<IDBObjectStore['put']>) {
    const request = originalPut.apply(this, args);
    if (this.transaction.db.name === 'AzureDiagramVersions') {
      request.addEventListener('success', () => this.transaction.abort());
    }
    return request;
  };
  try {
    let failed = false;
    try { await createSnapshot(document.nodes, document.edges, 'Aborted snapshot'); } catch { failed = true; }
    check(failed, 'A version snapshot must not report success before transaction commit.');
    check((await getAllVersions()).length === versionCount, 'Aborted snapshots must not appear in history.');
  } finally {
    IDBObjectStore.prototype.put = originalPut;
  }
  await deleteVersion(version.versionId);
  await discardDraft(key, 2);
  check(await readDraft(key) === null, 'Discarding a specific revision must commit.');

  const autosaveKey = `flush-${crypto.randomUUID()}`;
  const host = globalThis.document.createElement('div');
  globalThis.document.body.append(host);
  const root = createRoot(host);
  let autosave: ReturnType<typeof useDraftAutosave> | undefined;
  const Harness = ({ value }: { value: EditorDocument }) => {
    autosave = useDraftAutosave(value, () => undefined, { key: autosaveKey, enabled: true });
    return null;
  };
  const render = (value: EditorDocument) => flushSync(() => root.render(createElement(Harness, { value })));
  let flushRevisions = 0;
  try {
    render(document);
    const deadline = Date.now() + 5000;
    while (!autosave?.ready && Date.now() < deadline) await new Promise(resolve => setTimeout(resolve, 10));
    check(autosave?.ready, 'The autosave hook must be ready before flushing.');
    const firstEdit = { ...document, architecturePrompt: 'Before the write' };
    const latestEdit = { ...document, architecturePrompt: 'Edited while saving' };
    let editedDuringWrite = false;
    IDBObjectStore.prototype.put = function (...args: Parameters<IDBObjectStore['put']>) {
      const request = originalPut.apply(this, args);
      if (this.transaction.db.name === DRAFT_DATABASE && this.name === DRAFT_STORE && !editedDuringWrite) {
        editedDuringWrite = true;
        request.addEventListener('success', () => render(latestEdit), { once: true });
      }
      return request;
    };
    render(firstEdit);
    const committed = await autosave.flush();
    check(editedDuringWrite && committed?.revision === 2, 'Flush must drain edits arriving before an earlier transaction commits.');
    check(committed.document.architecturePrompt === latestEdit.architecturePrompt
      && (await readDraft(autosaveKey))?.document.architecturePrompt === latestEdit.architecturePrompt,
    'A completed flush must preserve the latest edit, not merely the first successful request.');
    flushRevisions = committed.revision;
  } finally {
    IDBObjectStore.prototype.put = originalPut;
    flushSync(() => root.unmount());
    host.remove();
    const remaining = await readDraft(autosaveKey);
    if (remaining) await discardDraft(autosaveKey, remaining.revision);
  }
  let startupSnapshots = 0;
  for (const delayedActivation of [true, false]) {
    const startupKey = `startup-${crypto.randomUUID()}`;
    const nextScopeKey = `next-${startupKey}`;
    const startupHost = globalThis.document.createElement('div');
    globalThis.document.body.append(startupHost);
    const startupRoot = createRoot(startupHost);
    let startupAutosave: ReturnType<typeof useDraftAutosave> | undefined;
    const StartupHarness = ({ value, enabled, scope }: { value: EditorDocument; enabled: boolean; scope: string }) => {
      startupAutosave = useDraftAutosave(value, () => undefined, { key: scope, enabled });
      return null;
    };
    const renderStartup = (value: EditorDocument, enabled: boolean, scope = startupKey) => flushSync(() => {
      startupRoot.render(createElement(StartupHarness, { value, enabled, scope }));
    });
    try {
      const initialScope = delayedActivation ? 'pending' : startupKey;
      renderStartup({ ...document, nodes: [] }, !delayedActivation, initialScope);
      renderStartup(document, !delayedActivation, initialScope);
      if (delayedActivation) renderStartup(document, true);
      const deadline = Date.now() + 5000;
      while (!startupAutosave?.ready && Date.now() < deadline) await new Promise(resolve => setTimeout(resolve, 10));
      check(startupAutosave?.ready, 'Startup storage must finish opening.');
      const committed = await startupAutosave.flush();
      check(committed?.document.nodes[0]?.id === 'storage-web',
        'A document restored before draft activation or its initial read must not become an unsaved baseline.');
      check((await readDraft(startupKey))?.document.nodes[0]?.id === 'storage-web',
        'Restored startup nodes must be durable regardless of identity and storage timing.');
      renderStartup(document, true, nextScopeKey);
      const scopeDeadline = Date.now() + 5000;
      while ((!startupAutosave.ready || startupAutosave.status !== 'idle') && Date.now() < scopeDeadline) {
        await new Promise(resolve => setTimeout(resolve, 10));
      }
      check(startupAutosave.ready && startupAutosave.status === 'idle', 'The new draft scope must finish opening.');
      check(await startupAutosave.flush() === null && await readDraft(nextScopeKey) === null,
        'Switching to another scope must not automatically copy the previous scope document.');
      startupSnapshots += 1;
    } finally {
      flushSync(() => startupRoot.unmount());
      startupHost.remove();
      const remaining = await readDraft(startupKey);
      if (remaining) await discardDraft(startupKey, remaining.revision);
      const nextScopeDraft = await readDraft(nextScopeKey);
      if (nextScopeDraft) await discardDraft(nextScopeKey, nextScopeDraft.revision);
    }
  }
  return { concurrentWrites: outcomes.length, draftRevisions: saved.revision, versionSaved: true, metadataCloned: true, abortHandled: true, flushRevisions, startupSnapshots };
}
