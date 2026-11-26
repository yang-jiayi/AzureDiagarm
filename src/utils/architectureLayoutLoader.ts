import { prepareElkLayout } from './elkLayoutRuntime';

export class LayoutPreparationError extends Error {
  constructor(readonly cause: unknown) {
    super('Layout files could not be prepared. Save your work and reload the application before trying again.');
    this.name = 'LayoutPreparationError';
  }
}

export async function prepareArchitectureLayout(engine: 'elk' | 'dagre') {
  try {
    if (engine === 'elk') {
      const adapter = await import('./elkLayoutEngine');
      await prepareElkLayout();
      return adapter;
    }
    return await import('./layoutEngine');
  } catch (cause) {
    throw new LayoutPreparationError(cause);
  }
}
