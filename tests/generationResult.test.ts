// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import { generationProducesCanvas } from '../src/utils/generationResult';

test('only topology generation enables canvas review and validation', () => {
  assert.equal(generationProducesCanvas('topology'), true);
  assert.equal(generationProducesCanvas('both'), true);
  assert.equal(generationProducesCanvas('blueprint'), false);
  assert.equal(generationProducesCanvas('reference'), false);
});
