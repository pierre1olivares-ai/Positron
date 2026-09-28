import assert from 'node:assert/strict';
import test from 'node:test';
import { changedIssueFields } from '../src/webparts/qstarIssueManager/domain/issueDraft';
import { normalizeRegion, REGIONS } from '../src/webparts/qstarIssueManager/domain/referenceData';
import type { IIssue } from '../src/webparts/qstarIssueManager/models/IIssue';

test('draft updates include only edited fields and exclude server metadata', () => {
  const baseline = { id: 4, qsNumber: 1004, eTag: '"1"', followUp: 'Old', status: 'In Progress', progressLog: [] } as unknown as IIssue;
  const draft = { ...baseline, followUp: 'New', eTag: '"2"', qsNumber: 99, progressLog: [{ ts: '', author: '', text: 'Unrelated' }] };
  assert.deepEqual(changedIssueFields(baseline, draft), { followUp: 'New' });
});

test('email-only reassignment invalidates the previous person lookup', () => {
  const baseline = { taskOwner: 'Alex', taskOwnerEmail: 'first@example.com', taskOwnerId: 3 } as IIssue;
  const patch = changedIssueFields(baseline, { ...baseline, taskOwnerEmail: 'second@example.com' });
  assert.deepEqual(patch, { taskOwner: 'Alex', taskOwnerEmail: 'second@example.com', taskOwnerId: undefined });
  assert.deepEqual(changedIssueFields(baseline, { ...baseline, taskOwner: '', taskOwnerEmail: '' }), {
    taskOwner: '', taskOwnerEmail: '', taskOwnerId: 0,
  });
});

test('legacy regions map to current choices without inventing replacements for custom values', () => {
  assert.equal(normalizeRegion('Germany'), 'Western Europe (Amsterdam)');
  assert.equal(normalizeRegion('Western Europe'), 'Western Europe (Amsterdam)');
  assert.equal(normalizeRegion('Custom region'), 'Custom region');
  for (const region of REGIONS) assert.equal(normalizeRegion(region), region);
});
