import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {mergeXlsxEdits} from './merge-edits.ts';
import {patchParts, partMatching} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

test('calculation settings and unknown workbook extensions survive and remain editable by a baseline', () => {
  const book = new Workbook();
  book.addWorksheet('Data');
  const add = (xml: string, count: number) =>
    xml.replace(
      '</workbook>',
      `<calcPr calcMode="manual" iterate="1" iterateCount="${count}"/><extLst><ext uri="urn:custom"><x:data xmlns:x="urn:custom">keep</x:data></ext></extLst></workbook>`,
    );
  const original = patchParts(writeXlsx(book), {'xl/workbook.xml': (xml) => add(xml, 100)});
  const roundtrip = writeXlsx(readXlsx(original));
  assert.match(partMatching(roundtrip, /^xl\/workbook.xml$/), /iterateCount="100"/);
  assert.match(partMatching(roundtrip, /^xl\/workbook.xml$/), /urn:custom/);
  const baseline = patchParts(writeXlsx(book), {'xl/workbook.xml': (xml) => add(xml, 100)});
  const edited = patchParts(baseline, {
    'xl/workbook.xml': (xml) => xml.replace('iterateCount="100"', 'iterateCount="200"'),
  });
  const out = mergeXlsxEdits(original, baseline, edited, {
    sheets: [{name: 'Data', originalName: 'Data', baselineName: 'Data'}],
  });
  assert.match(partMatching(out.data, /^xl\/workbook.xml$/), /iterateCount="200"/);
});
