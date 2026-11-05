import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'combined-native-features-keep-their-owners',
  cluster: 'fidelity',
  provenance: {source: 'authored-regression', ref: 'combined-native-features/sample.xlsx'},
  description:
    'Repeated saves retain a chart, shape, sparkline, checkbox, cell image, dynamic array, rich error and advanced conditional format in one worksheet.',
  behavior: [
    {
      name: 'three saves keep all independent features and their value metadata',
      expect(api: CorpusApi, assert: Assert) {
        const saves = api.combinedFeaturesThroughSaves('combined-native-features/sample.xlsx', 3);
        assert.equal(saves.length, 3);
        for (const saved of saves) {
          assert.deepEqual(
            {
              charts: saved.charts,
              shapes: saved.shapes,
              cellImages: saved.cellImages,
              sparklines: saved.sparklines,
              conditionalFormats: saved.conditionalFormats,
            },
            {charts: 1, shapes: 1, cellImages: 1, sparklines: 1, conditionalFormats: 1},
          );
          assert.deepEqual(saved.checkboxes, ['C4']);
          assert.deepEqual(saved.chartValues, [1, 3]);
          assert.deepEqual(saved.error, {error: '#SPILL!'});
          assert.ok(
            saved.array &&
              typeof saved.array === 'object' &&
              'dynamic' in saved.array &&
              saved.array.dynamic,
          );
        }
      },
    },
  ],
} satisfies Case;
