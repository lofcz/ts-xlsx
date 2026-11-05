import {tryDecodeRange} from './address.ts';
import type {ChartData, ChartOptions} from './chart.ts';
import {dateToSerial} from './date.ts';
import {INTERNAL} from './internal.ts';
import type {Workbook} from './workbook.ts';

/** A reference whose cache was retained because no bounded local range could be resolved. */
export interface ChartCacheDiagnostic {
  readonly sheet: string;
  readonly formula: string;
  readonly reason: 'unresolved-reference' | 'range-limit' | 'uncalculated-formula';
}

/** Refresh chart references using stored cell values and formula results, without evaluating formulas. */
export function resolveChartCaches(
  workbook: Workbook,
  home: string,
  options: ChartOptions,
  diagnostics: ChartCacheDiagnostic[] = [],
): ChartOptions {
  const resolve = <T extends string | number>(
    data: ChartData<T>,
    numeric: boolean,
  ): ChartData<T> => {
    if (data.formula === undefined) return data;
    const formula = data.formula.replace(/^=/, '');
    const named =
      workbook.definedNames.find((n) => n.name === formula && n.scope === home) ??
      workbook.definedNames.find((n) => n.name === formula && n.scope === undefined);
    const range = tryDecodeRange(named?.refersTo.replace(/^=/, '') ?? formula);
    const sheet = workbook.getWorksheet(range?.sheetName ?? named?.scope ?? home);
    const reject = (reason: ChartCacheDiagnostic['reason']): ChartData<T> => {
      diagnostics.push({sheet: home, formula, reason});
      return data;
    };
    if (
      !range ||
      !sheet ||
      range.top === undefined ||
      range.left === undefined ||
      range.bottom === undefined ||
      range.right === undefined
    )
      return reject('unresolved-reference');
    const count = (range.bottom - range.top + 1) * (range.right - range.left + 1);
    if (count > 100000) return reject('range-limit');
    const values: (string | number | null)[] = [];
    let present = false;
    for (let row = range.top; row <= range.bottom; row++) {
      for (let col = range.left; col <= range.right; col++) {
        const cell = sheet[INTERNAL].peekCell(row, col);
        present ||= cell !== undefined;
        let value = cell?.value;
        if (
          value &&
          typeof value === 'object' &&
          ('formula' in value || 'sharedFormula' in value)
        ) {
          if (value.result === undefined) return reject('uncalculated-formula');
          value = value.result;
        }
        if (value instanceof Date) value = dateToSerial(value, workbook.dateEpoch);
        if (value && typeof value === 'object' && 'richText' in value)
          value = value.richText.map((run) => run.text).join('');
        values.push(
          numeric
            ? typeof value === 'number' && Number.isFinite(value)
              ? value
              : null
            : typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
              ? String(value)
              : null,
        );
      }
    }
    if (!present) return reject('unresolved-reference');
    // T is fixed by the caller's numeric flag; values above only contain that kind or null.
    return {...data, values: values as (T | null)[]};
  };
  return {
    ...options,
    ...(options.categories
      ? {
          categories: resolve(
            options.categories,
            options.categoryAxis?.type === 'date' ||
              options.categories.values.some((value) => typeof value === 'number'),
          ),
        }
      : {}),
    series: options.series.map((series) => ({
      ...series,
      ...(series.nameFormula
        ? {
            name: String(
              resolve({formula: series.nameFormula, values: [series.name]}, false).values[0] ?? '',
            ),
          }
        : {}),
      values: resolve(series.values, true),
      ...(series.categories
        ? {
            categories: resolve(
              series.categories,
              options.categoryAxis?.type === 'date' ||
                series.categories.values.some((value) => typeof value === 'number'),
            ),
          }
        : {}),
      ...(series.categoryLevels
        ? {categoryLevels: series.categoryLevels.map((level) => resolve(level, false))}
        : {}),
      ...(series.errorBars
        ? {
            errorBars: series.errorBars.map((error) => ({
              ...error,
              ...(error.plus ? {plus: resolve(error.plus, true)} : {}),
              ...(error.minus ? {minus: resolve(error.minus, true)} : {}),
            })),
          }
        : {}),
      ...(series.xValues ? {xValues: resolve(series.xValues, true)} : {}),
      ...(series.bubbleSizes ? {bubbleSizes: resolve(series.bubbleSizes, true)} : {}),
    })),
  };
}
