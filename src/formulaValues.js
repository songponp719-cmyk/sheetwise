const MAX_FORMULA_RANGE_CELLS = 100_000;
const MAX_FORMULA_DEPTH = 256;

export async function recalculateFormulaValues(workbook, utils) {
  const { default: FormulaParser } = await import('fast-formula-parser');
  const FormulaError = FormulaParser.FormulaError;
  const formulaStates = new Map();
  const parsers = [];
  const formulaIssues = [];

  function getSheet(sheetName) {
    return workbook.Sheets[sheetName];
  }

  function getStateKey(sheetName, address) {
    return `${sheetName}!${address}`;
  }

  function getOriginalValue(cell) {
    return cell?.v ?? null;
  }

  function getCellValue(sheetName, address, depth) {
    const sheet = getSheet(sheetName);
    if (!sheet) return null;
    const cell = sheet[address];
    if (!cell?.f) return getOriginalValue(cell);

    const key = getStateKey(sheetName, address);
    if (formulaStates.get(key) === 'evaluated') return getOriginalValue(cell);
    if (formulaStates.get(key) === 'evaluating') {
      if (cell.v !== undefined && cell.v !== null) return cell.v;
      throw new Error(`Circular formula reference at ${key}`);
    }

    formulaStates.set(key, 'evaluating');
    let result;
    try {
      const position = utils.decode_cell(address);
      if (depth >= MAX_FORMULA_DEPTH) throw new Error('Formula dependency depth limit exceeded');
      result = getParser(depth).parse(cell.f, {
        row: position.r + 1,
        col: position.c + 1,
        sheet: sheetName,
      });
    } catch (error) {
      formulaIssues.push({
        sheet: sheetName,
        cell: address,
        message: error instanceof Error ? error.message : 'ไม่สามารถคำนวณสูตรนี้ได้',
      });
      if (cell.v === undefined || cell.v === null) {
        const isUnsupportedFunction = error instanceof FormulaError
          && error.message?.includes('is not implemented');
        cell.v = isUnsupportedFunction ? '#NAME?' : '#ERROR!';
        cell.t = 'e';
        delete cell.w;
      }
      formulaStates.set(key, 'evaluated');
      return getOriginalValue(cell);
    }

    if (result instanceof FormulaError) {
      if (result.error === '#NAME?') {
        formulaIssues.push({
          sheet: sheetName,
          cell: address,
          message: result.message || 'สูตรหรือชื่อที่อ้างอิงยังไม่รองรับ',
        });
      }
      if (result.error !== '#NAME?' || cell.v === undefined || cell.v === null) {
        cell.v = result.error;
        cell.t = 'e';
        delete cell.w;
      }
    } else {
      cell.v = result;
      cell.t = typeof result === 'number'
        ? 'n'
        : typeof result === 'boolean'
          ? 'b'
          : typeof result === 'string'
            ? 's'
            : 'e';
      if (cell.t === 'e') {
        cell.v = result == null ? '' : '#VALUE!';
      }
      delete cell.w;
    }

    formulaStates.set(key, 'evaluated');
    return getOriginalValue(cell);
  }

  function getRangeValues(reference, depth) {
    const rowCount = reference.to.row - reference.from.row + 1;
    const columnCount = reference.to.col - reference.from.col + 1;
    if (rowCount * columnCount > MAX_FORMULA_RANGE_CELLS) {
      throw new Error('Formula range exceeds the calculation limit');
    }
    const sheetName = reference.sheet || '';
    const rows = [];
    for (let row = reference.from.row; row <= reference.to.row; row += 1) {
      const values = [];
      for (let col = reference.from.col; col <= reference.to.col; col += 1) {
        values.push(getCellValue(sheetName, utils.encode_cell({ r: row - 1, c: col - 1 }), depth));
      }
      rows.push(values);
    }
    return rows;
  }

  function getParser(depth) {
    if (!parsers[depth]) {
      parsers[depth] = new FormulaParser({
        onCell: (reference) => getCellValue(
          reference.sheet,
          reference.address || utils.encode_cell({ r: reference.row - 1, c: reference.col - 1 }),
          depth + 1,
        ),
        onRange: (reference) => getRangeValues(reference, depth + 1),
        onVariable: () => null,
      });
    }
    return parsers[depth];
  }

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    Object.keys(sheet).forEach((address) => {
      if (/^[A-Z]+[0-9]+$/i.test(address) && sheet[address]?.f) {
        formulaStates.set(getStateKey(sheetName, address), 'pending');
      }
    });
  });

  for (const [key] of formulaStates) {
    const separator = key.lastIndexOf('!');
    getCellValue(key.slice(0, separator), key.slice(separator + 1), 0);
  }

  return formulaIssues;
}
