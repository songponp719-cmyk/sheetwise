import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowLeftRight, ArrowUpRight, Check, ChevronDown, ChevronLeft, ChevronRight, CircleHelp,
  CloudUpload, Download, FileSpreadsheet, Files, Filter, Info, LoaderCircle, Plus,
  Search, ShieldCheck, Sparkles, Table2, X,
} from 'lucide-react';

const FILE_LIMIT = 25 * 1024 * 1024;
const FILE_PREVIEW_PAGE_SIZE = 10;
const RESULTS_PAGE_SIZE = 25;

function normalizeHeader(value, index) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  return text || `คอลัมน์ ${index + 1}`;
}

function normalizeCell(value) {
  if (value == null) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'string') {
    const text = value.replace(/\u00a0/g, ' ').trim();
    if (!text) return '';
    const numeric = text.replace(/,/g, '');
    if (/^-?\d+(\.\d+)?$/.test(numeric)) return numeric;
    return text;
  }
  return value;
}

function findHeaderRowIndex(rows) {
  const candidates = rows.slice(0, 20);
  let bestIndex = 0;
  let bestScore = Number.NEGATIVE_INFINITY;
  candidates.forEach((row, index) => {
    const values = row.map((cell) => String(cell ?? '').replace(/\s+/g, ' ').trim()).filter(Boolean);
    if (!values.length) return;
    const uniqueCount = new Set(values.map((value) => value.toLocaleLowerCase())).size;
    const textCount = values.filter((value) => Number.isNaN(Number(value.replace(/,/g, '')))).length;
    const score = values.length * 4 + uniqueCount * 2 + textCount / values.length - index * 2.5;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });
  return bestIndex;
}

async function parseWorkbook(buffer, filename) {
  const XLSX = await import('@e965/xlsx');
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  if (!workbook.SheetNames.length) throw new Error('ไม่พบชีตข้อมูลในไฟล์นี้');
  const sheets = workbook.SheetNames.map((name) => {
    const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false });
    const nonEmptyRows = rawRows.filter((row) => row.some((cell) => String(cell ?? '').trim() !== ''));
    if (nonEmptyRows.length < 2) return { name, headers: [], rows: [] };

    const headerIndex = findHeaderRowIndex(nonEmptyRows);
    const headerRow = nonEmptyRows[headerIndex];
    const dataRows = nonEmptyRows.slice(headerIndex + 1);
    const width = Math.max(...nonEmptyRows.map((row) => row.length));
    const headers = Array.from({ length: width }, (_, index) => normalizeHeader(headerRow[index], index));
    const seen = new Map();
    const uniqueHeaders = headers.map((header) => {
      const count = (seen.get(header) || 0) + 1;
      seen.set(header, count);
      return count === 1 ? header : `${header} (${count})`;
    });
    const rows = dataRows.map((row, index) => {
      const record = { __rowNumber: index + headerIndex + 2 };
      uniqueHeaders.forEach((header, colIndex) => {
        record[header] = normalizeCell(row[colIndex]);
      });
      return record;
    }).filter((record) => uniqueHeaders.some((header) => record[header] !== ''));
    return { name, headers: uniqueHeaders, rows };
  });

  if (!sheets.some((sheet) => sheet.rows.length)) throw new Error('ไม่พบข้อมูลในชีตของไฟล์นี้');
  return { filename, sheets };
}

function keyFor(value) {
  return String(value ?? '').trim().toLocaleLowerCase();
}

function editDistanceWithin(left, right, maxDistance) {
  const a = Array.from(left);
  const b = Array.from(right);
  if (Math.abs(a.length - b.length) > maxDistance) return false;
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let row = 1; row <= a.length; row += 1) {
    const current = [row];
    let rowMinimum = row;
    for (let column = 1; column <= b.length; column += 1) {
      const cost = a[row - 1] === b[column - 1] ? 0 : 1;
      const distance = Math.min(
        previous[column] + 1,
        current[column - 1] + 1,
        previous[column - 1] + cost,
      );
      current.push(distance);
      rowMinimum = Math.min(rowMinimum, distance);
    }
    if (rowMinimum > maxDistance) return false;
    previous = current;
  }
  return previous[b.length] <= maxDistance;
}

function cellMatchesQuery(value, query) {
  const normalizedValue = String(value ?? '').trim().toLocaleLowerCase();
  if (!normalizedValue) return false;
  if (normalizedValue.includes(query)) return true;
  const queryLength = Array.from(query).length;
  if (queryLength < 3 || queryLength > 48) return false;
  const maxDistance = queryLength >= 8 ? 2 : 1;
  const words = normalizedValue.split(/[^\p{L}\p{N}]+/u).filter(Boolean);
  return words.some((word) => editDistanceWithin(word, query, maxDistance));
}

function rowMatchesQuery(row, query) {
  if (!query) return true;
  return Object.values(row.values).some((value) => cellMatchesQuery(value, query));
}

function getMatchingSheetNames(fileA, fileB) {
  const nonEmptyB = new Set(fileB?.sheets.filter((sheet) => sheet.rows.length).map((sheet) => sheet.name) || []);
  return fileA?.sheets.filter((sheet) => sheet.rows.length && nonEmptyB.has(sheet.name)).map((sheet) => sheet.name) || [];
}

function getSheet(file, name) {
  return file?.sheets.find((sheet) => sheet.name === name);
}

function getDefaultKeys(sheetA, sheetB, previousKeys = []) {
  const sharedKey = sheetA.headers.find((header) => sheetB.headers.includes(header));
  return [
    sheetA.headers.includes(previousKeys[0]) ? previousKeys[0] : sharedKey || sheetA.headers[0] || '',
    sheetB.headers.includes(previousKeys[1]) ? previousKeys[1] : sharedKey || sheetB.headers[0] || '',
  ];
}

function compareFiles(fileA, fileB, keyA, keyB) {
  const bucketsA = new Map();
  const bucketsB = new Map();
  const onlyA = [];
  const onlyB = [];
  const addToBucket = (buckets, row, key, unmatched) => {
    const normalized = keyFor(row[key]);
    if (!normalized) {
      unmatched.push(row);
      return;
    }
    if (!buckets.has(normalized)) buckets.set(normalized, []);
    buckets.get(normalized).push(row);
  };
  fileA.rows.forEach((row) => addToBucket(bucketsA, row, keyA, onlyA));
  fileB.rows.forEach((row) => addToBucket(bucketsB, row, keyB, onlyB));

  const matches = [];
  const allKeys = new Set([...bucketsA.keys(), ...bucketsB.keys()]);
  allKeys.forEach((key) => {
    const rowsA = bucketsA.get(key) || [];
    const rowsB = bucketsB.get(key) || [];
    const pairedCount = Math.min(rowsA.length, rowsB.length);
    for (let index = 0; index < pairedCount; index += 1) {
      matches.push({ key: keyFor(rowsA[index][keyA]), rowA: rowsA[index], rowB: rowsB[index] });
    }
    onlyA.push(...rowsA.slice(pairedCount));
    onlyB.push(...rowsB.slice(pairedCount));
  });
  return { matches, onlyA, onlyB };
}

function createCombinedRows(sheetResults, fileA, fileB) {
  const comparedByName = new Map(sheetResults.map((item) => [item.name, item]));
  const dataHeaders = [...new Set([
    ...fileA.sheets.flatMap((sheet) => sheet.headers.filter((header) => header !== comparedByName.get(sheet.name)?.keyA)),
    ...fileB.sheets.flatMap((sheet) => sheet.headers.filter((header) => header !== comparedByName.get(sheet.name)?.keyB)),
  ])];
  const usedHeaders = new Set(dataHeaders);
  const uniqueHeader = (preferred) => {
    let header = preferred;
    let suffix = 2;
    while (usedHeaders.has(header)) {
      header = `${preferred} (${suffix})`;
      suffix += 1;
    }
    usedHeaders.add(header);
    return header;
  };
  const sourceHeader = uniqueHeader('แหล่งข้อมูล');
  const keyHeader = uniqueHeader('Key สำหรับเปรียบเทียบ');
  const sheetHeader = uniqueHeader('ชีต');
  const groups = [];
  let groupIndex = 0;
  const addGroup = (sheetName, sortKey, entries) => {
    const groupId = groupIndex;
    groupIndex += 1;
    groups.push({
      sheetName,
      sortKey,
      groupIndex: groupId,
      rows: entries.map(({ source, row, key, rowNumber }) => {
        const values = {
          [sheetHeader]: sheetName,
          [sourceHeader]: source,
          [keyHeader]: row[key],
          ...Object.fromEntries(dataHeaders.map((header) => [header, row[header] ?? ''])),
        };
        return { id: `${source}-${rowNumber}-${groupId}`, values };
      }),
    });
  };

  const pairedNames = new Set(sheetResults.map(({ name }) => name));
  sheetResults.forEach(({ name, sheetA, sheetB, keyA, keyB, result }) => {
    result.matches.forEach(({ key, rowA, rowB }) => {
      addGroup(name, key, [
        { source: 'ไฟล์ 1', row: rowA, key: keyA, rowNumber: rowA.__rowNumber },
        { source: 'ไฟล์ 2', row: rowB, key: keyB, rowNumber: rowB.__rowNumber },
      ]);
    });
    result.onlyA.forEach((row) => {
      addGroup(name, keyFor(row[keyA]), [{ source: 'ไฟล์ 1', row, key: keyA, rowNumber: row.__rowNumber }]);
    });
    result.onlyB.forEach((row) => {
      addGroup(name, keyFor(row[keyB]), [{ source: 'ไฟล์ 2', row, key: keyB, rowNumber: row.__rowNumber }]);
    });
  });

  fileA.sheets.filter((sheet) => !pairedNames.has(sheet.name)).forEach((sheet) => {
    sheet.rows.forEach((row) => addGroup(sheet.name, '', [{ source: 'ไฟล์ 1', row, key: '', rowNumber: row.__rowNumber }]));
  });
  fileB.sheets.filter((sheet) => !pairedNames.has(sheet.name)).forEach((sheet) => {
    sheet.rows.forEach((row) => addGroup(sheet.name, '', [{ source: 'ไฟล์ 2', row, key: '', rowNumber: row.__rowNumber }]));
  });

  const sheetOrder = new Map([...fileA.sheets, ...fileB.sheets.filter((sheet) => !getSheet(fileA, sheet.name))]
    .map((sheet, index) => [sheet.name, index]));
  groups.sort((a, b) => (sheetOrder.get(a.sheetName) - sheetOrder.get(b.sheetName))
    || a.sortKey.localeCompare(b.sortKey, undefined, { numeric: true })
    || a.groupIndex - b.groupIndex);
  return groups.flatMap((group) => group.rows);
}

async function downloadWorkbook(sheets, filename) {
  const XLSX = await import('@e965/xlsx');
  const workbook = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const safeName = name.slice(0, 31);
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), safeName);
  });
  XLSX.writeFile(workbook, filename);
}

function DropZone({ index, file, onFile, onClear, error, inputRef, accept = '.xlsx,.xls,.csv' }) {
  const [dragging, setDragging] = useState(false);
  const handleFiles = (files) => {
    const selected = files?.[0];
    if (selected) onFile(index, selected);
  };
  return (
    <div
      className={`upload-card ${dragging ? 'is-dragging' : ''} ${file ? 'has-file' : ''}`}
      onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={(event) => { event.preventDefault(); setDragging(false); handleFiles(event.dataTransfer.files); }}
    >
      <input ref={inputRef} type="file" accept={accept} hidden onChange={(event) => {
        handleFiles(event.target.files);
        event.target.value = '';
      }} />
      {file ? (
        <div className="uploaded-file">
          <span className="file-icon"><FileSpreadsheet size={22} /></span>
          <div className="file-info">
            <strong title={file.filename}>{file.filename}</strong>
            <span>{file.sheets.reduce((total, sheet) => total + sheet.rows.length, 0).toLocaleString()} แถว · {file.sheets.length} ชีต</span>
          </div>
          <button className="icon-button clear-file" onClick={onClear} aria-label="ลบไฟล์"><X size={17} /></button>
        </div>
      ) : (
        <>
          <button className="upload-target" onClick={() => inputRef.current?.click()}>
            <span className="upload-icon"><CloudUpload size={23} /></span>
            <span className="upload-copy">
              <strong>เลือกไฟล์จากอุปกรณ์</strong>
              <span>หรือลากไฟล์มาวางที่นี่</span>
            </span>
            <ArrowUpRight size={17} className="upload-arrow" />
          </button>
          <span className="file-hint">รองรับ XLSX, XLS, CSV · ไม่เกิน 25 MB</span>
        </>
      )}
      {error && <p className="error-text">{error}</p>}
    </div>
  );
}

function TablePagination({ page, pageSize, total, onPageChange }) {
  const pageCount = Math.max(1, Math.ceil(total / pageSize));
  const currentPage = Math.min(page, pageCount - 1);
  const firstRow = total ? currentPage * pageSize + 1 : 0;
  const lastRow = Math.min((currentPage + 1) * pageSize, total);
  return (
    <div className="table-footer">
      <span>แสดง {firstRow.toLocaleString()}–{lastRow.toLocaleString()} จาก {total.toLocaleString()} รายการ</span>
      {pageCount > 1 && <div className="pagination-controls">
        <button className="pagination-button" disabled={currentPage === 0} onClick={() => onPageChange(currentPage - 1)} aria-label="หน้าก่อนหน้า"><ChevronLeft size={13} /></button>
        <span>{currentPage + 1} / {pageCount}</span>
        <button className="pagination-button" disabled={currentPage >= pageCount - 1} onClick={() => onPageChange(currentPage + 1)} aria-label="หน้าถัดไป"><ChevronRight size={13} /></button>
      </div>}
    </div>
  );
}

function DataTable({ file, sheet, fileIndex, activeSheet, onSheetChange, search, setSearch }) {
  const [page, setPage] = useState(0);
  const filteredRows = useMemo(() => {
    const query = search.trim().toLocaleLowerCase();
    if (!query) return sheet.rows;
    return sheet.rows.filter((row) => sheet.headers.some((header) => String(row[header] ?? '').toLocaleLowerCase().includes(query)));
  }, [sheet, search]);
  const pageCount = Math.max(1, Math.ceil(filteredRows.length / FILE_PREVIEW_PAGE_SIZE));
  const currentPage = Math.min(page, pageCount - 1);
  const visibleRows = filteredRows.slice(currentPage * FILE_PREVIEW_PAGE_SIZE, (currentPage + 1) * FILE_PREVIEW_PAGE_SIZE);
  return (
    <section className="data-panel">
      <div className="panel-title-row">
        <div className="panel-file-mark">{fileIndex === 0 ? '01' : '02'}</div>
        <div className="panel-title">
          <strong>{file.filename}</strong>
          <span>ชีต {sheet.name} · {filteredRows.length.toLocaleString()} จาก {sheet.rows.length.toLocaleString()} รายการ</span>
        </div>
        <label className="sheet-select-wrap"><span className="visually-hidden">เลือกชีตพรีวิวไฟล์ {fileIndex + 1}</span><select value={activeSheet} onChange={(event) => onSheetChange(event.target.value)}>{file.sheets.map((item) => <option key={item.name} value={item.name}>{item.name}</option>)}</select><ChevronDown size={14} /></label>
        <span className={`file-status ${fileIndex === 0 ? 'blue' : 'purple'}`}><Check size={13} /> พร้อมใช้งาน</span>
      </div>
      <label className="table-search">
        <Search size={16} />
        <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="ค้นหาทุกคอลัมน์..." />
        <Filter size={15} className="filter-mark" />
      </label>
      <div className="table-scroll">
        <table>
          <thead><tr>{sheet.headers.map((header) => <th key={header}>{header}<ChevronDown size={12} /></th>)}</tr></thead>
          <tbody>
            {visibleRows.map((row, index) => (
              <tr key={`${row.__rowNumber}-${index}`}>
                {sheet.headers.map((header) => <td key={header} title={String(row[header] ?? '')}>{String(row[header] ?? '') || <span className="empty-cell">—</span>}</td>)}
              </tr>
            ))}
            {!filteredRows.length && <tr><td className="table-empty" colSpan={Math.max(1, sheet.headers.length)}>{search ? 'ไม่พบข้อมูลที่ตรงกับคำค้นหา' : 'ชีตนี้ไม่มีข้อมูลสำหรับพรีวิว'}</td></tr>}
          </tbody>
        </table>
      </div>
      {filteredRows.length > 0 && <TablePagination page={currentPage} pageSize={FILE_PREVIEW_PAGE_SIZE} total={filteredRows.length} onPageChange={setPage} />}
    </section>
  );
}

function ResultTable({ rows, allRows, sheetNames, sheetFilter, onSheetFilterChange, search, onSearchChange, selectionMode, selectedIds, onToggleRow, onTogglePage, onRemoveRow }) {
  const [page, setPage] = useState(0);
  const columns = allRows.length ? Object.keys(allRows[0].values) : [];
  const pageCount = Math.max(1, Math.ceil(rows.length / RESULTS_PAGE_SIZE));
  const currentPage = Math.min(page, pageCount - 1);
  const pageRows = rows.slice(currentPage * RESULTS_PAGE_SIZE, (currentPage + 1) * RESULTS_PAGE_SIZE);
  const allVisibleSelected = pageRows.length > 0 && pageRows.every((row) => selectedIds.has(row.id));
  return (
    <>
      <div className="result-table-controls">
        <label className="table-search results-search">
          <Search size={16} />
          <input value={search} onChange={(event) => { onSearchChange(event.target.value); setPage(0); }} placeholder="ค้นหาแถวที่ตรงหรือใกล้เคียง..." />
          <Filter size={15} className="filter-mark" />
        </label>
        <label className="result-sheet-filter">
          <span>กรองชีต</span>
          <select value={sheetFilter} onChange={(event) => onSheetFilterChange(event.target.value)}>
            <option value="ทั้งหมด">ทุกชีต</option>
            {sheetNames.map((name) => <option key={name} value={name}>{name}</option>)}
          </select>
        </label>
      </div>
      {allRows.length ? (
        <div className="table-scroll results-scroll">
          <table><thead><tr>
            <th className="row-action-column"><span className="visually-hidden">นำแถวออก</span></th>
            {selectionMode && <th className="selection-column"><input type="checkbox" aria-label="เลือกทุกรายการในหน้านี้" checked={allVisibleSelected} onChange={() => onTogglePage(pageRows, allVisibleSelected)} /></th>}
            {columns.map((column) => <th key={column}>{column}</th>)}
          </tr></thead>
            <tbody>{pageRows.map((row) => (
              <tr key={row.id} className={selectedIds.has(row.id) ? 'selected-result-row' : ''}>
                <td className="row-action-column"><button className="remove-row-button" aria-label={`นำแถวจาก${row.values['แหล่งข้อมูล']} Key ${row.values['Key สำหรับเปรียบเทียบ']} ออก`} title="นำแถวนี้ออกจากตาราง" onClick={() => onRemoveRow(row.id)}><X size={14} /></button></td>
                {selectionMode && <td className="selection-column"><input type="checkbox" aria-label={`เลือกแถวจาก${row.values['แหล่งข้อมูล']} Key ${row.values['Key สำหรับเปรียบเทียบ']}`} checked={selectedIds.has(row.id)} onChange={() => onToggleRow(row.id)} /></td>}
                {columns.map((column) => <td key={column} title={String(row.values[column] ?? '')}>{String(row.values[column] ?? '') || <span className="empty-cell">—</span>}</td>)}
              </tr>
            ))}</tbody>
          </table>
          {!rows.length && <div className="no-results">ไม่พบแถวข้อมูลที่ตรงหรือใกล้เคียงกับคำค้นหา</div>}
        </div>
      ) : <div className="result-empty"><Files size={24} /><span>ไม่มีข้อมูลในชีตที่เลือก</span></div>}
      {rows.length > 0 && <TablePagination page={currentPage} pageSize={RESULTS_PAGE_SIZE} total={rows.length} onPageChange={setPage} />}
    </>
  );
}

function TableBuilderPage({ file, fileError, isReading, onFile, onClear, addedRows, outputSheets, onAddRow, onRemoveRow, onAddSheet, onRenameSheet, onExport }) {
  const [activeSheet, setActiveSheet] = useState('ทั้งหมด');
  const [activeOutputSheetId, setActiveOutputSheetId] = useState('');
  const [previewSheetWindowStart, setPreviewSheetWindowStart] = useState(0);
  const [multiSelectMode, setMultiSelectMode] = useState(false);
  const [selectedSourceIds, setSelectedSourceIds] = useState(() => new Set());
  const [isChoosingOutputSheet, setIsChoosingOutputSheet] = useState(false);
  const [batchOutputSheetId, setBatchOutputSheetId] = useState('');
  const [addSuccess, setAddSuccess] = useState(null);
  const [sheetNameDraft, setSheetNameDraft] = useState('');
  const [sheetNameError, setSheetNameError] = useState('');
  const [search, setSearch] = useState('');
  const [sourcePage, setSourcePage] = useState(0);
  const [outputPage, setOutputPage] = useState(0);
  const inputRef = useRef(null);
  const addSuccessTimers = useRef([]);
  const activeOutputSheet = outputSheets.find((sheet) => sheet.id === activeOutputSheetId) || outputSheets[0];
  const outputRows = addedRows.filter((item) => item.outputSheetId === activeOutputSheet?.id);
  const headers = useMemo(() => [...new Set(file?.sheets.flatMap((sheet) => sheet.headers) || [])], [file]);
  const sourceHeader = headers.includes('ชีตต้นทาง') ? 'ชีตต้นทาง (แหล่งที่มา)' : 'ชีตต้นทาง';
  useEffect(() => {
    setActiveSheet('ทั้งหมด');
    setSearch('');
    setSourcePage(0);
    setOutputPage(0);
    setPreviewSheetWindowStart(0);
    setSelectedSourceIds(new Set());
    setMultiSelectMode(false);
  }, [file]);
  useEffect(() => {
    if (!outputSheets.some((sheet) => sheet.id === activeOutputSheetId)) {
      setActiveOutputSheetId(outputSheets[0]?.id || '');
    }
  }, [outputSheets, activeOutputSheetId]);
  useEffect(() => {
    setSheetNameDraft(activeOutputSheet?.name || '');
    setSheetNameError('');
  }, [activeOutputSheet?.id, activeOutputSheet?.name]);
  useEffect(() => () => {
    addSuccessTimers.current.forEach((timer) => window.clearTimeout(timer));
  }, []);
  const sourceRows = useMemo(() => {
    if (!file) return [];
    const query = search.trim().toLocaleLowerCase();
    return file.sheets
      .filter((sheet) => activeSheet === 'ทั้งหมด' || sheet.name === activeSheet)
      .flatMap((sheet) => sheet.rows.map((row) => ({
        id: `${sheet.name}:${row.__rowNumber}`,
        sheetName: sheet.name,
        row,
      })))
      .filter(({ row }) => !query || headers.some((header) => String(row[header] ?? '').toLocaleLowerCase().includes(query)));
  }, [file, activeSheet, search, headers]);
  const allSourceRows = useMemo(() => file?.sheets.flatMap((sheet) => sheet.rows.map((row) => ({
    id: `${sheet.name}:${row.__rowNumber}`,
    sheetName: sheet.name,
    row,
  }))) || [], [file]);
  const selectedSourceRows = allSourceRows.filter(({ id }) => selectedSourceIds.has(id));
  const sourcePageCount = Math.max(1, Math.ceil(sourceRows.length / FILE_PREVIEW_PAGE_SIZE));
  const visibleSourceRows = sourceRows.slice(
    Math.min(sourcePage, sourcePageCount - 1) * FILE_PREVIEW_PAGE_SIZE,
    (Math.min(sourcePage, sourcePageCount - 1) + 1) * FILE_PREVIEW_PAGE_SIZE,
  );
  const outputPageCount = Math.max(1, Math.ceil(outputRows.length / FILE_PREVIEW_PAGE_SIZE));
  const visibleOutputRows = outputRows.slice(
    Math.min(outputPage, outputPageCount - 1) * FILE_PREVIEW_PAGE_SIZE,
    (Math.min(outputPage, outputPageCount - 1) + 1) * FILE_PREVIEW_PAGE_SIZE,
  );
  const visibleSourceIds = visibleSourceRows.map(({ id }) => id);
  const allVisibleSourceSelected = visibleSourceIds.length > 0 && visibleSourceIds.every((id) => selectedSourceIds.has(id));
  const sheetWindowEnd = Math.min(previewSheetWindowStart + 5, file?.sheets.length || 0);

  function showAddedMessage(count, sheetName) {
    addSuccessTimers.current.forEach((timer) => window.clearTimeout(timer));
    setAddSuccess({ count, sheetName, fading: false, id: Date.now() });
    addSuccessTimers.current = [
      window.setTimeout(() => setAddSuccess((current) => current && { ...current, fading: true }), 1000),
      window.setTimeout(() => setAddSuccess(null), 1500),
    ];
  }

  function toggleSourceRow(id) {
    setSelectedSourceIds((previous) => {
      const next = new Set(previous);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleVisibleSourceRows() {
    setSelectedSourceIds((previous) => {
      const next = new Set(previous);
      visibleSourceIds.forEach((id) => {
        if (allVisibleSourceSelected) next.delete(id);
        else next.add(id);
      });
      return next;
    });
  }

  function addOneSourceRow(item) {
    if (!activeOutputSheet) return;
    onAddRow({ ...item, outputSheetId: activeOutputSheet.id });
    showAddedMessage(1, activeOutputSheet.name);
    setSelectedSourceIds((previous) => {
      const next = new Set(previous);
      next.delete(item.id);
      return next;
    });
  }

  function confirmAddSelectedRows() {
    if (!batchOutputSheetId || !selectedSourceRows.length) return;
    const destinationSheet = outputSheets.find((sheet) => sheet.id === batchOutputSheetId);
    selectedSourceRows.forEach((item) => onAddRow({ ...item, outputSheetId: batchOutputSheetId }));
    showAddedMessage(selectedSourceRows.length, destinationSheet?.name || 'ชีตที่เลือก');
    setSelectedSourceIds(new Set());
    setIsChoosingOutputSheet(false);
  }

  function addOutputSheet() {
    let number = outputSheets.length + 1;
    let name = `ชีต ${number}`;
    while (outputSheets.some((sheet) => sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())) {
      number += 1;
      name = `ชีต ${number}`;
    }
    const sheet = { id: `builder-sheet-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`, name };
    onAddSheet(sheet);
    setActiveOutputSheetId(sheet.id);
    setOutputPage(0);
  }

  function saveSheetName() {
    const name = sheetNameDraft.trim();
    if (!name) {
      setSheetNameError('กรุณาระบุชื่อชีต');
      return;
    }
    if (name.length > 31) {
      setSheetNameError('ชื่อชีตต้องไม่เกิน 31 ตัวอักษร');
      return;
    }
    if (['\\', '/', '?', '*', '[', ']', ':'].some((character) => name.includes(character))) {
      setSheetNameError('ชื่อชีตห้ามมีอักขระ \\, /, ?, *, [, ], หรือ :');
      return;
    }
    if (outputSheets.some((sheet) => sheet.id !== activeOutputSheet?.id && sheet.name.toLocaleLowerCase() === name.toLocaleLowerCase())) {
      setSheetNameError('มีชีตชื่อนี้อยู่แล้ว');
      return;
    }
    onRenameSheet(activeOutputSheet.id, name);
    setSheetNameError('');
  }

  return (
    <>
      <section className="intro">
        <div className="eyebrow"><span className="eyebrow-dot" /> DATA WORKSPACE</div>
        <h1>สร้างตาราง<br className="mobile-break" /> <span>ในแบบของคุณ</span></h1>
        <p>เลือกดูชีต ค้นหาแถวที่ต้องการ แล้วกดปุ่มบวกเพื่อเพิ่มข้อมูลลงในตารางใหม่</p>
      </section>

      <section className="section-block table-builder-section">
        <div className="section-heading">
          <div><div className="section-kicker">STEP 01 <span>·</span> UPLOAD</div><h2>อัปโหลดไฟล์ Excel</h2><p>รองรับไฟล์ .xlsx และ .xls โดยประมวลผลในอุปกรณ์นี้</p></div>
          {file && <span className="file-count"><Files size={15} /> {file.sheets.length} ชีต</span>}
        </div>
        <DropZone
          index={0}
          file={file}
          onFile={(_, selectedFile) => onFile(selectedFile)}
          onClear={onClear}
          error={fileError}
          inputRef={inputRef}
          accept=".xlsx,.xls"
        />
        {isReading && <div className="reading-indicator"><LoaderCircle size={16} className="spin" /> กำลังอ่านและจัดระเบียบไฟล์ในเบราว์เซอร์...</div>}
      </section>

      {file && <>
        <section className="section-block builder-source-section">
          <div className="section-heading">
            <div><div className="section-kicker">STEP 02 <span>·</span> REVIEW DATA</div><h2>พรีวิวและเลือกข้อมูลจากไฟล์</h2><p>เลือกชีตหรือค้นหาข้อมูล จากนั้นกดปุ่ม + ที่แถวนั้น</p></div>
          </div>
          <div className="builder-controls">
            <label className="table-search">
              <Search size={16} />
              <input value={search} onChange={(event) => { setSearch(event.target.value); setSourcePage(0); }} placeholder="ค้นหาข้อมูลในทุกคอลัมน์..." />
              <Filter size={15} className="filter-mark" />
            </label>
            <div className="source-sheet-pager" role="group" aria-label="เลือกชีตพรีวิว">
              <button type="button" className={`source-all-sheets ${activeSheet === 'ทั้งหมด' ? 'active' : ''}`} onClick={() => { setActiveSheet('ทั้งหมด'); setSourcePage(0); }}>ทุกชีต</button>
              <button type="button" className="sheet-window-arrow" disabled={previewSheetWindowStart === 0} onClick={() => { setPreviewSheetWindowStart(Math.max(0, previewSheetWindowStart - 5)); setSourcePage(0); }} aria-label="ชีต 5 ชีตก่อนหน้า"><ChevronLeft size={18} /></button>
              <div className="source-sheet-dots">
                {file.sheets.slice(previewSheetWindowStart, sheetWindowEnd).map((sheet, index) => {
                  const sheetIndex = previewSheetWindowStart + index;
                  return <button key={sheet.name} type="button" className={`source-sheet-dot ${activeSheet === sheet.name ? 'active' : ''}`} onClick={() => { setActiveSheet(sheet.name); setSourcePage(0); }} aria-label={`แสดงชีต ${sheet.name}`} aria-pressed={activeSheet === sheet.name} title={sheet.name}>{sheetIndex + 1}</button>;
                })}
              </div>
              <button type="button" className="sheet-window-arrow" disabled={sheetWindowEnd >= file.sheets.length} onClick={() => { setPreviewSheetWindowStart(Math.min(Math.max(0, file.sheets.length - 1), previewSheetWindowStart + 5)); setSourcePage(0); }} aria-label="ชีต 5 ชีตถัดไป"><ChevronRight size={18} /></button>
              <span className="source-sheet-name" title={activeSheet === 'ทั้งหมด' ? 'ทุกชีต' : activeSheet}>{activeSheet === 'ทั้งหมด' ? 'ทุกชีต' : activeSheet}</span>
            </div>
            <div className="builder-selection-actions">
              <button type="button" className={`select-mode-button ${multiSelectMode ? 'active' : ''}`} aria-pressed={multiSelectMode} onClick={() => { setMultiSelectMode((enabled) => !enabled); setSelectedSourceIds(new Set()); }}>
                {multiSelectMode ? <X size={15} /> : <Check size={15} />} {multiSelectMode ? 'ยกเลิกการเลือก' : 'เลือกหลายแถว'}
              </button>
              <button type="button" className="export-button add-selected-button" disabled={!selectedSourceRows.length || !outputSheets.length} onClick={() => { setBatchOutputSheetId(activeOutputSheet?.id || outputSheets[0]?.id || ''); setIsChoosingOutputSheet(true); }}>
                <Plus size={16} /> เพิ่มข้อมูลที่เลือก ({selectedSourceRows.length.toLocaleString()})
              </button>
            </div>
          </div>
          <div className="table-scroll builder-table-scroll">
            <table>
              <thead><tr>
                <th className="builder-add-column">{multiSelectMode
                  ? <input type="checkbox" aria-label="เลือกทุกรายการในหน้านี้" checked={allVisibleSourceSelected} onChange={toggleVisibleSourceRows} disabled={!visibleSourceIds.length} />
                  : 'เพิ่ม'}</th>
                {activeSheet === 'ทั้งหมด' && <th>{sourceHeader}</th>}{headers.map((header) => <th key={header}>{header}</th>)}
              </tr></thead>
              <tbody>
                {visibleSourceRows.map(({ id, sheetName, row }) => {
                  return <tr key={id}>
                    <td className="builder-add-column">{multiSelectMode
                      ? <input type="checkbox" checked={selectedSourceIds.has(id)} aria-label={`เลือกแถว ${row.__rowNumber} จากชีต ${sheetName}`} onChange={() => toggleSourceRow(id)} />
                      : <button className="add-row-button" onClick={() => addOneSourceRow({ id, sheetName, row })} disabled={!activeOutputSheet} aria-label={`เพิ่มแถวจากชีต ${sheetName} ลงใน ${activeOutputSheet?.name || 'ตาราง'}`} title={`เพิ่มลงใน ${activeOutputSheet?.name || 'ตาราง'}`}><Plus size={17} /></button>}</td>
                    {activeSheet === 'ทั้งหมด' && <td>{sheetName}</td>}
                    {headers.map((header) => <td key={header} title={String(row[header] ?? '')}>{String(row[header] ?? '') || <span className="empty-cell">—</span>}</td>)}
                  </tr>;
                })}
                {!sourceRows.length && <tr><td className="table-empty" colSpan={headers.length + (activeSheet === 'ทั้งหมด' ? 2 : 1)}>{search ? 'ไม่พบข้อมูลที่ตรงกับคำค้นหา' : 'ชีตนี้ไม่มีข้อมูลสำหรับพรีวิว'}</td></tr>}
              </tbody>
            </table>
          </div>
          {sourceRows.length > 0 && <TablePagination page={Math.min(sourcePage, sourcePageCount - 1)} pageSize={FILE_PREVIEW_PAGE_SIZE} total={sourceRows.length} onPageChange={setSourcePage} />}
        </section>

        <section className="section-block builder-output-section">
          <div className="section-heading">
            <div><div className="section-kicker">STEP 03 <span>·</span> CUSTOM TABLE</div><h2>ตารางใหม่ที่ปรับแต่ง</h2><p>เพิ่มชีท ตั้งชื่อชีท และเลือกชีทปลายทางก่อนเพิ่มแถวจากพรีวิว</p></div>
            <div className="builder-output-actions"><span className="file-count"><Files size={15} /> {addedRows.length.toLocaleString()} แถว · {outputSheets.length} ชีต</span><button className="export-button" disabled={!outputSheets.length} onClick={onExport}><Download size={15} /> ส่งออก Excel</button></div>
          </div>
          <div className="builder-sheet-management">
            <div className="builder-sheet-tabs" role="tablist" aria-label="ชีตในตารางใหม่">
              {outputSheets.map((sheet) => <button key={sheet.id} type="button" role="tab" aria-selected={activeOutputSheet?.id === sheet.id} className={`builder-sheet-tab ${activeOutputSheet?.id === sheet.id ? 'active' : ''}`} onClick={() => { setActiveOutputSheetId(sheet.id); setOutputPage(0); }}>{sheet.name}<span>{addedRows.filter((item) => item.outputSheetId === sheet.id).length}</span></button>)}
              <button className="add-sheet-button" type="button" onClick={addOutputSheet}><Plus size={15} /> เพิ่มชีท</button>
            </div>
            {activeOutputSheet && <div className="builder-sheet-rename">
              <label htmlFor="builder-sheet-name">ชื่อชีท</label>
              <input id="builder-sheet-name" value={sheetNameDraft} maxLength={31} onChange={(event) => { setSheetNameDraft(event.target.value); setSheetNameError(''); }} onKeyDown={(event) => { if (event.key === 'Enter') saveSheetName(); }} />
              <button type="button" className="rename-sheet-button" onClick={saveSheetName}>บันทึกชื่อ</button>
              <span className="sheet-name-hint">ไม่เกิน 31 ตัวอักษร</span>
              {sheetNameError && <span className="sheet-name-error" role="alert">{sheetNameError}</span>}
            </div>}
          </div>
          <div className="table-scroll builder-table-scroll">
            <table>
              <thead><tr><th className="builder-add-column">นำออก</th><th>{sourceHeader}</th>{headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
              <tbody>
                {visibleOutputRows.map(({ id, sheetName, row }) => <tr key={id}>
                  <td className="builder-add-column"><button className="remove-row-button" onClick={() => onRemoveRow(id)} aria-label={`นำแถวจากชีต ${sheetName} ออกจากตาราง`} title="นำแถวนี้ออกจากตาราง"><X size={15} /></button></td>
                  <td>{sheetName}</td>
                  {headers.map((header) => <td key={header} title={String(row[header] ?? '')}>{String(row[header] ?? '') || <span className="empty-cell">—</span>}</td>)}
                </tr>)}
                {!outputRows.length && <tr><td className="table-empty" colSpan={headers.length + 2}>ชีตนี้ยังไม่มีข้อมูล — เลือกแถวจากพรีวิวแล้วกดปุ่ม +</td></tr>}
              </tbody>
            </table>
          </div>
          {outputRows.length > 0 && <TablePagination page={Math.min(outputPage, outputPageCount - 1)} pageSize={FILE_PREVIEW_PAGE_SIZE} total={outputRows.length} onPageChange={setOutputPage} />}
        </section>
      </>}
      {isChoosingOutputSheet && <div className="confirm-overlay" onMouseDown={(event) => {
        if (event.target === event.currentTarget) setIsChoosingOutputSheet(false);
      }} onKeyDown={(event) => {
        if (event.key === 'Escape') setIsChoosingOutputSheet(false);
      }}>
        <section className="confirm-dialog builder-add-dialog" role="dialog" aria-modal="true" aria-labelledby="builder-add-title" aria-describedby="builder-add-description">
          <span className="confirm-icon builder-add-icon"><Files size={20} /></span>
          <h2 id="builder-add-title">เลือกชีตปลายทาง</h2>
          <p id="builder-add-description">เพิ่มข้อมูลที่เลือก {selectedSourceRows.length.toLocaleString()} แถวลงในชีตที่ต้องการ</p>
          <label className="destination-sheet-select">
            <span>ชีตที่จะเพิ่มข้อมูล</span>
            <select value={batchOutputSheetId} onChange={(event) => setBatchOutputSheetId(event.target.value)}>
              {outputSheets.map((sheet) => <option key={sheet.id} value={sheet.id}>{sheet.name}</option>)}
            </select>
          </label>
          <div className="confirm-actions">
            <button className="confirm-cancel" autoFocus onClick={() => setIsChoosingOutputSheet(false)}>ยกเลิก</button>
            <button className="confirm-add" disabled={!batchOutputSheetId} onClick={confirmAddSelectedRows}><Plus size={16} /> เพิ่มข้อมูล</button>
          </div>
        </section>
      </div>}
      {addSuccess && <div key={addSuccess.id} className={`builder-add-success ${addSuccess.fading ? 'is-fading' : ''}`} role="status" aria-live="polite">
        <span className="builder-add-success-icon"><Check size={23} strokeWidth={2.8} /></span>
        <span className="builder-add-success-copy"><strong>เพิ่มข้อมูลเรียบร้อยแล้ว</strong><span>{addSuccess.count.toLocaleString()} แถว · ชีต “{addSuccess.sheetName}”</span></span>
      </div>}
    </>
  );
}

export default function App() {
  const [activePage, setActivePage] = useState('compare');
  const [files, setFiles] = useState([null, null]);
  const [errors, setErrors] = useState(['', '']);
  const [keySelections, setKeySelections] = useState({});
  const [activeSheets, setActiveSheets] = useState(['', '']);
  const [searches, setSearches] = useState(['', '']);
  const [result, setResult] = useState(null);
  const [resultSearch, setResultSearch] = useState('');
  const [resultSheetFilter, setResultSheetFilter] = useState('ทั้งหมด');
  const [removedIds, setRemovedIds] = useState(() => new Set());
  const [selectedIds, setSelectedIds] = useState(() => new Set());
  const [selectionMode, setSelectionMode] = useState(false);
  const [pendingRemovalIds, setPendingRemovalIds] = useState([]);
  const [builderFile, setBuilderFile] = useState(null);
  const [builderError, setBuilderError] = useState('');
  const [isBuilderReading, setIsBuilderReading] = useState(false);
  const [builderRows, setBuilderRows] = useState([]);
  const [builderSheets, setBuilderSheets] = useState([]);
  const [isReading, setIsReading] = useState(false);
  const [notice, setNotice] = useState('');
  const inputRefs = [useRef(null), useRef(null)];
  const builderRowInstance = useRef(0);
  const ready = files.every(Boolean);

  function addBuilderRow(item) {
    builderRowInstance.current += 1;
    setBuilderRows((previous) => [...previous, {
      ...item,
      sourceId: item.id,
      id: `${item.id}::${builderRowInstance.current}`,
    }]);
  }

  async function handleBuilderFile(file) {
    setBuilderError('');
    const extension = file.name.split('.').pop()?.toLowerCase();
    if (!['xlsx', 'xls'].includes(extension)) {
      setBuilderError('กรุณาเลือกไฟล์ Excel นามสกุล .xlsx หรือ .xls');
      return;
    }
    if (file.size > FILE_LIMIT) {
      setBuilderError('ไฟล์มีขนาดเกิน 25 MB');
      return;
    }
    setIsBuilderReading(true);
    try {
      const parsed = await parseWorkbook(await file.arrayBuffer(), file.name);
      setBuilderFile(parsed);
      setBuilderRows([]);
      setBuilderSheets([{ id: `builder-sheet-${Date.now()}`, name: 'ตาราง 1' }]);
    } catch (error) {
      setBuilderError(error instanceof Error ? error.message : 'อ่านไฟล์ไม่สำเร็จ');
    } finally {
      setIsBuilderReading(false);
    }
  }

  async function exportBuilderRows() {
    if (!builderSheets.length) return;
    try {
      const headers = [...new Set(builderFile?.sheets.flatMap((sheet) => sheet.headers) || [])];
      const sourceHeader = headers.includes('ชีตต้นทาง') ? 'ชีตต้นทาง (แหล่งที่มา)' : 'ชีตต้นทาง';
      await downloadWorkbook(builderSheets.map((sheet) => ({
        name: sheet.name,
        rows: builderRows.filter((item) => item.outputSheetId === sheet.id).map(({ sheetName, row }) => ({
          [sourceHeader]: sheetName,
          ...Object.fromEntries(headers.map((header) => [header, row[header] ?? ''])),
        })),
      })), 'sheetwise-custom-table.xlsx');
      setNotice(`ส่งออกตารางที่ปรับแต่ง ${builderRows.length.toLocaleString()} แถวแล้ว`);
    } catch (error) {
      setNotice(`ส่งออกไฟล์ไม่สำเร็จ: ${error instanceof Error ? error.message : 'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ'}`);
    }
  }

  async function handleFile(index, file) {
    setNotice('');
    setErrors((previous) => previous.map((error, position) => position === index ? '' : error));
    const extension = file.name.split('.').pop()?.toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(extension)) {
      setErrors((previous) => previous.map((error, position) => position === index ? 'ไฟล์ประเภทนี้ยังไม่รองรับ' : error));
      return;
    }
    if (file.size > FILE_LIMIT) {
      setErrors((previous) => previous.map((error, position) => position === index ? 'ไฟล์มีขนาดเกิน 25 MB' : error));
      return;
    }
    setIsReading(true);
    try {
      const buffer = await file.arrayBuffer();
      const parsed = await parseWorkbook(buffer, file.name);
      setFiles((previous) => previous.map((current, position) => position === index ? parsed : current));
      setActiveSheets((previous) => previous.map((sheetName, position) => position === index
        ? (parsed.sheets.some((sheet) => sheet.name === sheetName) ? sheetName : parsed.sheets.find((sheet) => sheet.rows.length)?.name || parsed.sheets[0].name)
        : sheetName));
      setResult(null);
      setResultSearch('');
      setResultSheetFilter('ทั้งหมด');
      setRemovedIds(new Set());
      setSelectedIds(new Set());
      setPendingRemovalIds([]);
    } catch (error) {
      setErrors((previous) => previous.map((message, position) => position === index ? (error instanceof Error ? error.message : 'อ่านไฟล์ไม่สำเร็จ') : message));
    } finally {
      setIsReading(false);
    }
  }

  function clearFile(index) {
    setFiles((previous) => previous.map((file, position) => position === index ? null : file));
    setKeySelections({});
    setActiveSheets((previous) => previous.map((sheet, position) => position === index ? '' : sheet));
    setErrors((previous) => previous.map((error, position) => position === index ? '' : error));
    setResult(null);
    setResultSearch('');
    setResultSheetFilter('ทั้งหมด');
    setRemovedIds(new Set());
    setSelectedIds(new Set());
    setPendingRemovalIds([]);
  }

  const matchingSheetNames = ready ? getMatchingSheetNames(files[0], files[1]) : [];
  const sheetResults = ready && result ? result : [];
  const resultCounts = result ? sheetResults.reduce((counts, sheetResult) => ({
    matched: counts.matched + sheetResult.result.matches.length,
    onlyA: counts.onlyA + sheetResult.result.onlyA.length,
    onlyB: counts.onlyB + sheetResult.result.onlyB.length,
  }), {
    matched: files[0].sheets.filter((sheet) => !matchingSheetNames.includes(sheet.name)).reduce((count, sheet) => count + sheet.rows.length, 0),
    onlyA: 0,
    onlyB: files[1].sheets.filter((sheet) => !matchingSheetNames.includes(sheet.name)).reduce((count, sheet) => count + sheet.rows.length, 0),
  }) : null;

  function startCompare() {
    if (!ready) return;
    const invalidMapping = matchingSheetNames.find((name) => {
      const keys = keySelections[name];
      const sheetA = getSheet(files[0], name);
      const sheetB = getSheet(files[1], name);
      const [keyA, keyB] = getDefaultKeys(sheetA, sheetB, keys);
      return !sheetA.headers.includes(keyA) || !sheetB.headers.includes(keyB);
    });
    if (invalidMapping) return;
    const compared = matchingSheetNames.map((name) => {
      const sheetA = getSheet(files[0], name);
      const sheetB = getSheet(files[1], name);
      const [keyA, keyB] = getDefaultKeys(sheetA, sheetB, keySelections[name]);
      return { name, sheetA, sheetB, keyA, keyB, result: compareFiles(sheetA, sheetB, keyA, keyB) };
    });
    setResult(compared);
    setResultSearch('');
    setResultSheetFilter('ทั้งหมด');
    setRemovedIds(new Set());
    setSelectedIds(new Set());
    setPendingRemovalIds([]);
    setSelectionMode(false);
  }

  function toggleRow(id) {
    setSelectedIds((previous) => {
      const next = new Set(previous);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function togglePage(rows, allSelected) {
    setSelectedIds((previous) => {
      const next = new Set(previous);
      rows.forEach((row) => {
        if (allSelected) next.delete(row.id);
        else next.add(row.id);
      });
      return next;
    });
  }

  function removeRows(ids) {
    setRemovedIds((previous) => new Set([...previous, ...ids]));
    setSelectedIds((previous) => {
      const next = new Set(previous);
      ids.forEach((id) => next.delete(id));
      return next;
    });
  }

  function requestRemoveRows(ids) {
    if (ids.length) setPendingRemovalIds(ids);
  }

  function confirmRemoveRows() {
    removeRows(pendingRemovalIds);
    setPendingRemovalIds([]);
  }

  const combinedRows = result && ready
    ? createCombinedRows(sheetResults, files[0], files[1])
    : [];
  const visibleRows = combinedRows.filter((row) => !removedIds.has(row.id));
  const resultSheetNames = [...new Set(visibleRows.map((row) => String(row.values['ชีต'])))];
  const sheetFilteredRows = resultSheetFilter === 'ทั้งหมด'
    ? visibleRows
    : visibleRows.filter((row) => row.values['ชีต'] === resultSheetFilter);
  const displayRows = sheetFilteredRows.filter((row) => rowMatchesQuery(row, resultSearch.trim().toLocaleLowerCase()));
  const selectedRows = displayRows.filter((row) => selectedIds.has(row.id));

  async function exportVisibleRows() {
    if (!displayRows.length) return;
    try {
      await downloadWorkbook([{
        name: 'ข้อมูลผลลัพธ์',
        rows: displayRows.map((row) => row.values),
      }], 'sheetwise-comparison-results.xlsx');
      setNotice(`ส่งออกข้อมูล ${displayRows.length.toLocaleString()} แถวที่ตรงกับตัวกรองแล้ว`);
    } catch (error) {
      setNotice(`ส่งออกไฟล์ไม่สำเร็จ: ${error instanceof Error ? error.message : 'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ'}`);
    }
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand" href="#" aria-label="Sheetwise home">
          <span className="brand-mark"><Table2 size={20} strokeWidth={2.2} /></span>
          <span>sheetwise<span className="brand-dot">.</span></span>
        </a>
        <div className="side-label">WORKSPACE</div>
        <nav className="side-nav" aria-label="เมนูหลัก">
          <button className={`nav-link ${activePage === 'compare' ? 'active' : ''}`} onClick={() => setActivePage('compare')} aria-current={activePage === 'compare' ? 'page' : undefined}><span><ArrowLeftRight size={17} /></span> เปรียบเทียบไฟล์</button>
          <button className={`nav-link ${activePage === 'builder' ? 'active' : ''}`} onClick={() => setActivePage('builder')} aria-current={activePage === 'builder' ? 'page' : undefined}><span><Table2 size={17} /></span> สร้างตาราง</button>
        </nav>
        <div className="sidebar-bottom">
          <div className="privacy-card">
            <span className="privacy-icon"><ShieldCheck size={18} /></span>
            <strong>ข้อมูลของคุณปลอดภัย</strong>
            <p>ประมวลผลบนอุปกรณ์ของคุณ ไม่มีการอัปโหลดไฟล์ขึ้นเซิร์ฟเวอร์</p>
          </div>
          <div className="profile-row"><span className="avatar">S</span><div><strong>พื้นที่ทำงาน</strong><span>โหมดส่วนตัว</span></div><CircleHelp size={17} /></div>
        </div>
      </aside>

      <main className="main-content">
        <header className="topbar"><div className="breadcrumbs"><span>Workspace</span><span className="crumb-slash">/</span><strong>{activePage === 'compare' ? 'เปรียบเทียบไฟล์' : 'สร้างตาราง'}</strong></div><div className="topbar-right"><span className="private-pill"><span /> ประมวลผลในเครื่อง</span><button className="help-button" title="ช่วยเหลือ"><CircleHelp size={19} /></button></div></header>
        <div className="page-wrap">
          {activePage === 'builder' ? <TableBuilderPage
            file={builderFile}
            fileError={builderError}
            isReading={isBuilderReading}
            onFile={handleBuilderFile}
            onClear={() => { setBuilderFile(null); setBuilderRows([]); setBuilderSheets([]); setBuilderError(''); }}
            addedRows={builderRows}
            outputSheets={builderSheets}
            onAddRow={addBuilderRow}
            onRemoveRow={(id) => setBuilderRows((previous) => previous.filter((row) => row.id !== id))}
            onAddSheet={(sheet) => setBuilderSheets((previous) => [...previous, sheet])}
            onRenameSheet={(id, name) => setBuilderSheets((previous) => previous.map((sheet) => sheet.id === id ? { ...sheet, name } : sheet))}
            onExport={exportBuilderRows}
          /> : <>
          <section className="intro">
            <div className="eyebrow"><span className="eyebrow-dot" /> DATA WORKSPACE</div>
            <h1>เปรียบเทียบข้อมูล<br className="mobile-break" /> <span>ได้อย่างมั่นใจ</span></h1>
            <p>ค้นหาความแตกต่างระหว่างไฟล์ของคุณได้ในไม่กี่ขั้นตอน<br className="desktop-break" /> ข้อมูลทั้งหมดอยู่บนอุปกรณ์ของคุณ ปลอดภัยและเป็นส่วนตัว</p>
            <div className="step-track">
              <div className="step active"><span>1</span><div><strong>อัปโหลดไฟล์</strong><small>เลือกข้อมูล 2 ชุด</small></div></div>
              <span className="step-line" />
              <div className={`step ${ready ? 'active' : ''}`}><span>2</span><div><strong>เลือกคอลัมน์</strong><small>กำหนด Key สำหรับเทียบ</small></div></div>
              <span className="step-line" />
              <div className={`step ${result ? 'active' : ''}`}><span>3</span><div><strong>ดูผลลัพธ์</strong><small>เปรียบเทียบและส่งออก</small></div></div>
            </div>
          </section>

          <section className="section-block upload-section">
            <div className="section-heading"><div><div className="section-kicker">STEP 01 <span>·</span> UPLOAD</div><h2>เลือกไฟล์ที่ต้องการเปรียบเทียบ</h2><p>รองรับ Excel หลายชีตและ CSV — ระบบจับคู่ชีตที่ชื่อเหมือนกัน</p></div><span className="file-count"><Files size={15} /> 2 ไฟล์</span></div>
            <div className="upload-grid">
              <div className="upload-column"><div className="upload-label"><span className="file-number blue-bg">01</span><div><strong>ไฟล์ต้นฉบับ</strong><span>ข้อมูลชุดแรกที่ต้องการตรวจสอบ</span></div></div><DropZone index={0} file={files[0]} onFile={handleFile} onClear={() => clearFile(0)} error={errors[0]} inputRef={inputRefs[0]} /></div>
              <div className="swap-symbol"><ArrowLeftRight size={16} /></div>
              <div className="upload-column"><div className="upload-label"><span className="file-number purple-bg">02</span><div><strong>ไฟล์ที่ต้องการเปรียบเทียบ</strong><span>ข้อมูลชุดที่สองสำหรับนำมาเทียบ</span></div></div><DropZone index={1} file={files[1]} onFile={handleFile} onClear={() => clearFile(1)} error={errors[1]} inputRef={inputRefs[1]} /></div>
            </div>
            <div className="cleaning-note"><span className="sparkle"><Sparkles size={15} /></span><span><strong>จัดระเบียบข้อมูลอัตโนมัติ</strong> <i>แถวว่างถูกลบ · ช่องว่างถูกตัด · หัวตารางที่ว่างถูกตั้งชื่อให้</i></span><button title="การจัดการข้อมูล"><Info size={15} /></button></div>
          </section>

          {isReading && <div className="reading-indicator"><LoaderCircle size={16} className="spin" /> กำลังอ่านและจัดระเบียบไฟล์ในเบราว์เซอร์...</div>}

          {files.some(Boolean) && <>
            <section className="section-block preview-section">
              <div className="section-heading"><div><div className="section-kicker">STEP 02 <span>·</span> REVIEW & MAP</div><h2>{ready ? 'ตรวจสอบข้อมูลและเลือก Key' : 'พรีวิวข้อมูลที่อัปโหลด'}</h2><p>{ready ? 'เลือกคอลัมน์ที่ระบุรายการเดียวกันในทั้งสองไฟล์' : 'ตรวจสอบข้อมูลในตารางได้ทันที เมื่ออัปโหลดไฟล์แต่ละชุด'}</p></div>{ready && <span className="preview-tag"><Check size={14} /> Preview พร้อมแล้ว</span>}</div>
              {ready && <div className="key-mapping">
                <div className="mapping-copy"><span className="mapping-icon"><ArrowLeftRight size={17} /></span><div><strong>จับคู่ชีตที่ชื่อเหมือนกัน</strong><span>ตั้งค่า Key แยกสำหรับแต่ละคู่ชีต; ชีตที่ไม่มีชื่อคู่กันจะแสดงเป็นรายการเฉพาะไฟล์</span></div></div>
                <div className="sheet-mappings">
                  {matchingSheetNames.length ? matchingSheetNames.map((name) => {
                    const sheetA = getSheet(files[0], name);
                    const sheetB = getSheet(files[1], name);
                    const keys = getDefaultKeys(sheetA, sheetB, keySelections[name]);
                    return <div className="sheet-mapping-row" key={name}>
                      <strong className="sheet-mapping-name" title={name}>{name}</strong>
                      {[sheetA, sheetB].map((sheet, index) => <label className="select-wrap" key={`${name}-${index}`}><span className={`select-label ${index ? 'purple-text' : ''}`}>ไฟล์ {index + 1}</span><select value={keys[index]} onChange={(event) => {
                        setKeySelections((previous) => ({
                          ...previous,
                          [name]: getDefaultKeys(sheetA, sheetB, previous[name]).map((key, position) => position === index ? event.target.value : key),
                        }));
                        setResult(null);
                        setRemovedIds(new Set());
                        setSelectedIds(new Set());
                      }}>{sheet.headers.map((header) => <option key={header} value={header}>{header}</option>)}</select><ChevronDown size={15} /></label>)}
                    </div>;
                  }) : <p className="no-matching-sheets">ไม่พบชีตที่มีข้อมูลและชื่อเหมือนกัน ระบบจะแสดงข้อมูลจากทุกชีตเป็นรายการเฉพาะของแต่ละไฟล์</p>}
                </div>
              </div>}
              <div className={`preview-grid ${files.filter(Boolean).length === 1 ? 'single-preview' : ''}`}>
                {files.map((file, index) => file && <DataTable
                  key={`${file.filename}-${index}`}
                  file={file}
                  sheet={getSheet(file, activeSheets[index]) || file.sheets[0]}
                  fileIndex={index}
                  activeSheet={activeSheets[index] || file.sheets[0].name}
                  onSheetChange={(name) => setActiveSheets((previous) => previous.map((sheetName, position) => position === index ? name : sheetName))}
                  search={searches[index]}
                  setSearch={(value) => setSearches((previous) => previous.map((search, position) => position === index ? value : search))}
                />)}
              </div>
            </section>
          </>}

          {ready && <section className="compare-bar">
              <div className="compare-bar-copy"><span className="compare-icon"><ArrowLeftRight size={18} /></span><div><strong>พร้อมเริ่มเปรียบเทียบ?</strong><span>ไฟล์ของคุณจะถูกประมวลผลบนอุปกรณ์นี้เท่านั้น</span></div></div>
              <button className="compare-button" onClick={startCompare}><ArrowLeftRight size={16} /> เริ่มเปรียบเทียบ <ArrowUpRight size={15} /></button>
            </section>}

          {result && <section className="section-block results-section">
            <div className="section-heading results-heading"><div><div className="section-kicker">STEP 03 <span>·</span> COMPARISON RESULTS</div><h2>หน้าต่างผลลัพธ์</h2><p>เปรียบเทียบชีตที่ชื่อเหมือนกันทุกชีต แถวที่จับคู่ได้เรียงไฟล์ 1 ก่อน แล้วตามด้วยไฟล์ 2</p></div><div className="result-actions">
              <button className={`select-mode-button ${selectionMode ? 'active' : ''}`} onClick={() => {
                setSelectionMode((enabled) => !enabled);
                setSelectedIds(new Set());
              }} aria-pressed={selectionMode}>{selectionMode ? <X size={15} /> : <Check size={15} />} {selectionMode ? 'ยกเลิกการเลือก' : 'เลือกข้อมูล'}</button>
              <button className="remove-selected-button" disabled={!selectedRows.length} onClick={() => requestRemoveRows(selectedRows.map((row) => row.id))}><X size={15} /> นำข้อมูลที่เลือกออก ({selectedRows.length.toLocaleString()})</button>
              <button className="export-button" disabled={!displayRows.length} onClick={exportVisibleRows}><Download size={15} /> ส่งออก Excel</button>
            </div></div>
            <div className="result-summary">
              <div className="summary-intro"><span className="summary-check"><Check size={19} /></span><div><strong>เปรียบเทียบเสร็จแล้ว</strong><span>เหลือข้อมูลในตาราง {visibleRows.length.toLocaleString()} จาก {combinedRows.length.toLocaleString()} แถว</span></div></div>
              <div className="summary-stat"><strong>{resultCounts.matched.toLocaleString()}</strong><span>คู่ที่จับคู่ได้</span></div>
              <div className="summary-stat"><strong>{resultCounts.onlyA.toLocaleString()}</strong><span>เฉพาะไฟล์ 1</span></div>
              <div className="summary-stat"><strong>{resultCounts.onlyB.toLocaleString()}</strong><span>เฉพาะไฟล์ 2</span></div>
            </div>
            <div className="results-card">
              <div className="result-table-heading"><div><strong>ข้อมูลรวมจากทั้งสองไฟล์</strong><span>{selectionMode ? 'เลือกแถวที่ต้องการนำออกจากตารางด้วยช่องทำเครื่องหมาย' : 'ใช้ปุ่ม “เลือกข้อมูล” เพื่อเลือกหลายแถว · กด × ท้ายแถวเพื่อนำออกทันที'}</span></div><span className="selection-count">{selectedRows.length.toLocaleString()} รายการที่เลือก</span></div>
              <ResultTable
                key={resultSheetFilter}
                rows={displayRows}
                allRows={visibleRows}
                sheetNames={resultSheetNames}
                sheetFilter={resultSheetFilter}
                onSheetFilterChange={setResultSheetFilter}
                search={resultSearch}
                onSearchChange={setResultSearch}
                selectionMode={selectionMode}
                selectedIds={selectedIds}
                onToggleRow={toggleRow}
                onTogglePage={togglePage}
                onRemoveRow={(id) => requestRemoveRows([id])}
              />
              <div className="result-footnote"><Info size={14} /><span>การนำแถวออกจะซ่อนแถวนั้นจากตารางและไฟล์ส่งออกปัจจุบัน โดยไม่แก้ไขไฟล์ต้นฉบับ</span></div>
            </div>
          </section>}
          </>}

          <footer className="page-footer"><span>© 2025 Sheetwise</span><span><ShieldCheck size={14} /> ไฟล์ของคุณไม่ถูกส่งออกจากอุปกรณ์</span><a href="mailto:support@sheetwise.local">ต้องการความช่วยเหลือ?</a></footer>
        </div>
      </main>
      {pendingRemovalIds.length > 0 && <div className="confirm-overlay" onMouseDown={(event) => {
        if (event.target === event.currentTarget) setPendingRemovalIds([]);
      }} onKeyDown={(event) => {
        if (event.key === 'Escape') setPendingRemovalIds([]);
      }}>
        <section className="confirm-dialog" role="dialog" aria-modal="true" aria-labelledby="confirm-title" aria-describedby="confirm-description">
          <span className="confirm-icon"><X size={20} /></span>
          <h2 id="confirm-title">ยืนยันการนำข้อมูลออก</h2>
          <p id="confirm-description">ต้องการนำ {pendingRemovalIds.length.toLocaleString()} แถวออกจากตารางผลลัพธ์หรือไม่? ข้อมูลนี้จะไม่ถูกรวมในไฟล์ Excel ที่ส่งออก แต่ไฟล์ต้นฉบับจะไม่เปลี่ยนแปลง</p>
          <div className="confirm-actions">
            <button className="confirm-cancel" autoFocus onClick={() => setPendingRemovalIds([])}>ยกเลิก</button>
            <button className="confirm-delete" onClick={confirmRemoveRows}><X size={15} /> ยืนยันนำออก</button>
          </div>
        </section>
      </div>}
      {notice && <div className="toast"><Info size={16} />{notice}<button onClick={() => setNotice('')} aria-label="ปิด"><X size={15} /></button></div>}
    </div>
  );
}
