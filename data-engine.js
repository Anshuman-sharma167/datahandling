(function (global) {
  'use strict';

  var fields = ['name', 'date', 'program', 'value'];
  var nextId = 0;

  function blank(value) {
    return value == null || (typeof value === 'string' && value.trim() === '');
  }

  function normalized(value) {
    return blank(value) ? '' : String(value).trim().replace(/\s+/g, ' ').toLowerCase();
  }

  function headerField(header) {
    var text = normalized(header).replace(/[_-]+/g, ' ');
    if (text.indexOf('name') !== -1) return 'name';
    if (text.indexOf('date') !== -1) return 'date';
    if (text.indexOf('program') !== -1 || text.indexOf('activity') !== -1) return 'program';
    if (text.indexOf('value') !== -1 || text.indexOf('amount') !== -1 || text.indexOf('count') !== -1) return 'value';
    return null;
  }

  function tableRows(table, file, format, context) {
    if (!table.length) return [];
    if (context) table[0].forEach(function (header) {
      if (blank(header)) return;
      var label = String(header).trim().slice(0, 120);
      if (context.headers.indexOf(label) === -1 && context.headers.length < 50) context.headers.push(label);
    });
    var mapping = table[0].map(headerField);
    return table.slice(1).filter(function (cells) {
      return cells.some(function (cell) { return !blank(cell); });
    }).map(function (cells) {
      var row = {
        id: 'row-' + (++nextId), name: null, date: null, program: null,
        value: null, source_file: file.name, source_format: format
      };
      mapping.forEach(function (field, index) {
        if (field && blank(row[field]) && !blank(cells[index])) row[field] = cells[index];
      });
      return row;
    });
  }

  function parseDelimited(text) {
    var lines = text.split(/\r?\n/).filter(function (line) { return line.trim(); });
    if (!/[\t,;|]/.test(text) && lines.some(function (line) { return /\S\s{2,}\S/.test(line); })) {
      return lines.map(function (line) { return line.trim().split(/\s{2,}/); });
    }
    var result = global.Papa.parse(text, { skipEmptyLines: 'greedy' });
    var table = result.data || [];
    if (table.length && table[0].length === 1) {
      table = text.split(/\r?\n/).filter(function (line) { return line.trim(); })
        .map(function (line) { return line.trim().split(/\s{2,}/); });
    }
    return table;
  }

  function pdfLines(items) {
    var lines = [];
    items.forEach(function (item) {
      if (typeof item.str !== 'string' || !item.str.trim()) return;
      var transform = item.transform || [];
      var y = Number(transform[5]) || 0;
      var line = lines.find(function (entry) { return Math.abs(entry.y - y) <= 2; });
      if (!line) {
        line = { y: y, cells: [] };
        lines.push(line);
      }
      line.cells.push({ x: Number(transform[4]) || 0, width: Number(item.width) || 0, fontSize: Math.abs(Number(transform[0])) || 10, text: item.str });
    });
    return lines.sort(function (a, b) { return b.y - a.y; }).map(function (line) {
      var previous = null;
      return line.cells.sort(function (a, b) { return a.x - b.x; }).map(function (cell) {
        var separator = '';
        if (previous) separator = cell.x - previous.x - previous.width > Math.max(6, previous.fontSize * 0.8) ? '  ' : ' ';
        previous = cell;
        return separator + cell.text;
      }).join('').trim();
    }).filter(Boolean);
  }

  async function parseFile(file, context) {
    var extension = String(file.name || '').split('.').pop().toLowerCase();
    if (extension === 'csv') {
      return new Promise(function (resolve, reject) {
        global.Papa.parse(file, {
          skipEmptyLines: 'greedy',
          complete: function (result) { resolve(tableRows(result.data || [], file, extension, context)); },
          error: reject
        });
      });
    }
    if (extension === 'xlsx' || extension === 'xls') {
      var workbook = global.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
      context.sheets = workbook.SheetNames.slice(0, 50);
      return workbook.SheetNames.reduce(function (rows, sheetName) {
        var table = global.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
          header: 1, defval: null, blankrows: false, raw: true
        });
        var date1904 = !!(workbook.Workbook && workbook.Workbook.WBProps && workbook.Workbook.WBProps.date1904);
        if (table.length) {
          table[0].forEach(function (header, column) {
            if (headerField(header) !== 'date') return;
            table.slice(1).forEach(function (cells) {
              if (typeof cells[column] !== 'number') return;
              var decoded = global.XLSX.SSF.parse_date_code(cells[column], { date1904: date1904 });
              cells[column] = decoded ? isoDate(decoded.y, decoded.m, decoded.d) : null;
            });
          });
        }
        return rows.concat(tableRows(table, file, extension, context));
      }, []);
    }
    if (extension === 'pdf') {
      var task = global.pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) });
      var document = await task.promise;
      var output = [];
      var hasText = false;
      var header = null;
      try {
        for (var pageNumber = 1; pageNumber <= document.numPages; pageNumber++) {
          var page = await document.getPage(pageNumber);
          var content = await page.getTextContent();
          var lines = pdfLines(content.items);
          if (lines.some(function (line) { return line.trim(); })) hasText = true;
          var table = parseDelimited(lines.join('\n'));
          var headerIndex = table.findIndex(function (cells) {
            return cells.filter(function (cell) { return headerField(cell); }).length >= 2;
          });
          if (headerIndex >= 0) {
            header = table[headerIndex];
            table = table.slice(headerIndex);
          } else if (header) {
            table.unshift(header);
          } else {
            continue;
          }
          table = [table[0]].concat(table.slice(1).filter(function (cells) {
            return !cells.every(function (cell, index) { return normalized(cell) === normalized(header[index]); });
          }));
          output = output.concat(tableRows(table, file, extension, context));
        }
        if (!hasText) return [{ unsupportedFile: true, reason: 'scanned document, OCR not supported', source_file: file.name, source_format: extension }];
        return output;
      } finally {
        await task.destroy();
      }
    }
    return [];
  }

  function validRows(rows) {
    return Array.from(rows || []).filter(function (row) { return row && !row.unsupportedFile; });
  }

  function isoDate(year, month, day) {
    year = Number(year);
    month = Number(month);
    day = Number(day);
    if (year < 100) year += year < 70 ? 2000 : 1900;
    if (year < 100 || year > 9999 || month < 1 || month > 12 || day < 1) return null;
    var date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
    return date.toISOString().slice(0, 10);
  }

  function cleanDate(value) {
    if (value instanceof Date) return Number.isFinite(value.getTime()) ? value.toISOString().slice(0, 10) : null;
    if (typeof value === 'number' && Number.isFinite(value)) {
      var decoded = global.XLSX.SSF.parse_date_code(value);
      return decoded ? isoDate(decoded.y, decoded.m, decoded.d) : null;
    }
    var text = String(value).trim();
    var match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T\s].*)?$/);
    if (match) return isoDate(match[1], match[2], match[3]);
    match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/);
    if (match) {
      var first = Number(match[1]);
      var second = Number(match[2]);
      return first > 12 ? isoDate(match[3], second, first) : isoDate(match[3], first, second);
    }
    var months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    match = text.match(/^(\d{1,2})[\s-]+([a-z]+)[\s,\-]+(\d{2}|\d{4})$/i);
    if (match) return isoDate(match[3], months.indexOf(match[2].slice(0, 3).toLowerCase()) + 1, match[1]);
    match = text.match(/^([a-z]+)[\s-]+(\d{1,2}),?[\s-]+(\d{2}|\d{4})$/i);
    if (match) return isoDate(match[3], months.indexOf(match[1].slice(0, 3).toLowerCase()) + 1, match[2]);
    return null;
  }

  function numeric(value) {
    if (blank(value)) return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    var text = String(value).trim().replace(/^[\p{Sc}]\s*|\s*[\p{Sc}]$/gu, '').replace(/,/g, '');
    if (/^\(.*\)$/.test(text)) text = '-' + text.slice(1, -1);
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) return null;
    var number = Number(text);
    return Number.isFinite(number) ? number : null;
  }

  /** Parse supported files into canonical rows, including markers for scanned PDFs. */
  async function parseFiles(fileList) {
    var result = [];
    var documents = [];
    for (var file of Array.from(fileList || [])) {
      var context = { title: String(file.name || '').slice(0, 200), format: String(file.name || '').split('.').pop().toLowerCase(), headers: [], sheets: [] };
      var rows = await parseFile(file, context);
      context.recordCount = rows.filter(function (row) { return !row.unsupportedFile; }).length;
      if (context.recordCount) documents.push(context);
      result = result.concat(rows);
    }
    Object.defineProperty(result, 'documentContext', { value: documents, enumerable: false });
    return result;
  }

  /** Clean canonical fields and return an audit log without mutating input rows. */
  function sanitize(rows) {
    var log = [];
    var cleanedRows = validRows(rows).map(function (original) {
      var row = Object.assign({}, original);
      Object.keys(row).forEach(function (key) {
        if (typeof row[key] === 'string') row[key] = row[key].trim();
      });
      fields.forEach(function (field) {
        var value = original[field];
        var cleaned = blank(value) ? null : typeof value === 'string' ? value.trim() : value;
        var reason = 'standardized field';
        if (cleaned === null && field !== 'program') reason = 'missing field';
        else if (cleaned !== null && field === 'name') {
          cleaned = String(cleaned).toLowerCase().replace(/(^|[\s\-'’])\p{L}/gu, function (letter) { return letter.toUpperCase(); });
        } else if (cleaned !== null && field === 'date') {
          cleaned = cleanDate(cleaned);
          if (cleaned === null) reason = 'invalid date';
        } else if (cleaned !== null && field === 'value') {
          cleaned = numeric(cleaned);
          if (cleaned === null) reason = 'invalid numeric value';
        }
        row[field] = cleaned;
        if (reason === 'missing field' || value !== cleaned) {
          log.push({ rowId: row.id, field: field, originalValue: value == null ? null : value, cleanedValue: cleaned, reason: reason });
        }
      });
      return row;
    });
    return { rows: cleanedRows, log: log };
  }

  /** Merge exact and corroborated fuzzy matches, returning merge and review logs. */
  function mergeDuplicates(rows) {
    var output = [];
    var mergeLog = [];
    var possibleDuplicates = [];
    validRows(rows).forEach(function (original) {
      var row = Object.assign({}, original);
      var name = normalized(row.name);
      var date = normalized(row.date);
      var program = normalized(row.program);
      var kept = name && date ? output.find(function (candidate) {
        return normalized(candidate.name) === name && normalized(candidate.date) === date;
      }) : null;
      var matchType = 'exact';
      var confidence = 1;
      if (!kept && name && output.length) {
        var candidates = output.filter(function (candidate) { return !blank(candidate.name); })
          .map(function (candidate) { return { row: candidate, name: normalized(candidate.name) }; });
        var fuse = new global.Fuse(candidates, { keys: ['name'], threshold: 0.3, includeScore: true, ignoreLocation: true });
        var matches = fuse.search(name).filter(function (match) { return match.score <= 0.3; });
        matches.forEach(function (match) {
          var candidate = match.item.row;
          var corroborated = (date && normalized(candidate.date) === date) || (program && normalized(candidate.program) === program);
          if (corroborated && !kept) {
            kept = candidate;
            matchType = 'fuzzy';
            confidence = 1 - match.score;
          } else if (!corroborated) {
            possibleDuplicates.push({ rowIds: [candidate.id, row.id], reason: 'Similar names without a matching date or program' });
          }
        });
      }
      if (!kept) {
        output.push(row);
        return;
      }
      var reconciled = [];
      fields.concat(['source_file', 'source_format']).forEach(function (field) {
        if (blank(kept[field]) && !blank(row[field])) {
          kept[field] = row[field];
          reconciled.push(field);
        }
      });
      mergeLog.push({ keptRowId: kept.id, mergedFromRowIds: [row.id], matchType: matchType, confidence: confidence, fieldsReconciled: reconciled });
    });
    return { rows: output, mergeLog: mergeLog, possibleDuplicates: possibleDuplicates };
  }

  /** Flag numeric values outside two population standard deviations within each program. */
  function detectOutliers(rows) {
    var groups = new Map();
    validRows(rows).forEach(function (row) {
      var value = numeric(row.value);
      if (value === null) return;
      var key = normalized(row.program);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push({ row: row, value: value });
    });
    var outliers = [];
    groups.forEach(function (group) {
      if (group.length < 3) return;
      var mean = group.reduce(function (sum, item) { return sum + item.value / group.length; }, 0);
      var deviation = Math.sqrt(group.reduce(function (sum, item) {
        return sum + Math.pow(item.value - mean, 2) / group.length;
      }, 0));
      group.forEach(function (item) {
        if (item.value > mean + 2 * deviation || item.value < mean - 2 * deviation) {
          outliers.push({ rowId: item.row.id, field: 'value', value: item.value, reason: 'Value is more than two standard deviations from the program mean' });
        }
      });
    });
    return outliers;
  }

  /** Compute beneficiary and recorded-value metrics with contributing row IDs. */
  function computeMetrics(rows) {
    var input = validRows(rows);
    var excluded = new Set(detectOutliers(input).map(function (entry) { return entry.rowId; }));
    var beneficiaries = input.filter(function (row) { return !blank(row.name); });
    var recorded = input.filter(function (row) { return numeric(row.value) !== null && !excluded.has(row.id); });
    return [
      { key: 'uniqueBeneficiaries', label: 'Unique beneficiaries', value: new Set(beneficiaries.map(function (row) { return row.id; })).size,
        definition: 'Unique beneficiaries served — distinct people with at least one valid record.',
        sourceRowIds: Array.from(new Set(beneficiaries.map(function (row) { return row.id; }))) },
      { key: 'totalRecordedValue', label: 'Total recorded value', value: recorded.reduce(function (sum, row) { return sum + numeric(row.value); }, 0),
        definition: 'Total recorded program value — sum of value field, outliers excluded.',
        sourceRowIds: recorded.map(function (row) { return row.id; }) }
    ];
  }

  /** Build a summary containing only aggregate counts and metric metadata. */
  function buildDatasetSummary(rows, sanitizeLog, mergeLog, possibleDuplicates, outliers, metrics) {
    var input = validRows(rows);
    return {
      totalRows: input.length,
      fieldsUsed: fields.filter(function (field) { return input.some(function (row) { return !blank(row[field]); }); }),
      metrics: Array.from(metrics || []).map(function (metric) {
        return { key: metric.key, label: metric.label, value: metric.value, definition: metric.definition };
      }),
      mergeCount: new Set(Array.from(mergeLog || []).flatMap(function (entry) { return entry.mergedFromRowIds || []; })).size,
      possibleDuplicateCount: Array.from(possibleDuplicates || []).length,
      outlierCount: Array.from(outliers || []).length,
      missingFieldCount: Array.from(sanitizeLog || []).filter(function (entry) { return entry.reason === 'missing field'; }).length
    };
  }

  /** Return a badge color using inclusive green and yellow upper bounds. */
  function getBadgeLevel(value, thresholds) {
    if (value <= thresholds.greenMax) return 'green';
    if (value <= thresholds.yellowMax) return 'yellow';
    return 'red';
  }

  global.DataEngine = {
    parseFiles: parseFiles,
    sanitize: sanitize,
    mergeDuplicates: mergeDuplicates,
    detectOutliers: detectOutliers,
    computeMetrics: computeMetrics,
    buildDatasetSummary: buildDatasetSummary,
    getBadgeLevel: getBadgeLevel
  };
})(window);
