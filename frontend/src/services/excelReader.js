import * as XLSX from 'xlsx';

const BLANK = new Set([null, undefined, '', '#N/A', '#VALUE!', '#REF!', '#DIV/0!', '0', '0.0']);

function val(ws, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  const cell = ws[addr];
  if (!cell) return '';
  const s = String(cell.v ?? '').trim();
  return BLANK.has(s) ? '' : s;
}

function dateVal(ws, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  const cell = ws[addr];
  if (!cell) return '';
  if (cell.t === 'd') {
    const d = new Date(cell.v);
    return d.toISOString().slice(0, 10);
  }
  if (cell.t === 'n' && cell.w) return cell.w;
  if (cell.v == null) return '';
  return String(cell.v).trim();
}

function num(ws, row, col, decimals = 2) {
  const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  const cell = ws[addr];
  if (!cell) return '';
  const n = parseFloat(cell.v);
  if (isNaN(n)) return String(cell.v ?? '').trim();
  return n.toFixed(decimals);
}

export function readVibAnalyzer(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const cert = wb.Sheets['Vib Analyzer Cert'];
  if (!cert) throw new Error('Sheet "Vib Analyzer Cert" not found in workbook');

  const info = {};

  // Analyzer / Meter Information
  info.manufacturer = val(cert, 5, 3);
  info.model = val(cert, 6, 3);
  info.serial = val(cert, 7, 3);
  info.cal_tech = val(cert, 8, 3);
  info.cal_date = val(cert, 9, 3);
  info.cal_due = val(cert, 10, 3);

  // Analyzer Settings
  info.analyzer_mode = val(cert, 5, 8);
  info.freq_max = val(cert, 6, 8);
  info.freq_min = val(cert, 7, 8);
  info.freq_unit = val(cert, 8, 8);
  info.lines_resolution = val(cert, 9, 8);
  info.avg_points = val(cert, 10, 8);
  info.window_type = val(cert, 11, 8);
  info.sensor_input_sens = val(cert, 12, 8);
  info.sensor_input_unit = val(cert, 12, 9);

  // Test type detection (J4 header)
  info.test_type = val(cert, 4, 10);

  // Test parameters (column L-M, rows 5-8)
  info.test_params = {};
  for (let r = 5; r <= 8; r++) {
    const label = val(cert, r, 12);
    const value = val(cert, r, 13);
    if (label && value) info.test_params[label] = value;
  }

  // Test Equipment
  info.pvc_model = val(cert, 15, 3);
  info.pvc_serial = val(cert, 15, 4);
  info.pvc_cal_date = dateVal(cert, 15, 6);
  info.pvc_sensitivity = num(cert, 15, 8);
  info.pvc_sens_unit = val(cert, 15, 9);
  info.pvc_tolerance = num(cert, 15, 10);
  info.pvc_deviation = num(cert, 15, 12);

  info.sensor_model = val(cert, 16, 3);
  info.sensor_serial = val(cert, 16, 4);
  info.sensor_cal_date = dateVal(cert, 16, 6);
  info.sensor_sensitivity = num(cert, 16, 8);
  info.sensor_sens_unit = val(cert, 16, 9);
  info.sensor_tolerance = num(cert, 16, 10);
  info.sensor_deviation = num(cert, 16, 12);

  // Customer
  info.customer = val(cert, 33, 7);

  // Note
  info.note = val(cert, 41, 7);

  return info;
}

export function mergeInfo(accInfo, velInfo) {
  const merged = {};
  const sources = [accInfo, velInfo].filter(Boolean);
  if (!sources.length) return merged;

  const commonKeys = [
    'manufacturer', 'model', 'serial', 'cal_tech', 'cal_date', 'cal_due',
    'analyzer_mode', 'freq_max', 'freq_min', 'freq_unit', 'lines_resolution',
    'avg_points', 'window_type', 'sensor_input_sens', 'sensor_input_unit',
    'customer', 'pvc_model', 'pvc_serial', 'pvc_cal_date', 'pvc_sensitivity',
    'pvc_sens_unit', 'pvc_tolerance', 'pvc_deviation', 'sensor_model',
    'sensor_serial', 'sensor_cal_date', 'sensor_sensitivity', 'sensor_sens_unit',
    'sensor_tolerance', 'sensor_deviation', 'note',
  ];

  for (const key of commonKeys) {
    merged[key] = '';
    for (const src of sources) {
      if (src[key]) { merged[key] = src[key]; break; }
    }
  }

  // Test-specific info kept separate
  if (accInfo) {
    merged.acc_test_type = accInfo.test_type || 'Linearity Test';
    merged.acc_params = accInfo.test_params || {};
  } else {
    merged.acc_test_type = '';
    merged.acc_params = {};
  }

  if (velInfo) {
    merged.vel_test_type = velInfo.test_type || 'Frequency Response Test';
    merged.vel_params = velInfo.test_params || {};
  } else {
    merged.vel_test_type = '';
    merged.vel_params = {};
  }

  return merged;
}
