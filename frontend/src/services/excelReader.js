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
  const data = wb.Sheets['Vib Analyzer Data'];
  const cert = wb.Sheets['Vib Analyzer Cert'];
  if (!data && !cert) throw new Error('Neither "Vib Analyzer Data" nor "Vib Analyzer Cert" sheet found');

  // Prefer Data sheet (has actual values); Cert sheet has formulas that may not be cached
  const src = data || cert;

  // Lookup tables for coded fields on the Data sheet
  const MODE_MAP = { '1': 'Spectra', '2': 'Overall', '3': 'Time waveform' };
  const FREQ_UNIT_MAP = { '1': 'Hz', '2': 'CPM' };
  const WINDOW_MAP = { '1': 'Hanning', '2': 'Hamming', '3': 'Flattop', '4': 'Uniform' };
  const SENS_UNIT_MAP = { '1': 'mV/EU', '2': 'V/EU' };

  const info = {};

  // Analyzer / Meter Information (same positions on both sheets)
  info.manufacturer = val(src, 5, 3);
  info.model = val(src, 6, 3);
  info.serial = val(src, 7, 3);
  info.cal_tech = val(src, 8, 3);
  info.cal_date = val(src, 9, 3);
  info.cal_due = val(src, 10, 3);

  // Analyzer Settings
  if (data) {
    // Data sheet stores numeric codes for some fields
    const modeRaw = val(src, 5, 8);
    info.analyzer_mode = MODE_MAP[modeRaw] || modeRaw;
    info.freq_max = val(src, 6, 8);
    info.freq_min = val(src, 7, 8);
    const funitRaw = val(src, 8, 8);
    info.freq_unit = FREQ_UNIT_MAP[funitRaw] || funitRaw;
    info.lines_resolution = val(src, 9, 8);
    info.avg_points = val(src, 10, 8);
    const winRaw = val(src, 11, 8);
    info.window_type = WINDOW_MAP[winRaw] || winRaw;
    info.sensor_input_sens = val(src, 12, 8);
    const sunitRaw = val(src, 13, 8);
    info.sensor_input_unit = SENS_UNIT_MAP[sunitRaw] || sunitRaw;
  } else {
    info.analyzer_mode = val(src, 5, 8);
    info.freq_max = val(src, 6, 8);
    info.freq_min = val(src, 7, 8);
    info.freq_unit = val(src, 8, 8);
    info.lines_resolution = val(src, 9, 8);
    info.avg_points = val(src, 10, 8);
    info.window_type = val(src, 11, 8);
    info.sensor_input_sens = val(src, 12, 8);
    info.sensor_input_unit = val(src, 12, 9);
  }

  // Test type detection (J4 header)
  info.test_type = val(cert || src, 4, 10);

  // Test parameters (column L-M, rows 5-8 on Cert sheet)
  info.test_params = {};
  const paramSrc = cert || src;
  for (let r = 5; r <= 8; r++) {
    const label = val(paramSrc, r, 12);
    const value = val(paramSrc, r, 13);
    if (label && value) info.test_params[label] = value;
  }

  // Test Equipment
  if (data) {
    // Data sheet: PVC at row 16, Sensor at row 17; sensitivity at col G(7), unit at H(8), tolerance at K(11)
    info.pvc_model = val(src, 16, 3);
    info.pvc_serial = val(src, 16, 4);
    info.pvc_cal_date = dateVal(src, 16, 6);
    info.pvc_sensitivity = num(src, 16, 7);
    info.pvc_sens_unit = val(src, 16, 8);
    info.pvc_tolerance = num(src, 16, 11);
    info.pvc_deviation = num(src, 16, 12);

    info.sensor_model = val(src, 17, 3);
    info.sensor_serial = val(src, 17, 4);
    info.sensor_cal_date = dateVal(src, 17, 6);
    info.sensor_sensitivity = num(src, 17, 7);
    info.sensor_sens_unit = val(src, 17, 8);
    info.sensor_tolerance = num(src, 17, 11);
    info.sensor_deviation = num(src, 17, 12);
  } else {
    // Cert sheet: PVC at row 15, Sensor at row 16
    info.pvc_model = val(src, 15, 3);
    info.pvc_serial = val(src, 15, 4);
    info.pvc_cal_date = dateVal(src, 15, 6);
    info.pvc_sensitivity = num(src, 15, 8);
    info.pvc_sens_unit = val(src, 15, 9);
    info.pvc_tolerance = num(src, 15, 10);
    info.pvc_deviation = num(src, 15, 12);

    info.sensor_model = val(src, 16, 3);
    info.sensor_serial = val(src, 16, 4);
    info.sensor_cal_date = dateVal(src, 16, 6);
    info.sensor_sensitivity = num(src, 16, 8);
    info.sensor_sens_unit = val(src, 16, 9);
    info.sensor_tolerance = num(src, 16, 10);
    info.sensor_deviation = num(src, 16, 12);
  }

  // Customer & Note (literal values on Cert sheet)
  info.customer = val(cert || src, 33, 7);
  info.note = val(cert || src, 41, 7);

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
