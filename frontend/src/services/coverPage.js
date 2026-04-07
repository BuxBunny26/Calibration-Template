import { jsPDF } from 'jspdf';
import 'jspdf-autotable';

const RED = [194, 4, 11];
const BLACK = [26, 26, 26];
const GREY = [77, 77, 77];
const LIGHT_GREY = [245, 245, 247];
const WHITE = [255, 255, 255];

const PAGE_W = 210; // A4 mm
const MARGIN = 15;
const CONTENT_W = PAGE_W - 2 * MARGIN;

function drawHeader(doc, info) {
  const pageH = doc.internal.pageSize.getHeight();

  // Try to draw logo (loaded as base64 before calling)
  if (info._logoBase64) {
    try {
      const logoH = 20;
      const logoW = logoH * 488 / 511;
      doc.addImage(info._logoBase64, 'PNG', MARGIN, 2, logoW, logoH);
    } catch { /* skip logo if failed */ }
  }

  // Header right
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.setTextColor(...BLACK);
  doc.text('CALIBRATION CERTIFICATE', PAGE_W - MARGIN, 8, { align: 'right' });

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(...GREY);
  doc.text('WearCheck ARC \u2014 Condition Monitoring Division', PAGE_W - MARGIN, 12, { align: 'right' });

  doc.setFontSize(6.5);
  doc.setTextColor(...BLACK);
  doc.text('ISO 16063-21  |  NMISA', PAGE_W - MARGIN, 15, { align: 'right' });

  // Footer
  const serial = info.serial || '';
  const calDate = info.cal_date || '';
  const mfr = info.manufacturer || '';
  const model = info.model || '';
  const genDateShort = new Date().toISOString().slice(0, 10);
  const certNum = serial ? `VIB-${serial}-${genDateShort}` : '\u2014';
  const equipLine = `${mfr} ${model}  |  S/N: ${serial}  |  Cert: ${certNum}`;

  doc.setDrawColor(200, 200, 200);
  doc.setLineWidth(0.2);
  doc.line(MARGIN, pageH - 14, PAGE_W - MARGIN, pageH - 14);

  doc.setFontSize(6.5);
  doc.setTextColor(...BLACK);
  doc.text(equipLine, MARGIN, pageH - 10);

  const now = new Date();
  const genDate = now.toISOString().slice(0, 16).replace('T', ' ');
  doc.text(`Generated: ${genDate}  |  Page ${doc.getCurrentPageInfo().pageNumber}`, PAGE_W - MARGIN, pageH - 10, { align: 'right' });

  doc.setFontSize(6);
  doc.setTextColor(...GREY);
  doc.text(
    'This certificate is issued in accordance with WearCheck ARC quality management procedures aligned with ISO/IEC 17025.',
    MARGIN, pageH - 6.5,
  );
}

function sectionHeader(doc, text, y) {
  doc.setFillColor(...LIGHT_GREY);
  doc.rect(MARGIN, y, CONTENT_W, 6, 'F');
  doc.setDrawColor(220, 220, 220);
  doc.line(MARGIN, y + 6, MARGIN + CONTENT_W, y + 6);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(9);
  doc.setTextColor(...BLACK);
  doc.text(text, MARGIN + 3, y + 4.2);
  return y + 7;
}

function subSectionHeader(doc, text, y) {
  doc.setFillColor(245, 245, 245);
  doc.rect(MARGIN, y, CONTENT_W, 5, 'F');
  doc.setDrawColor(220, 220, 220);
  doc.line(MARGIN, y + 5, MARGIN + CONTENT_W, y + 5);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor(...GREY);
  doc.text(text, MARGIN + 5, y + 3.5);
  return y + 6;
}

function fieldRows(doc, fields, y, stacked) {
  const colW = CONTENT_W / 2;
  const labelW = colW * 0.42;
  const rowH = 5;

  if (stacked) {
    // Full-width rows: one field per row
    const fullLabelW = CONTENT_W * 0.18;
    for (let i = 0; i < fields.length; i++) {
      doc.setDrawColor(230, 230, 230);
      doc.line(MARGIN, y + rowH, MARGIN + CONTENT_W, y + rowH);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);
      doc.setTextColor(...BLACK);
      doc.text(fields[i][0], MARGIN + 3, y + 3.5);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.text(String(fields[i][1] || ''), MARGIN + fullLabelW, y + 3.5);

      y += rowH;
    }
    return y + 1;
  }

  for (let i = 0; i < fields.length; i += 2) {
    // Draw light separator
    doc.setDrawColor(230, 230, 230);
    doc.line(MARGIN, y + rowH, MARGIN + CONTENT_W, y + rowH);

    // Left pair
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setTextColor(...BLACK);
    doc.text(fields[i][0], MARGIN + 3, y + 3.5);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text(String(fields[i][1] || ''), MARGIN + labelW, y + 3.5);

    // Right pair
    if (i + 1 < fields.length) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);
      doc.text(fields[i + 1][0], MARGIN + colW + 3, y + 3.5);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.text(String(fields[i + 1][1] || ''), MARGIN + colW + labelW, y + 3.5);
    }

    y += rowH;
  }
  return y + 1;
}

export async function generateCoverPage(info) {
  // Load logo
  try {
    const resp = await fetch('/WearCheck Logo.png');
    const blob = await resp.blob();
    const reader = new FileReader();
    info._logoBase64 = await new Promise((resolve) => {
      reader.onload = () => resolve(reader.result);
      reader.readAsDataURL(blob);
    });
  } catch { /* no logo */ }

  const doc = new jsPDF({ unit: 'mm', format: 'a4' });

  drawHeader(doc, info);

  let y = 24;

  // Title
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  doc.setTextColor(...BLACK);
  doc.text('CERTIFICATE OF CALIBRATION', PAGE_W / 2, y, { align: 'center' });
  y += 3;
  doc.setDrawColor(200, 200, 200);
  doc.setLineWidth(0.3);
  doc.line(MARGIN, y, PAGE_W - MARGIN, y);
  y += 3;

  // Certificate Information
  const serial = info.serial || '';
  const calDate = info.cal_date || '';
  const now = new Date().toISOString().slice(0, 10);
  const certNum = serial ? `VIB-${serial}-${now}` : '\u2014';

  y = sectionHeader(doc, 'CERTIFICATE INFORMATION', y);
  y = fieldRows(doc, [
    ['Certificate No:', certNum],
    ['Date of Issue:', now],
    ['Calibration Date:', calDate],
    ['Next Due:', info.cal_due || ''],
    ['Customer:', info.customer || ''],
    ['Calibration Technician:', info.cal_tech || ''],
  ], y);

  // Analyzer Information
  y = sectionHeader(doc, 'ANALYZER / METER INFORMATION', y + 1);
  y = fieldRows(doc, [
    ['Manufacturer:', info.manufacturer || ''],
    ['Model:', info.model || ''],
    ['Serial No:', serial],
    ['Description:', 'Vibration Analyzer'],
  ], y);

  // Analyzer Settings
  y = sectionHeader(doc, 'ANALYZER SETTINGS', y + 1);
  y = fieldRows(doc, [
    ['Analyzer Mode:', info.analyzer_mode || ''],
    ['Frequency Max:', `${info.freq_max || ''} ${info.freq_unit || ''}`],
    ['Frequency Min:', `${info.freq_min || ''} ${info.freq_unit || ''}`],
    ['Lines of Resolution:', info.lines_resolution || ''],
    ['Averaging Points:', info.avg_points || ''],
    ['Window Type:', info.window_type || ''],
    ['Sensor Input Sensitivity:', `${info.sensor_input_sens || ''} ${info.sensor_input_unit || ''}`],
    ['', ''],
  ], y);

  // Tests Performed
  y = sectionHeader(doc, 'TESTS PERFORMED', y + 1);
  let testNum = 0;

  if (info.acc_test_type) {
    testNum++;
    y = subSectionHeader(doc, `Test ${testNum}: ${info.acc_test_type}`, y);
    const accFields = Object.entries(info.acc_params || {}).map(([k, v]) => [k, v]);
    if (accFields.length % 2 !== 0) accFields.push(['', '']);
    if (accFields.length) y = fieldRows(doc, accFields, y);
  }

  if (info.vel_test_type) {
    testNum++;
    y = subSectionHeader(doc, `Test ${testNum}: ${info.vel_test_type}`, y);
    const velFields = Object.entries(info.vel_params || {}).map(([k, v]) => [k, v]);
    if (velFields.length % 2 !== 0) velFields.push(['', '']);
    if (velFields.length) y = fieldRows(doc, velFields, y);
  }

  // Test Equipment
  y = sectionHeader(doc, 'TEST EQUIPMENT', y + 1);
  let pvcLine = `PVC M/N: ${info.pvc_model || ''}, S/N: ${info.pvc_serial || ''}`;
  if (info.pvc_cal_date) pvcLine += `, Cal: ${info.pvc_cal_date}`;
  let sensorLine = `Sensor M/N: ${info.sensor_model || ''}, S/N: ${info.sensor_serial || ''}`;
  if (info.sensor_cal_date) sensorLine += `, Cal: ${info.sensor_cal_date}`;

  const pvcSens = `${info.pvc_sensitivity || ''} ${info.pvc_sens_unit || ''}`.trim();
  const sensorSens = `${info.sensor_sensitivity || ''} ${info.sensor_sens_unit || ''}`.trim();

  y = fieldRows(doc, [
    ['PVC:', pvcLine],
    ['Sensitivity:', pvcSens],
    ['Sensor:', sensorLine],
    ['Sensitivity:', sensorSens],
  ], y, true);

  // Procedure
  y = sectionHeader(doc, 'PROCEDURE', y + 1);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.setTextColor(...BLACK);
  doc.text('Back-to-Back Comparison per ISO 16063-21', MARGIN + 3, y + 3.5);
  y += 6;

  doc.setFont('helvetica', 'italic');
  doc.setFontSize(6.5);
  doc.setTextColor(...GREY);
  const traceText = 'Traceability: The measurements reported herein are traceable to NIST (USA) and PTB (Germany) through an unbroken chain of calibrations.';
  const splitTrace = doc.splitTextToSize(traceText, CONTENT_W - 6);
  doc.text(splitTrace, MARGIN + 3, y + 2);
  y += splitTrace.length * 3 + 2;

  const uncertText = 'Uncertainty: The reported expanded uncertainty is based on a standard uncertainty multiplied by a coverage factor k=2, providing a level of confidence of approximately 95%.';
  const splitUncert = doc.splitTextToSize(uncertText, CONTENT_W - 6);
  doc.text(splitUncert, MARGIN + 3, y + 2);
  y += splitUncert.length * 3 + 2;

  if (info.note) {
    const noteText = `Note: ${info.note}`;
    const splitNote = doc.splitTextToSize(noteText, CONTENT_W - 6);
    doc.text(splitNote, MARGIN + 3, y + 2);
    y += splitNote.length * 3 + 2;
  }

  // Authorisation
  y = sectionHeader(doc, 'AUTHORISATION', y + 2);
  y += 2;

  const blockW = (CONTENT_W - 10) / 2;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7.5);
  doc.setTextColor(...BLACK);
  doc.text('Calibrated By:', MARGIN + 3, y + 3);
  doc.text('Approved By:', MARGIN + blockW + 13, y + 3);
  y += 6;

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.text(info.cal_tech || '', MARGIN + 3, y + 3);
  y += 5;

  // Signature lines
  doc.setDrawColor(...BLACK);
  doc.setLineWidth(0.3);
  doc.line(MARGIN + 3, y + 8, MARGIN + 3 + blockW, y + 8);
  doc.line(MARGIN + blockW + 13, y + 8, MARGIN + blockW + 13 + blockW, y + 8);
  y += 10;

  doc.setFont('helvetica', 'italic');
  doc.setFontSize(6.5);
  doc.setTextColor(...GREY);
  doc.text('Name', MARGIN + 3, y + 2);
  doc.text('Signature', MARGIN + 3 + blockW * 0.55, y + 2);
  doc.text('Name', MARGIN + blockW + 13, y + 2);
  doc.text('Signature', MARGIN + blockW + 13 + blockW * 0.55, y + 2);
  y += 6;

  // Date lines
  const today = new Date().toLocaleDateString('en-ZA', { year: 'numeric', month: '2-digit', day: '2-digit' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.setTextColor(...BLACK);
  doc.text(`Date: ${today}`, MARGIN + 3, y + 3);
  doc.text('Date:', MARGIN + blockW + 13, y + 3);
  y += 5;

  doc.setDrawColor(...BLACK);
  doc.line(MARGIN + 3, y, MARGIN + 3 + blockW, y);
  doc.line(MARGIN + blockW + 13, y, MARGIN + blockW + 13 + blockW, y);

  return doc.output('arraybuffer');
}
