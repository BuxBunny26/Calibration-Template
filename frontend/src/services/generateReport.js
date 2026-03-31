import { PDFDocument } from 'pdf-lib';
import { readVibAnalyzer, mergeInfo } from './excelReader';
import { generateCoverPage } from './coverPage';
import { supabase, supabaseConfigured } from './supabase';

function generateId() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = (Math.random() * 16) | 0;
    return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
  });
}

function classifyFile(filename) {
  const lower = filename.toLowerCase();
  const ext = lower.split('.').pop();
  if (!['xlsx', 'pdf'].includes(ext)) return { kind: null, ext };
  const kind = lower.includes('acc') ? 'acc' : lower.includes('vel') ? 'vel' : null;
  return { kind, ext };
}

async function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function mergePdfs(coverBytes, dataPdfBuffers) {
  const merged = await PDFDocument.create();

  // Add cover page
  const coverDoc = await PDFDocument.load(coverBytes);
  const coverPages = await merged.copyPages(coverDoc, coverDoc.getPageIndices());
  coverPages.forEach(p => merged.addPage(p));

  // Add data PDFs
  for (const buf of dataPdfBuffers) {
    const src = await PDFDocument.load(buf);
    const pages = await merged.copyPages(src, src.getPageIndices());
    pages.forEach(p => merged.addPage(p));
  }

  return merged.save();
}

function triggerDownload(bytes, filename) {
  const blob = new Blob([bytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function processGroup(groupFiles) {
  const jobId = generateId();

  // Classify files
  const slots = {};
  for (const file of groupFiles) {
    const { kind, ext } = classifyFile(file.name);
    if (kind && ext) slots[`${kind}_${ext}`] = file;
  }

  const hasAcc = slots.acc_xlsx || slots.acc_pdf;
  const hasVel = slots.vel_xlsx || slots.vel_pdf;

  if (!hasAcc && !hasVel) throw new Error('No ACC or VEL files found.');
  if (hasAcc && !slots.acc_xlsx) throw new Error('ACC PDF provided but ACC Excel is missing.');
  if (hasAcc && !slots.acc_pdf) throw new Error('ACC Excel provided but ACC PDF is missing.');
  if (hasVel && !slots.vel_xlsx) throw new Error('VEL PDF provided but VEL Excel is missing.');
  if (hasVel && !slots.vel_pdf) throw new Error('VEL Excel provided but VEL PDF is missing.');

  // Read Excel data
  let accInfo = null, velInfo = null;
  if (slots.acc_xlsx) {
    const buf = await readFileAsArrayBuffer(slots.acc_xlsx);
    accInfo = readVibAnalyzer(buf);
  }
  if (slots.vel_xlsx) {
    const buf = await readFileAsArrayBuffer(slots.vel_xlsx);
    velInfo = readVibAnalyzer(buf);
  }

  const merged = mergeInfo(accInfo, velInfo);

  // Generate cover page
  const coverBytes = await generateCoverPage(merged);

  // Read data PDFs
  const dataPdfBuffers = [];
  if (slots.acc_pdf) dataPdfBuffers.push(await readFileAsArrayBuffer(slots.acc_pdf));
  if (slots.vel_pdf) dataPdfBuffers.push(await readFileAsArrayBuffer(slots.vel_pdf));

  // Merge all into final PDF
  const finalBytes = await mergePdfs(coverBytes, dataPdfBuffers);

  const model = (merged.model || 'Unknown').replace(/\s/g, '_').replace(/\//g, '-');
  const serial = (merged.serial || 'Unknown').replace(/\s/g, '_');
  const calDate = (merged.cal_date || '').replace(/\//g, '-') || 'undated';
  const outputName = `FullReport_${model}_${serial}_${calDate}.pdf`;

  // Upload to Supabase if configured
  let downloadUrl = null;
  if (supabaseConfigured) {
    try {
      const remotePath = `certificates/${jobId}/${outputName}`;
      const { error: uploadErr } = await supabase.storage
        .from('calibration-files')
        .upload(remotePath, new Blob([finalBytes], { type: 'application/pdf' }), {
          contentType: 'application/pdf',
          upsert: true,
        });

      if (!uploadErr) {
        const { data } = supabase.storage
          .from('calibration-files')
          .getPublicUrl(remotePath);
        downloadUrl = data?.publicUrl || null;
      }

      // Save job record
      await supabase.from('certificate_jobs').upsert({
        id: jobId,
        status: 'completed',
        equipment_model: merged.model || '',
        serial_number: merged.serial || '',
        manufacturer: merged.manufacturer || '',
        calibration_date: merged.cal_date || '',
        calibration_due: merged.cal_due || '',
        calibration_tech: merged.cal_tech || '',
        customer: merged.customer || '',
        input_files: JSON.stringify(groupFiles.map(f => ({ name: f.name, size: f.size }))),
        output_file_path: remotePath,
        output_file_name: outputName,
        completed_at: new Date().toISOString(),
      });
    } catch (e) {
      console.warn('Supabase upload/save failed:', e);
    }
  }

  // Always trigger direct download (works without backend)
  triggerDownload(finalBytes, outputName);

  // Create a blob URL as fallback when Supabase is not configured
  if (!downloadUrl) {
    const blob = new Blob([finalBytes], { type: 'application/pdf' });
    downloadUrl = URL.createObjectURL(blob);
  }

  return {
    job_id: jobId,
    status: 'completed',
    output_file_name: outputName,
    download_url: downloadUrl,
    model: merged.model || '',
    serial: merged.serial || '',
    _pdfBytes: finalBytes,
  };
}
