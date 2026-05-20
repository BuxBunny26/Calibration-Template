// Signature lookup + base64 loader for jsPDF embedding.
//
// PNGs live in /public/Signatures and are fetched on first use, then cached
// (as data URLs) so subsequent PDFs are instant.

const SIGNATURE_MAP = {
  'andrew robb':  'Andrew_Robb_Signature-removebg-preview.png',
  'andrew':       'Andrew_Robb_Signature-removebg-preview.png',
  'robb':         'Andrew_Robb_Signature-removebg-preview.png',
  'edward jnr':   'Edward_Jnr_Signature-removebg-preview.png',
  'edward':       'Edward_Jnr_Signature-removebg-preview.png',
  'eddie jnr':    'Edward_Jnr_Signature-removebg-preview.png',
  'eddie':        'Edward_Jnr_Signature-removebg-preview.png',
};

const cache = {};

function resolveFile(name) {
  if (!name) return null;
  const key = String(name).trim().toLowerCase().replace(/\s+/g, ' ');
  if (!key) return null;
  if (SIGNATURE_MAP[key]) return SIGNATURE_MAP[key];
  for (const k of Object.keys(SIGNATURE_MAP)) {
    if (k.includes(key) || key.includes(k)) return SIGNATURE_MAP[k];
  }
  return null;
}

async function loadAsDataURL(file) {
  if (cache[file]) return cache[file];
  const resp = await fetch(`/Signatures/${encodeURIComponent(file)}`);
  if (!resp.ok) return null;
  const blob = await resp.blob();
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
  cache[file] = { dataUrl, naturalSize: await getImageSize(dataUrl) };
  return cache[file];
}

function getImageSize(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 0, h: 0 });
    img.src = dataUrl;
  });
}

// Returns { dataUrl, width, height } sized to fit inside maxW x maxH (mm),
// preserving aspect ratio. Returns null if no signature found.
export async function getSignature(name, maxW = 35, maxH = 12) {
  const file = resolveFile(name);
  if (!file) return null;
  try {
    const entry = await loadAsDataURL(file);
    if (!entry || !entry.naturalSize.w) return null;
    const { w, h } = entry.naturalSize;
    // Fit inside maxW x maxH (mm), preserving aspect ratio.
    const scale = Math.min(maxW / w, maxH / h);
    return {
      dataUrl: entry.dataUrl,
      width: w * scale,
      height: h * scale,
    };
  } catch {
    return null;
  }
}
