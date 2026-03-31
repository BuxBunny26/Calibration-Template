const API_URL = process.env.REACT_APP_API_URL || 'http://localhost:5000';

export async function generateCertificate(files) {
  const formData = new FormData();
  files.forEach(f => formData.append('files', f));

  const res = await fetch(`${API_URL}/api/generate`, {
    method: 'POST',
    body: formData,
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: 'Generation failed' }));
    throw new Error(err.error || 'Generation failed');
  }

  return res.json();
}
