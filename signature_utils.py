"""
Signature utilities — look up a signatory's signature image by name and
return a sized reportlab Image flowable for placement in PDF tables.

Signature PNGs live in the ``Signatures/`` folder next to this module.
Add new names by extending ``SIGNATURE_MAP`` below.
"""

import os

from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image


SIGNATURES_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "Signatures"
)

# Map lowercase name fragments -> signature file (relative to SIGNATURES_DIR).
# Lookup tries the full lowercase name first, then any fragment that appears
# in (or contains) the supplied name.
SIGNATURE_MAP = {
    "andrew robb":  "Andrew_Robb_Signature-removebg-preview.png",
    "andrew":       "Andrew_Robb_Signature-removebg-preview.png",
    "robb":         "Andrew_Robb_Signature-removebg-preview.png",
    "edward jnr":   "Edward_Jnr_Signature-removebg-preview.png",
    "edward":       "Edward_Jnr_Signature-removebg-preview.png",
    "eddie jnr":    "Edward_Jnr_Signature-removebg-preview.png",
    "eddie":        "Edward_Jnr_Signature-removebg-preview.png",
    "louis":        "Louis_s_Signature-removebg-preview.png",
    "philip":       "Philip's Signature.png",
}


def _resolve_file(name):
    if not name:
        return None
    key = " ".join(str(name).strip().lower().split())
    if not key:
        return None
    # Exact match
    if key in SIGNATURE_MAP:
        return SIGNATURE_MAP[key]
    # Token / substring match
    for k, v in SIGNATURE_MAP.items():
        if k in key or key in k:
            return v
    return None


def get_signature_image(name, max_width=35 * mm, max_height=12 * mm):
    """Return a reportlab ``Image`` flowable for ``name`` sized to fit the
    given bounding box (preserving aspect ratio). Returns ``""`` if no
    matching signature file is found, so the result can be dropped straight
    into a Table cell.
    """
    fname = _resolve_file(name)
    if not fname:
        return ""
    path = os.path.join(SIGNATURES_DIR, fname)
    if not os.path.exists(path):
        return ""
    try:
        iw, ih = ImageReader(path).getSize()
        if iw <= 0 or ih <= 0:
            return ""
        scale = min(max_width / iw, max_height / ih)
        return Image(path, width=iw * scale, height=ih * scale)
    except Exception:
        return ""
