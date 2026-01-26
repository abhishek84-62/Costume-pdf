#!/usr/bin/env python3
"""
Flask backend for Costume-pdf
- Convert PDF <-> PPTX using LibreOffice (soffice)
- Detect blank pages/slides by rasterizing pages (pdf2image / poppler) and checking whiteness
- If no blank page found, generate a blank page matching the source page size
- Export blank page as PDF or PNG
"""
import os
import tempfile
import subprocess
from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
from PIL import Image
from pypdf import PdfReader, PdfWriter

ALLOWED_EXT = {'.pdf', '.ppt', '.pptx'}
DEFAULT_DPI = 300
BLANK_THRESHOLD = 0.02  # fraction of non-white pixels below which page is considered blank

app = Flask(__name__)

def run_cmd(cmd):
    proc = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"Command failed: {cmd}\nstdout: {proc.stdout}\nstderr: {proc.stderr}")
    return proc.stdout

def libreoffice_convert(input_path, out_dir, target_format):
    """Convert input file to target_format ('pdf' or 'pptx') using soffice headless."""
    if target_format not in ('pdf', 'pptx'):
        raise ValueError('target_format must be pdf or pptx')
    # Ensure output dir exists
    os.makedirs(out_dir, exist_ok=True)
    cmd = f'soffice --headless --invisible --convert-to {target_format} "{input_path}" --outdir "{out_dir}"'
    run_cmd(cmd)
    base = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(out_dir, f'{base}.{target_format}')
    if os.path.exists(out_path):
        return out_path
    # try to find anything produced in out_dir with base prefix
    for f in os.listdir(out_dir):
        if f.startswith(base + '.') and f.lower().endswith('.' + target_format):
            return os.path.join(out_dir, f)
    raise RuntimeError('LibreOffice conversion failed or output not found')

def render_pdf_pages(pdf_path, dpi=DEFAULT_DPI):
    """Render PDF pages to PIL Images using pdf2image (requires poppler)."""
    # convert_from_path returns list of PIL.Image
    images = convert_from_path(pdf_path, dpi=dpi, fmt='png')
    return images

def fraction_nonwhite_pixels(pil_img):
    """Return fraction of pixels that are NOT near-white in the given PIL.Image."""
    # Downscale for performance
    base_width = 800
    img = pil_img.convert('RGB')
    if img.width > base_width:
        ratio = base_width / img.width
        img = img.resize((base_width, int(img.height * ratio)), resample=Image.BILINEAR)
    pixels = img.getdata()
    total = 0
    nonwhite = 0
    for r, g, b in pixels:
        total += 1
        if not (r > 245 and g > 245 and b > 245):
            nonwhite += 1
    return nonwhite / total if total else 0.0

def extract_pdf_page(pdf_path, page_index, dest_path):
    """Extract a single page (0-based) from PDF and save to dest_path."""
    reader = PdfReader(pdf_path)
    if page_index < 0 or page_index >= len(reader.pages):
        raise IndexError('page_index out of range')
    writer = PdfWriter()
    writer.add_page(reader.pages[page_index])
    with open(dest_path, 'wb') as f:
        writer.write(f)
    return dest_path

def create_blank_pdf_like(pdf_path, dest_path, page_index=0):
    """Create a one-page blank PDF whose page size matches page_index of pdf_path."""
    reader = PdfReader(pdf_path)
    if len(reader.pages) == 0:
        # default A4 if no page
        width = 595.276  # points
        height = 841.89
    else:
        idx = max(0, min(page_index, len(reader.pages)-1))
        p = reader.pages[idx]
        # Get mediabox width/height in points
        mediabox = p.mediabox
        width = float(mediabox.right) - float(mediabox.left)
        height = float(mediabox.top) - float(mediabox.bottom)
    writer = PdfWriter()
    writer.add_blank_page(width=width, height=height)
    with open(dest_path, 'wb') as f:
        writer.write(f)
    return dest_path

@app.route('/api/convert', methods=['POST'])
def api_convert():
    """Convert entire file between pdf and pptx. Accepts file and targetFormat('pdf'|'pptx')."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    target = (request.form.get('targetFormat') or '').lower()
    if target not in ('pdf', 'pptx'):
        return jsonify({'error': 'targetFormat must be pdf or pptx'}), 400
    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify({'error': 'Unsupported file type'}), 400

    with tempfile.TemporaryDirectory() as td:
        in_path = os.path.join(td, filename)
        file.save(in_path)
        try:
            out_path = libreoffice_convert(in_path, td, target)
        except Exception as e:
            return jsonify({'error': 'Conversion failed', 'details': str(e)}), 500
        return send_file(out_path, as_attachment=True)

@app.route('/api/extract-blank', methods=['POST'])
def api_extract_blank():
    """Find the first blank page/slide in uploaded PDF/PPTX and return it as PDF or PNG.
    If none found, create a blank page matching the source size.
    Form-data:
      - file: upload
      - outputType: 'pdf'|'image'  (defaults to 'pdf')
      - dpi: (optional) rasterization DPI
    """
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    output_type = (request.form.get('outputType') or 'pdf').lower()
    if output_type not in ('pdf', 'image'):
        return jsonify({'error': 'outputType must be pdf or image'}), 400
    try:
        dpi = int(request.form.get('dpi') or DEFAULT_DPI)
    except ValueError:
        dpi = DEFAULT_DPI

    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify({'error': 'Unsupported file type'}), 400

    with tempfile.TemporaryDirectory() as td:
        in_path = os.path.join(td, filename)
        file.save(in_path)

        # If PPTX -> convert to PDF for detection
        if ext in ('.ppt', '.pptx'):
            try:
                pdf_path = libreoffice_convert(in_path, td, 'pdf')
            except Exception as e:
                return jsonify({'error': 'PPTX->PDF conversion failed', 'details': str(e)}), 500
        else:
            pdf_path = in_path

        # Render pages
        try:
            pages = render_pdf_pages(pdf_path, dpi=dpi)
        except Exception as e:
            return jsonify({'error': 'Rendering PDF failed', 'details': str(e)}), 500

        # Detect first blank page
        found_index = -1
        for i, img in enumerate(pages):
            frac = fraction_nonwhite_pixels(img)
            if frac < BLANK_THRESHOLD:
                found_index = i
                break

        extracted_pdf = os.path.join(td, 'extracted.pdf')
        if found_index == -1:
            # Create blank PDF like first page
            try:
                create_blank_pdf_like(pdf_path, extracted_pdf, page_index=0)
                generated = True
            except Exception as e:
                return jsonify({'error': 'Failed to create blank page', 'details': str(e)}), 500
        else:
            try:
                extract_pdf_page(pdf_path, found_index, extracted_pdf)
                generated = False
            except Exception as e:
                return jsonify({'error': 'Failed to extract page', 'details': str(e)}), 500

        if output_type == 'pdf':
            out_name = ('blank-generated.pdf' if found_index == -1 else f'blank-page-{found_index+1}.pdf')
            return send_file(extracted_pdf, as_attachment=True, download_name=out_name)
        else:
            # Render the single-page PDF to PNG
            try:
                imgs = convert_from_path(extracted_pdf, dpi=dpi, fmt='png')
            except Exception as e:
                return jsonify({'error': 'Failed to render extracted page to image', 'details': str(e)}), 500
            if not imgs:
                return jsonify({'error': 'No image produced'}), 500
            img_path = os.path.join(td, 'out.png')
            imgs[0].save(img_path, format='PNG')
            out_name = ('blank-generated.png' if found_index == -1 else f'blank-page-{found_index+1}.png')
            return send_file(img_path, as_attachment=True, download_name=out_name)

if __name__ == '__main__':
    # For development only. In production use gunicorn/uvicorn + proper config.
    app.run(host='0.0.0.0', port=5000, debug=True)