'''Read a PDF into text plus chart blocks an agent cannot skip.

An image is a chart when it covers at least 8 percent of the page. Logos and
icons stay in the page text only. Each chart is rendered to a PNG and, when
Tesseract is available, OCR'd into the same extract as the prose. Otherwise
the extract names the PNG the agent still has to open.
'''

import re
from pathlib import Path

CHART_AREA_FRACTION = 0.08
_PAGE_REF = re.compile(r'(?:p\.|page)\s*(\d+)', re.IGNORECASE)
_MARKERS = ('chart', 'ocr', 'read from axis', 'no numeric labels')


def is_chart(area_fraction):
    '''True when an embedded image is large enough to be a chart.'''
    try:
        fraction = float(area_fraction)
    except (TypeError, ValueError):
        return False
    return fraction >= CHART_AREA_FRACTION


def read_pdf(path, out_dir, name=None, reader=None, ocr=None):
    '''Extract page text and write a chart block for every large image.

    ``reader`` and ``ocr`` are test hooks. The default reader uses PyMuPDF.
    ``ocr`` receives the PNG path and returns text, or None when OCR did not
    run. The return value is one document for ``build_baseline``.
    '''
    path = Path(path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    raw_pages = (reader or _pymupdf_pages)(path)
    extract_lines = []
    pages = []
    for raw in raw_pages:
        page_no = int(raw['page'])
        text = raw.get('text') or ''
        extract_lines.append('=== page %d ===' % page_no)
        if text:
            extract_lines.append(text.rstrip())
        charts = []
        chart_images = [img for img in (raw.get('images') or []) if is_chart(img.get('area_fraction'))]
        page_png = out_dir / ('page_%d.png' % page_no)
        png_bytes = raw.get('png')
        if chart_images and png_bytes:
            page_png.write_bytes(png_bytes)
        page_ocr = _clean_ocr(raw.get('ocr'))
        if chart_images and page_ocr is None and png_bytes and page_png.is_file():
            page_ocr = _clean_ocr((ocr or _ocr_png)(page_png))
        for img in chart_images:
            img_ocr = _clean_ocr(img['ocr']) if 'ocr' in img else page_ocr
            charts.append({
                'png': str(page_png),
                'ocr': img_ocr,
                'needs_visual_read': img_ocr is None,
            })
            extract_lines.append('--- chart p. %d ---' % page_no)
            if img_ocr:
                extract_lines.append('OCR: %s' % img_ocr)
            else:
                extract_lines.append('UNREAD CHART: open %s' % page_png.name)
        pages.append({'page': page_no, 'text': text, 'charts': charts})
    extract_path = out_dir / ('%s_extract.txt' % path.stem)
    extract_path.write_text('\n'.join(extract_lines) + ('\n' if extract_lines else ''), encoding='utf-8')
    return {
        'name': name if name else path.stem,
        'path': str(path),
        'extract': str(extract_path),
        'pages': pages,
    }


def unread_charts(documents, article=None, attachment=None, attachments=None):
    '''Chart pages whose matching source does not cite the page.

    A citation is one leaf whose evidence contains ``p. N`` or ``page N`` and
    one of ``chart``, ``OCR``, ``read from axis``, or ``no numeric labels``.
    The article document must be named ``article``. An attachment document
    must use the same name as that attachment.
    '''
    unread = []
    for doc in _as_list(documents):
        if not isinstance(doc, dict):
            continue
        payload = _payload_for(doc.get('name'), article, attachment, attachments)
        evidences = _evidences(payload)
        for page in doc.get('pages') or []:
            page_no = page.get('page')
            for chart in page.get('charts') or []:
                if _cited(evidences, page_no):
                    continue
                unread.append({
                    'name': doc.get('name'),
                    'page': page_no,
                    'png': chart.get('png'),
                    'ocr': chart.get('ocr'),
                })
    return unread


def _cited(evidences, page_no):
    try:
        page_no = int(page_no)
    except (TypeError, ValueError):
        return False
    for text in evidences:
        low = text.lower()
        if not any(marker in low for marker in _MARKERS):
            continue
        if any(int(match) == page_no for match in _PAGE_REF.findall(text)):
            return True
    return False


def _evidences(node):
    found = []
    _walk_evidence(node, found)
    return found


def _walk_evidence(node, found):
    if isinstance(node, dict):
        if 'value' in node and 'evidence' in node:
            evidence = node.get('evidence')
            if isinstance(evidence, str) and evidence.strip():
                found.append(evidence)
        for value in node.values():
            _walk_evidence(value, found)
    elif isinstance(node, list):
        for value in node:
            _walk_evidence(value, found)


def _payload_for(doc_name, article, attachment, attachments):
    name = str(doc_name or '').strip()
    if name.lower() == 'article':
        return article or {}
    for item in _attachment_items(attachment, attachments):
        if str(item.get('name') or '') == name:
            return item.get('leaves') or {}
    return {}


def _attachment_items(attachment, attachments):
    items = []
    if attachment:
        if isinstance(attachment, dict) and 'leaves' in attachment and 'name' in attachment:
            items.append(attachment)
        else:
            items.append({'name': None, 'leaves': attachment})
    for item in attachments or []:
        if isinstance(item, dict) and 'leaves' in item:
            items.append({'name': item.get('name'), 'leaves': item.get('leaves')})
        elif item:
            items.append({'name': None, 'leaves': item})
    return items


def _as_list(documents):
    if not documents:
        return []
    if isinstance(documents, dict):
        return [documents]
    return list(documents)


def _clean_ocr(value):
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _pymupdf_pages(path):
    try:
        import fitz
    except ImportError as exc:
        raise ImportError(
            'PyMuPDF is required to read a PDF. Install it with pip install -e ".[pdf]".'
        ) from exc
    doc = fitz.open(str(path))
    try:
        pages = []
        for index, page in enumerate(doc):
            rect = page.rect
            page_area = float(rect.width * rect.height) or 1.0
            images = []
            try:
                infos = page.get_image_info() or []
            except Exception:
                infos = []
            for info in infos:
                bbox = info.get('bbox') if isinstance(info, dict) else None
                if not bbox or len(bbox) < 4:
                    continue
                width = abs(float(bbox[2]) - float(bbox[0]))
                height = abs(float(bbox[3]) - float(bbox[1]))
                images.append({'area_fraction': (width * height) / page_area})
            png = None
            if any(is_chart(img['area_fraction']) for img in images):
                png = page.get_pixmap(matrix=fitz.Matrix(2, 2)).tobytes('png')
            pages.append({
                'page': index + 1,
                'text': page.get_text() or '',
                'images': images,
                'png': png,
            })
        return pages
    finally:
        doc.close()


def _ocr_png(png_path):
    try:
        import pytesseract
        text = pytesseract.image_to_string(str(png_path))
    except Exception:
        return None
    return _clean_ocr(text)
