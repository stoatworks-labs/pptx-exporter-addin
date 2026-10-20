#!/usr/bin/env python3
"""Open the Office and ODF samples the tests write, with other people's readers.

    SAMPLES_DIR=.out/office npx vitest run test/office.test.ts
    .venv/bin/pip install python-pptx python-docx odfpy     (once)
    .venv/bin/python scripts/check-office.py .out/office

python-pptx for the .pptx/.ppsx/.potx copies and the flattened deck,
python-docx for the Word handouts, odfpy for the .odp; every zip's CRCs and
every XML part's well-formedness besides. It prints what each reader saw and
exits non-zero on the first thing it couldn't read. None of this is
PowerPoint or Word: it catches broken packages, not every repair dialog.
"""
import glob
import os
import sys
import zipfile

from lxml import etree

import pptx.api
from pptx import Presentation
from pptx.enum.action import PP_ACTION
from pptx.util import Emu
from docx import Document
from odf import draw
from odf import presentation as odfpres
from odf import teletype
from odf import text as odftext
from odf.namespaces import PRESENTATIONNS, XLINKNS
from odf.opendocument import load as load_odf

PML = 'application/vnd.openxmlformats-officedocument.presentationml.'
SHOW_TYPES = (PML + 'presentation.main+xml', PML + 'slideshow.main+xml', PML + 'template.main+xml')
ODP_MIME = 'application/vnd.oasis.opendocument.presentation'
FIXTURE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'test', 'fixtures', 'sample-16x9.pptx')

# python-pptx only opens a package whose main part is a .pptx or .pptm; a
# show or a template is the same package with another content type.
pptx.api._is_pptx_package = lambda part: part.content_type in SHOW_TYPES

problems = []


def problem(msg):
    problems.append(msg)
    print(f'  !! {msg}')


def check_zip(path):
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        if bad:
            problem(f'{os.path.basename(path)}: CRC mismatch in {bad}')
        for info in z.infolist():
            if info.filename.endswith(('.xml', '.rels')):
                try:
                    etree.fromstring(z.read(info.filename))
                except etree.XMLSyntaxError as e:
                    problem(f'{os.path.basename(path)}: {info.filename}: {e}')
        return z.namelist()


def short(s, n=60):
    s = s.replace('\n', ' / ')
    return s if len(s) <= n else s[: n - 1] + '…'


def describe_click(prs, click):
    if click.action == PP_ACTION.HYPERLINK:
        return click.hyperlink.address
    if click.action == PP_ACTION.NAMED_SLIDE:
        return f'jump to slide {prs.slides.index(click.target_slide) + 1}'
    return click.action.name.lower()


def check_pptx(path):
    names = check_zip(path)
    prs = Presentation(path)
    main = prs.part.content_type.rsplit('.', 2)[-2]
    print(f'{os.path.basename(path)}: {len(prs.slides)} slides, {Emu(prs.slide_width).inches:.2f} × '
          f'{Emu(prs.slide_height).inches:.2f} in, main part "{main}", {len(names)} zip entries')
    # Reading every part through relationships fails on a dangling target.
    parts = list(prs.part.package.iter_parts())
    for i, slide in enumerate(prs.slides, 1):
        title = slide.shapes.title.text if slide.shapes.title is not None else ''
        hidden = slide._element.get('show') in ('0', 'false')
        notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ''
        pics = [sh for sh in slide.shapes if sh.shape_type == 13]
        for p in pics:
            if not p.image.blob:
                problem(f'{os.path.basename(path)} slide {i}: empty picture')
        trans = slide._element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}transition')
        effect = trans[0].tag.split('}')[1] if trans is not None and len(trans) else ''
        links = [describe_click(prs, sh.click_action) for sh in slide.shapes if sh.click_action.action != PP_ACTION.NONE]
        print(f'  {i:2} {"[hidden] " if hidden else ""}{short(title, 40)!r}'
              f'{f"  pictures={len(pics)}" if pics else ""}'
              f'{f"  transition={effect}" if effect else ""}'
              f'{f"  links={links}" if links else ""}'
              f'{f"  notes={short(notes)!r}" if notes else ""}')
    print(f'  {len(parts)} parts reachable, core title {prs.core_properties.title!r}, '
          f'author {prs.core_properties.author!r}, last modified by {prs.core_properties.last_modified_by!r}')
    return prs


def check_docx(path):
    check_zip(path)
    doc = Document(path)
    heads = [p.text for p in doc.paragraphs if p.style.name.startswith('Heading')]
    shapes = doc.inline_shapes
    sizes = sorted({f'{Emu(s.width).cm:.1f}×{Emu(s.height).cm:.1f} cm' for s in shapes})
    for s in shapes:
        rid = s._inline.graphic.graphicData.pic.blipFill.blip.embed
        if not doc.part.related_parts[rid].blob:
            problem(f'{os.path.basename(path)}: empty picture {rid}')
    sec = doc.sections[0]
    print(f'{os.path.basename(path)}: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables, '
          f'{len(shapes)} pictures ({", ".join(sizes)}), page {Emu(sec.page_width).cm:.1f}×{Emu(sec.page_height).cm:.1f} cm, '
          f'title {doc.core_properties.title!r}')
    for h in heads[:4]:
        print(f'  heading {h!r}')
    notes = [p.text for p in doc.paragraphs if p.text and not p.style.name.startswith('Heading')]
    bold = [r.text for p in doc.paragraphs for r in p.runs if r.bold]
    print(f'  first text: {[short(n, 40) for n in notes[:4]]}  bold runs: {bold[:4]}')


def check_odp(path):
    names = check_zip(path)
    with zipfile.ZipFile(path) as z:
        first = z.infolist()[0]
        if first.filename != 'mimetype' or first.compress_type != zipfile.ZIP_STORED or first.extra:
            problem(f'{os.path.basename(path)}: mimetype must be the first entry, stored, with no extra field')
        if z.read('mimetype').decode() != ODP_MIME:
            problem(f'{os.path.basename(path)}: wrong mimetype')
        manifest = etree.fromstring(z.read('META-INF/manifest.xml'))
        mns = '{urn:oasis:names:tc:opendocument:xmlns:manifest:1.0}'
        listed = {e.get(mns + 'full-path') for e in manifest.iter(mns + 'file-entry')} - {'/'}
        files = set(names) - {'mimetype', 'META-INF/manifest.xml'}
        if listed != files:
            problem(f'{os.path.basename(path)}: manifest {sorted(listed ^ files)} out of step with the zip')
    doc = load_odf(path)
    if doc.mimetype != ODP_MIME:
        problem(f'{os.path.basename(path)}: odfpy reads mimetype {doc.mimetype}')
    pages = doc.presentation.getElementsByType(draw.Page)
    print(f'{os.path.basename(path)}: {len(pages)} pages, mimetype ok, manifest lists {len(listed)} files')
    for i, page in enumerate(pages, 1):
        frames = page.getElementsByType(draw.Frame)
        image = page.getElementsByType(draw.Image)[0]
        href = image.getAttrNS(XLINKNS, 'href')
        if href not in names:
            problem(f'{os.path.basename(path)} page {i}: picture {href} missing')
        svg_title = [str(t) for t in frames[0].childNodes if t.qname[1] == 'title']
        style = page.getAttribute('stylename')
        notes = page.getElementsByType(odfpres.Notes)
        note_text = ' / '.join(teletype.extractText(p) for p in notes[0].getElementsByType(odftext.P)) if notes else ''
        print(f'  {i:2} {"[hidden] " if style == "dp2" else ""}{short(svg_title[0] if svg_title else "", 40)!r} '
              f'{href}{f"  notes={short(note_text)!r}" if note_text else ""}')


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else os.environ.get('SAMPLES_DIR', '.out/office')
    files = sorted(glob.glob(os.path.join(folder, '*')))
    office = [f for f in files if f.endswith(('.pptx', '.ppsx', '.potx', '.docx', '.odp'))]
    if not office:
        sys.exit(f'No .pptx/.ppsx/.potx/.docx/.odp files in {folder} — run the tests with SAMPLES_DIR set first.')
    original = Presentation(FIXTURE)
    for path in office:
        try:
            if path.endswith('.docx'):
                check_docx(path)
            elif path.endswith('.odp'):
                check_odp(path)
            else:
                prs = check_pptx(path)
                if 'flattened' in path:
                    if (prs.slide_width, prs.slide_height) != (original.slide_width, original.slide_height):
                        problem('flattened deck: slide size differs from the original')
                    if len(prs.slides) != len(original.slides):
                        problem('flattened deck: slide count differs from the original')
        except Exception as e:  # noqa: BLE001 — any reader failure is the finding
            problem(f'{os.path.basename(path)}: {type(e).__name__}: {e}')
        print()
    if problems:
        print(f'{len(problems)} problem(s).')
        sys.exit(1)
    print(f'All {len(office)} files read cleanly.')


if __name__ == '__main__':
    main()
