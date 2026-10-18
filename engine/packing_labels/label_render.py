"""Turning courier label pages into the pictures Stitch Labels puts on the packing labels, and doing that for many
labels at once in worker processes.

Rendering a 4x6 label at 600 dpi takes PyMuPDF about a fifth of a second and holds Python while it works, so a job of
200+ labels spent most of its time here, one label after another. render_many() splits the labels over a few worker
processes (python -m engine.packing_labels.label_render JOBFILE) and puts their results back in order. The pictures
are exactly what the app itself would draw: the same code, the same pixels (lossless grayscale, Flate-compressed as
the PDF stores them).

This module loads only PyMuPDF and Pillow, so a worker doesn't load the app, its log or its data. Where workers
can't run (the .exe, a handful of labels, or a worker failing), render_many() returns None and the stitcher renders
the labels itself."""
import json
import math
import os
import struct
import subprocess
import sys
import tempfile
import time
import zlib
from collections import Counter

import pymupdf as fitz
from PIL import Image, ImageOps

DPI = 600
PAD_MM = 1.0  # white kept around the cropped label
WORKER_MIN_LABELS = 24  # fewer labels than this: starting workers would cost more than it saves
WORKER_TIMEOUT = 900    # seconds a worker may take before its labels are rendered here instead


def text_angle(page):
    """Which way the label's text runs as the page is shown: 0 (upright), 90, 180 or 270 degrees clockwise.
    Decided by the text itself (most characters win), plus the page's own rotation flag."""
    votes = Counter()
    for block in page.get_text('dict')['blocks']:
        for line in block.get('lines', []):
            dx, dy = line['dir']
            angle = round(math.degrees(math.atan2(dy, dx)) / 90) % 4 * 90  # y points down, so this is clockwise
            votes[angle] += sum(len(s['text'].strip()) for s in line['spans'])
    return ((votes.most_common(1)[0][0] if votes else 0) + page.rotation) % 360


def render_label(page, dpi=DPI):
    """The courier label as a cropped, lossless grayscale PIL image, turned so its text reads upright
    (a 4x6 label printed sideways on a landscape page comes back the right way up)."""
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False)
    img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
    angle = text_angle(page)
    if angle:
        img = img.rotate(angle, expand=True, fillcolor=255)  # PIL turns anti-clockwise: undoes the clockwise angle
    box = ImageOps.invert(img).point(lambda v: 255 if v > 24 else 0).getbbox()  # ignore near-white noise
    if box:
        pad = round(PAD_MM / 25.4 * dpi)
        box = (max(0, box[0] - pad), max(0, box[1] - pad), min(img.width, box[2] + pad), min(img.height, box[3] + pad))
        img = img.crop(box)
    return img


def turn_for(region_w, region_h, img):
    """img as it goes into a region of this shape: a landscape label turned a quarter (reading bottom to top) for a
    portrait region, so it stays near full size instead of shrinking to the region's width; grayscale."""
    if (img.width > img.height) != (region_w > region_h):
        img = img.rotate(-90, expand=True, fillcolor=255)
    return img if img.mode == "L" else img.convert("L")


def compress_label(img):
    """The label's grayscale pixels, Flate-compressed (lossless) as the PDF stores them."""
    return zlib.compress(img.tobytes(), 6)


def _render_pages(pdf_paths, jobs, region_wh, progress_path=None):
    """[(width, height, flate)] for jobs [(file name, 1-based page)], in order. progress_path: a worker writes how
    many it has done there after each label, so the app can show the count on the loading overlay."""
    docs, out = {}, []
    for name, page in jobs:
        doc = docs.get(name) or docs.setdefault(name, fitz.open(pdf_paths[name]))
        img = turn_for(*region_wh, render_label(doc[page - 1]))
        out.append((img.width, img.height, compress_label(img)))
        if progress_path:
            try:
                with open(progress_path, 'w') as f:
                    f.write(str(len(out)))
            except OSError:
                pass
    for doc in docs.values():
        doc.close()
    return out


def worker_count(labels):
    """How many worker processes for this many labels: none from the .exe (its executable is the app itself) or for a
    handful of labels; else one per 16 labels, up to the cores less one, at most 8."""
    if getattr(sys, 'frozen', False) or labels < WORKER_MIN_LABELS:
        return 0
    return max(2, min(8, (os.cpu_count() or 2) - 1, math.ceil(labels / 16)))


def _count(path):
    try:
        with open(path) as f:
            return int(f.read() or 0)
    except (OSError, ValueError):
        return 0


def render_many(courier_pdfs, jobs, region_wh, app_dir, on_progress=None):
    """Renders jobs [(file name, 1-based page)] of courier_pdfs [(name, bytes)] in worker processes: [(width, height,
    flate)] in the jobs' order, or None when that isn't possible (then render here). app_dir: the folder holding the
    engine package, where the workers start. on_progress(done, total): the labels the workers have done so far, a few
    times a second, for the loading overlay."""
    n = worker_count(len(jobs))
    if not n:
        return None
    with tempfile.TemporaryDirectory(prefix="wh_render_") as tmp:
        pdf_paths = {}
        for i, (name, data) in enumerate(courier_pdfs):
            pdf_paths[name] = os.path.join(tmp, f"courier_{i}.pdf")
            with open(pdf_paths[name], 'wb') as f:
                f.write(data)
        size = math.ceil(len(jobs) / n)
        chunks = [jobs[i:i + size] for i in range(0, len(jobs), size)]
        procs = []
        for i, chunk in enumerate(chunks):
            spec = os.path.join(tmp, f"job_{i}.json")
            with open(spec, 'w', encoding='utf-8') as f:
                json.dump({'pdfs': pdf_paths, 'jobs': chunk, 'region': list(region_wh),
                           'out': os.path.join(tmp, f"out_{i}.bin"), 'progress': os.path.join(tmp, f"done_{i}.txt")}, f)
            procs.append(subprocess.Popen(
                [sys.executable, "-m", "engine.packing_labels.label_render", spec], cwd=app_dir,
                stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)))
        # While they work: their counts added up for the overlay, until all are done or the time is up
        deadline = time.time() + WORKER_TIMEOUT
        while any(p.poll() is None for p in procs) and time.time() < deadline:
            if on_progress:
                on_progress(sum(_count(os.path.join(tmp, f"done_{i}.txt")) for i in range(len(procs))), len(jobs))
            time.sleep(0.3)
        results, failed = [], []
        for i, (proc, chunk) in enumerate(zip(procs, chunks)):
            if proc.poll() is None:
                proc.kill()
                proc.communicate()
                failed.append(f"worker {i + 1} took too long")
                continue
            _, err = proc.communicate()
            if proc.returncode != 0:
                failed.append(f"worker {i + 1}: {err.decode(errors='replace').strip().splitlines()[-1:]}")
                continue
            got = _read_results(os.path.join(tmp, f"out_{i}.bin"))
            if len(got) != len(chunk):
                failed.append(f"worker {i + 1} returned {len(got)} of {len(chunk)} labels")
                continue
            results += got
        if failed:
            raise RuntimeError("; ".join(failed))
        return results


def _write_results(path, results):
    with open(path, 'wb') as f:
        for w, h, data in results:
            f.write(struct.pack('<III', w, h, len(data)))
            f.write(data)


def _read_results(path):
    out = []
    with open(path, 'rb') as f:
        while True:
            head = f.read(12)
            if len(head) < 12:
                return out
            w, h, n = struct.unpack('<III', head)
            out.append((w, h, f.read(n)))


if __name__ == '__main__':  # a worker: python -m engine.packing_labels.label_render JOBFILE
    with open(sys.argv[1], encoding='utf-8') as f:
        spec = json.load(f)
    _write_results(spec['out'], _render_pages(spec['pdfs'], [tuple(j) for j in spec['jobs']], spec['region'],
                                              spec.get('progress')))
