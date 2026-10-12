"""Restore original raster pixels after Office's lossy PDF conversion.

Only confidently matched image resources are replaced. The PDF content streams,
font resources, placement matrices, clipping and vector graphics stay intact.
"""
from io import BytesIO
import math
import posixpath
from pathlib import Path
import tempfile
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from PIL import Image, ImageChops, ImageStat
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject, NumberObject


def _thumbnail(image):
    rgba = image.convert("RGBA")
    canvas = Image.new("RGBA", rgba.size, "white")
    canvas.alpha_composite(rgba)
    return canvas.convert("RGB").resize((160, 160), Image.Resampling.LANCZOS)


def _sources(pptx_path):
    sources = []
    seen = set()
    with ZipFile(pptx_path) as archive:
        # Brightness, recoloring, blur, etc. can legitimately change exported
        # pixels. Do not undo them just because thumbnails look similar.
        blocked = set()
        drawing = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        relationship = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        for owner in archive.namelist():
            if not owner.startswith(("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/")) or not owner.endswith(".xml"):
                continue
            root = ET.fromstring(archive.read(owner))
            ids = set()
            for blip in root.iter(drawing + "blip"):
                effects = [child for child in blip if child.tag != drawing + "extLst"
                           and not (child.tag == drawing + "lum" and all(v == "0" for v in child.attrib.values()))
                           and not (child.tag == drawing + "alphaModFix" and child.get("amt", "100000") == "100000")]
                if effects and blip.get(relationship + "embed"):
                    ids.add(blip.get(relationship + "embed"))
            if ids:
                folder, basename = posixpath.split(owner)
                rel_file = folder + "/_rels/" + basename + ".rels"
                if rel_file in archive.namelist():
                    for rel in ET.fromstring(archive.read(rel_file)):
                        if rel.get("Id") in ids and rel.get("TargetMode") != "External":
                            blocked.add(posixpath.normpath(posixpath.join(folder, rel.get("Target", ""))))
        for name in archive.namelist():
            if not name.startswith("ppt/media/") or name in blocked:
                continue
            data = archive.read(name)
            # Ignore vectors/audio/video: keep their native Office rendering.
            try:
                with Image.open(BytesIO(data)) as image:
                    image.load()
                    if image.format not in {"PNG", "JPEG", "TIFF", "BMP", "GIF"}:
                        continue
                    rgba = image.convert("RGBA")
            except (OSError, ValueError):
                continue
            import hashlib
            digest = hashlib.sha256(rgba.tobytes()).digest()
            if digest in seen:
                continue
            seen.add(digest)
            sources.append((name, rgba, _thumbnail(rgba)))
    return sources


def _lossless_stream(writer, image):
    rgba = image.convert("RGBA")
    stream = DecodedStreamObject()
    stream.set_data(rgba.convert("RGB").tobytes())
    stream.update({NameObject("/Type"): NameObject("/XObject"),
                   NameObject("/Subtype"): NameObject("/Image"),
                   NameObject("/Width"): NumberObject(rgba.width),
                   NameObject("/Height"): NumberObject(rgba.height),
                   NameObject("/ColorSpace"): NameObject("/DeviceRGB"),
                   NameObject("/BitsPerComponent"): NumberObject(8)})
    alpha = rgba.getchannel("A")
    if alpha.getextrema() != (255, 255):
        mask = DecodedStreamObject()
        mask.set_data(alpha.tobytes())
        mask.update({NameObject("/Type"): NameObject("/XObject"),
                     NameObject("/Subtype"): NameObject("/Image"),
                     NameObject("/Width"): NumberObject(rgba.width),
                     NameObject("/Height"): NumberObject(rgba.height),
                     NameObject("/ColorSpace"): NameObject("/DeviceGray"),
                     NameObject("/BitsPerComponent"): NumberObject(8)})
        stream[NameObject("/SMask")] = writer._add_object(mask.flate_encode())
    return stream.flate_encode()


def restore_original_images(pptx_path, pdf_path):
    sources = _sources(pptx_path)
    result = {"restored": [], "unmatched": [], "unchanged": 0}
    if not sources:
        return result
    path = Path(pdf_path)
    reader = PdfReader(path)
    if reader.is_encrypted:
        raise ValueError("PDF зашифрован: восстановление изображений без изменения защиты невозможно.")
    writer = PdfWriter(clone_from=reader)
    processed = set()
    for page in writer.pages:
        # Includes image resources inside Form XObjects, not just top-level pictures.
        for exported in list(page.images):
            ref = exported.indirect_reference
            if ref is None or ref.idnum in processed or exported.image is None:
                continue
            processed.add(ref.idnum)
            native = exported.image.convert("RGBA")
            thumbnail = _thumbnail(native)
            candidates = []
            for name, original, sample in sources:
                ratio_difference = abs(original.width / original.height - native.width / native.height)
                if ratio_difference > .003 or original.width < native.width or original.height < native.height:
                    continue
                error = math.sqrt(sum(x * x for x in ImageStat.Stat(
                    ImageChops.difference(sample, thumbnail)).rms) / 3)
                candidates.append((error, name, original))
            candidates.sort(key=lambda candidate: candidate[0])
            # Don't remove Office effects or choose an ambiguous similar image.
            if (not candidates or candidates[0][0] > 8
                    or (len(candidates) > 1 and candidates[1][0] - candidates[0][0] < 1)):
                result["unmatched"].append(exported.name)
                continue
            error, source_name, original = candidates[0]
            if original.size == native.size and ImageChops.difference(original, native).getbbox(alpha_only=False) is None:
                result["unchanged"] += 1
                continue
            replacement = _lossless_stream(writer, original)
            existing = ref.get_object()
            for key in ("/Interpolate", "/Intent", "/OC"):
                if key in existing:
                    replacement[NameObject(key)] = existing[key]
            # Preserve the resource object number so every existing use, including
            # nested forms and repeated pictures, keeps its original placement.
            # pypdf is constrained to major version 6; regression tests cover this API.
            writer._replace_object(ref, replacement)
            result["restored"].append({"resource": exported.name, "source": source_name,
                                        "before": native.size, "after": original.size})
    if result["restored"]:
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".pdf", delete=False) as output:
                temporary = Path(output.name)
                writer.write(output)
            # Complete the rewrite before replacing the intermediate Office PDF.
            import os
            os.replace(temporary, path)
        finally:
            if temporary:
                temporary.unlink(missing_ok=True)
    writer.close()
    return result
