"""Restore original raster pixels after Office's lossy PDF conversion.

Only confidently matched image resources are replaced. The PDF content streams,
font resources, placement matrices, clipping and vector graphics stay intact.
"""
from io import BytesIO
from copy import copy
from collections import OrderedDict
import hashlib
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


def _resources(resources, seen=None, prefix=""):
    seen = set() if seen is None else seen
    resources = resources.get_object() if hasattr(resources, "get_object") else resources
    objects = resources.get("/XObject", {})
    objects = objects.get_object() if hasattr(objects, "get_object") else objects
    for name, reference in objects.items():
        obj = reference.get_object()
        identity = getattr(reference, "idnum", id(obj))
        if identity in seen:
            continue
        seen.add(identity)
        if obj.get("/Subtype") == "/Image":
            yield prefix + str(name), reference, obj
        elif obj.get("/Subtype") == "/Form":
            yield from _resources(obj.get("/Resources", {}), seen, prefix + str(name) + "/")


def _fast_image(obj):
    """Decode Office RGB/gray resources without re-encoding them as JPEG2000."""
    colour = obj.get("/ColorSpace")
    mode = {"/DeviceRGB": "RGB", "/DeviceGray": "L"}.get(str(colour))
    if not mode or obj.get("/BitsPerComponent") != 8 or obj.get("/ImageMask") or obj.get("/Mask"):
        return None
    if obj.get("/Decode") not in (None, [0, 1] * len(mode)):
        return None
    size = int(obj["/Width"]), int(obj["/Height"])
    filters = obj.get("/Filter", [])
    if isinstance(filters, str):
        filters = [filters]
    data = obj.get_data()
    if "/DCTDecode" in filters or "/JPXDecode" in filters:
        with Image.open(BytesIO(data)) as decoded:
            image = decoded.convert("RGBA")
    else:
        image = Image.frombytes(mode, size, data).convert("RGBA")
    mask = obj.get("/SMask")
    if mask and str(mask) != "/None":
        if mask.get_object().get("/Matte") is not None:
            return None
        alpha = _fast_image(mask.get_object())
        if alpha is None or alpha.size != size:
            return None
        image.putalpha(alpha.getchannel("R"))
    return image


def _resource_key(obj):
    digest = hashlib.sha256()
    digest.update(obj._data)
    for key in ("/Width", "/Height", "/BitsPerComponent", "/ColorSpace", "/Filter", "/Decode",
                "/DecodeParms", "/Mask", "/SMaskInData", "/Matte"):
        digest.update(str(obj.get(key)).encode())
    mask = obj.get("/SMask")
    if mask and str(mask) != "/None":
        digest.update(_resource_key(mask.get_object()))
    return digest.digest()


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
            digest = hashlib.sha256(rgba.tobytes()).digest()
            if digest in seen:
                continue
            seen.add(digest)
            sources.append((name, rgba, _thumbnail(rgba)))
    return sources


def _lossless_stream(writer, image, cache=None, key=None):
    if cache is not None and key in cache:
        rgb, alpha = cache[key]
        stream = copy(rgb)
        if alpha is not None:
            stream[NameObject("/SMask")] = writer._add_object(copy(alpha))
        return stream
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
    mask = None
    if alpha.getextrema() != (255, 255):
        mask = DecodedStreamObject()
        mask.set_data(alpha.tobytes())
        mask.update({NameObject("/Type"): NameObject("/XObject"),
                     NameObject("/Subtype"): NameObject("/Image"),
                     NameObject("/Width"): NumberObject(rgba.width),
                     NameObject("/Height"): NumberObject(rgba.height),
                     NameObject("/ColorSpace"): NameObject("/DeviceGray"),
                     NameObject("/BitsPerComponent"): NumberObject(8)})
        mask = mask.flate_encode()
    stream = stream.flate_encode()
    if cache is not None:
        # These prototypes contain no object references from any output PDF.
        cache[key] = (stream, mask)
        stream = copy(stream)
    if mask is not None:
        stream[NameObject("/SMask")] = writer._add_object(copy(mask))
    return stream


def restore_original_images(pptx_path, pdf_path, *, restorer=None):
    sources = restorer.sources if restorer is not None else _sources(pptx_path)
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
        for resource_name, ref, obj in _resources(page.get("/Resources", {})):
            if not hasattr(ref, "idnum") or ref.idnum in processed:
                continue
            processed.add(ref.idnum)
            key = _resource_key(obj)
            decision = restorer.matches.get(key) if restorer is not None else None
            if decision is not None:
                source_name, original, unchanged = decision
            else:
                native = _fast_image(obj)
                if native is None:
                    # Preserve the full pypdf decoder for uncommon colour spaces.
                    exported = next((i for i in page.images if i.indirect_reference == ref), None)
                    if exported is None or exported.image is None:
                        result["unmatched"].append(resource_name)
                        continue
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
                if (not candidates or candidates[0][0] > 8
                        or (len(candidates) > 1 and candidates[1][0] - candidates[0][0] < 1)):
                    source_name, original, unchanged = None, None, False
                else:
                    _, source_name, original = candidates[0]
                    unchanged = (original.size == native.size and
                                 ImageChops.difference(original, native).getbbox(alpha_only=False) is None)
                if restorer is not None:
                    if len(restorer.matches) >= 64:
                        restorer.matches.popitem(last=False)
                    restorer.matches[key] = (source_name, original, unchanged)
            if original is None:
                result["unmatched"].append(resource_name)
                continue
            if unchanged:
                result["unchanged"] += 1
                continue
            replacement = _lossless_stream(writer, original,
                                            restorer.compressed if restorer is not None else None,
                                            source_name)
            for option in ("/Interpolate", "/Intent", "/OC"):
                if option in obj:
                    replacement[NameObject(option)] = obj[option]
            writer._replace_object(ref, replacement)
            result["restored"].append({"resource": resource_name, "source": source_name,
                                        "before": (int(obj["/Width"]), int(obj["/Height"])),
                                        "after": original.size})
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


class ImageRestorer:
    """Batch-scoped cache: decode and compress each original picture only once."""
    def __init__(self, pptx_path):
        self.sources = _sources(pptx_path)
        self.compressed = {}
        self.matches = OrderedDict()

    def restore(self, pdf_path):
        return restore_original_images(None, pdf_path, restorer=self)
