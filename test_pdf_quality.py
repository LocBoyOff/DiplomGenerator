from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from PIL import Image, ImageChops
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (BooleanObject, DecodedStreamObject, DictionaryObject,
                           NameObject, NumberObject)

from pdf_quality import restore_original_images


class QualityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='pdf_quality_')
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.pptx = self.root / 'template.pptx'
        self.pdf = self.root / 'document.pdf'

    def pattern(self, size=(320, 200)):
        image = Image.new('RGBA', size)
        pixels = image.load()
        for x in range(size[0]):
            for y in range(size[1]):
                pixels[x, y] = (x * 7 % 256, y * 5 % 256, (x + y) * 3 % 256,
                                0 if x < 10 else 255)
        return image

    def source(self, *images):
        with ZipFile(self.pptx, 'w') as archive:
            for index, image in enumerate(images):
                data = BytesIO()
                image.save(data, format='PNG')
                archive.writestr(f'ppt/media/image{index}.png', data.getvalue())

    def exported(self, image):
        writer = PdfWriter()
        page = writer.add_blank_page(600, 400)
        jpeg = BytesIO()
        image.convert('RGB').save(jpeg, format='JPEG', quality=85)
        stream = DecodedStreamObject()
        stream.set_data(jpeg.getvalue())
        stream.update({NameObject('/Type'): NameObject('/XObject'),
                       NameObject('/Subtype'): NameObject('/Image'),
                       NameObject('/Width'): NumberObject(image.width),
                       NameObject('/Height'): NumberObject(image.height),
                       NameObject('/ColorSpace'): NameObject('/DeviceRGB'),
                       NameObject('/BitsPerComponent'): NumberObject(8),
                       NameObject('/Filter'): NameObject('/DCTDecode'),
                       NameObject('/Interpolate'): BooleanObject(True)})
        mask = DecodedStreamObject()
        mask.set_data(image.getchannel('A').tobytes())
        mask.update({NameObject('/Type'): NameObject('/XObject'),
                     NameObject('/Subtype'): NameObject('/Image'),
                     NameObject('/Width'): NumberObject(image.width),
                     NameObject('/Height'): NumberObject(image.height),
                     NameObject('/ColorSpace'): NameObject('/DeviceGray'),
                     NameObject('/BitsPerComponent'): NumberObject(8)})
        stream[NameObject('/SMask')] = writer._add_object(mask.flate_encode())
        font = DictionaryObject({NameObject('/Type'): NameObject('/Font'),
                                 NameObject('/Subtype'): NameObject('/Type1'),
                                 NameObject('/BaseFont'): NameObject('/Helvetica')})
        page[NameObject('/Resources')] = DictionaryObject({
            NameObject('/XObject'): DictionaryObject({NameObject('/Bg'): writer._add_object(stream)}),
            NameObject('/Font'): DictionaryObject({NameObject('/F1'): writer._add_object(font)}),
        })
        content = DecodedStreamObject()
        content.set_data(b'q 600 0 0 400 0 0 cm /Bg Do Q\nBT /F1 24 Tf 50 200 Td (Diploma) Tj ET')
        page[NameObject('/Contents')] = writer._add_object(content)
        writer.write(self.pdf)
        writer.close()

    def test_restore_exact_pixels_alpha_text_and_placement(self):
        original = self.pattern()
        self.source(original)
        self.exported(original)
        before = PdfReader(self.pdf)
        content = before.pages[0].get_contents().get_data()
        report = restore_original_images(self.pptx, self.pdf)
        after = PdfReader(self.pdf)
        self.assertEqual(len(report['restored']), 1)
        self.assertEqual(after.pages[0].extract_text(), 'Diploma')
        self.assertEqual(after.pages[0].get_contents().get_data(), content)
        self.assertEqual(after.pages[0].mediabox, before.pages[0].mediabox)
        image = after.pages[0].images[0]
        self.assertIsNone(ImageChops.difference(image.image.convert('RGBA'), original).getbbox(alpha_only=False))
        self.assertEqual(image.indirect_reference.get_object()['/Filter'], '/FlateDecode')
        self.assertTrue(bool(image.indirect_reference.get_object()['/Interpolate']))

    def test_restore_downsampled_image(self):
        original = Image.new('RGBA', (640, 400))
        original.putdata([(x * 255 // 640, y * 255 // 400, 120, 255)
                          for y in range(400) for x in range(640)])
        self.source(original)
        self.exported(original.resize((320, 200), Image.Resampling.LANCZOS))
        report = restore_original_images(self.pptx, self.pdf)
        self.assertEqual(len(report['restored']), 1)
        self.assertEqual(PdfReader(self.pdf).pages[0].images[0].image.size, original.size)

    def test_do_not_replace_a_different_or_ambiguous_picture(self):
        first = Image.new('RGBA', (320, 200), (255, 0, 0, 255))
        second = Image.new('RGBA', (320, 200), (0, 0, 255, 255))
        self.source(first)
        self.exported(second)
        before = self.pdf.read_bytes()
        report = restore_original_images(self.pptx, self.pdf)
        self.assertFalse(report['restored'])
        self.assertEqual(self.pdf.read_bytes(), before)
        # Two near-identical images cannot safely be distinguished after JPEG.
        self.source(first, Image.new('RGBA', first.size, (254, 0, 0, 255)))
        self.exported(first)
        report = restore_original_images(self.pptx, self.pdf)
        self.assertFalse(report['restored'])

    def test_no_source_images(self):
        self.source()
        self.exported(self.pattern())
        before = self.pdf.read_bytes()
        self.assertFalse(restore_original_images(self.pptx, self.pdf)['restored'])
        self.assertEqual(self.pdf.read_bytes(), before)

    def test_do_not_undo_template_brightness_effect(self):
        image = self.pattern()
        self.source(image)
        with ZipFile(self.pptx, 'a') as archive:
            archive.writestr('ppt/slides/slide1.xml',
                             '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                             'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                             '<a:blip r:embed="rId1"><a:lum bright="500"/></a:blip></p:sld>')
            archive.writestr('ppt/slides/_rels/slide1.xml.rels',
                             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                             '<Relationship Id="rId1" Target="../media/image0.png"/></Relationships>')
        self.exported(image)
        before = self.pdf.read_bytes()
        self.assertFalse(restore_original_images(self.pptx, self.pdf)['restored'])
        self.assertEqual(self.pdf.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
