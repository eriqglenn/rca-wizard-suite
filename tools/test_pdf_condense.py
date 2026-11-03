"""HCAD rotation, clipping, and one-page attachment regression checks."""
from pathlib import Path
from collections import Counter
import argparse
import hashlib
import importlib
import json
import sys
import tempfile
import unittest

import fitz

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from core.pdf_condense import condense_pdf_to_one_page, condense_hcad_pdfs


def words(doc):
    return Counter(word for page in doc for word in page.get_text().split())


def upright(doc):
    lines=[line for page in doc for block in page.get_text('dict')['blocks']
           for line in block.get('lines',[]) if sum(c.isalnum() for span in line.get('spans',[]) for c in span['text']) >= 2]
    return bool(lines) and all(abs(line['dir'][0]-1)<.001 and abs(line['dir'][1])<.001 for line in lines)


class CondenseTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.folder=Path(self.tmp.name)

    def source(self, rotation=0, blank_second=False, crop=False):
        doc=fitz.open()
        for i in range(2):
            page=doc.new_page(width=792,height=612)
            if not (i==1 and blank_second):
                for j,(x,y) in enumerate([(45,60),(670,60),(45,555),(670,555)]):
                    page.insert_text((x,y),f'P{i+1}CORNER{j+1}',fontsize=9)
                page.draw_rect(fitz.Rect(40,40,752,570),color=(0,0,0),width=1)
            if crop:
                page.set_cropbox(fitz.Rect(20,20,772,592))
        source=self.folder/f'source_{rotation}_{crop}_{blank_second}.pdf'
        doc.save(source);doc.close()
        if rotation:
            # Construct an upright display with saved rotation, preserving the
            # crop box in the fixture as well as in the production code.
            from pypdf import PdfReader, PdfWriter
            import io
            reader=PdfReader(io.BytesIO(source.read_bytes()));writer=PdfWriter()
            for page in reader.pages:
                page.rotation=(360-rotation)%360
                page.transfer_rotation_to_content()
                page.rotation=rotation
                writer.add_page(page)
            with source.open('wb') as out:writer.write(out)
        return source

    def check(self, source):
        before=source.read_bytes()
        out=self.folder/(source.stem+'_condensed.pdf')
        self.assertEqual(condense_pdf_to_one_page(source,out),out)
        self.assertEqual(source.read_bytes(),before)
        with fitz.open(source) as src,fitz.open(out) as result:
            self.assertEqual(len(result),1)
            self.assertEqual(result[0].rotation,0)
            self.assertEqual(tuple(result[0].rect),(0,0,612,792))
            self.assertEqual(words(src),words(result))
            self.assertTrue(upright(result))
            for word in result[0].get_text('words'):
                self.assertTrue(fitz.Rect(17,17,595,775).contains(fitz.Rect(word[:4])),word)
        return out

    def test_rotated_records_keep_every_corner_and_read_upright(self):
        for angle in (0,90,180,270):
            for crop in (False,True):
                with self.subTest(angle=angle,crop=crop):self.check(self.source(angle,crop=crop))

    def test_single_remaining_wide_page_fits_after_blank_page_removed(self):
        for angle in (0,90,180,270):self.check(self.source(angle,blank_second=True))

    def test_single_page_is_passed_through_unchanged(self):
        source=self.source(90)
        with fitz.open(source) as doc:
            doc.delete_page(1)
            single=self.folder/'single.pdf';doc.save(single)
        before=single.read_bytes()
        self.assertEqual(condense_pdf_to_one_page(single,self.folder/'unused.pdf'),single)
        self.assertEqual(single.read_bytes(),before)

    def test_blank_or_invalid_input_is_preserved(self):
        source=self.folder/'blank.pdf'
        with fitz.open() as doc:
            doc.new_page();doc.new_page();doc.save(source)
        self.assertEqual(condense_pdf_to_one_page(source,self.folder/'blank_out.pdf'),source)
        invalid=self.folder/'invalid.pdf';invalid.write_bytes(b'invalid test')
        self.assertEqual(condense_pdf_to_one_page(invalid,self.folder/'invalid_out.pdf'),invalid)

    def test_batch_preserves_record_order(self):
        paths=[self.source(90),self.source(270)]
        outputs=condense_hcad_pdfs(paths,self.folder/'batch')
        self.assertEqual(len(outputs),2)
        self.assertIn('01_'+paths[0].stem,outputs[0].name)
        self.assertIn('02_'+paths[1].stem,outputs[1].name)

    def test_scanned_records_keep_image_content(self):
        from PIL import Image
        import io
        image=Image.new('RGB',(900,600),'white')
        for x in range(100,800):
            for y in range(50,100):image.putpixel((x,y),(0,0,0))
        stream=io.BytesIO();image.save(stream,format='PNG')
        source=self.folder/'scan.pdf'
        with fitz.open() as doc:
            for i in range(2):
                page=doc.new_page(width=792,height=612)
                page.insert_image(fitz.Rect(20,20,770,570),stream=stream.getvalue())
                page.set_rotation(270);page.remove_rotation();page.set_rotation(90)
            doc.save(source)
        out=condense_pdf_to_one_page(source,self.folder/'scan_out.pdf')
        with fitz.open(out) as doc:
            self.assertEqual(len(doc),1)
            images=doc[0].get_image_info()
            self.assertEqual(len(images),2)
            for image in images:self.assertTrue(fitz.Rect(17,17,595,775).contains(fitz.Rect(image['bbox'])))

    def test_all_four_wizards_use_shared_condensing_function(self):
        for key in ('mls.mailout_engine','mls.templates_2026','mls_renewal.wizard','mbl.wizard','mbl.mbl_mailout_engine','mbl_renewal.wizard'):
            module=importlib.import_module('wizards.'+key)
            self.assertIs(module.condense_hcad_pdfs,condense_hcad_pdfs)


def check_records(folder, output):
    output.mkdir(parents=True,exist_ok=True)
    records=[]
    for source in sorted(folder.glob('*.pdf')):
        before=hashlib.sha256(source.read_bytes()).hexdigest()
        target=output/source.name
        actual=condense_pdf_to_one_page(source,target)
        assert actual==target,source.name
        with fitz.open(source) as src,fitz.open(target) as doc:
            assert len(doc)==1 and upright(doc),source.name
            assert words(src)==words(doc),(source.name,words(src)-words(doc))
            for word in doc[0].get_text('words'):
                assert fitz.Rect(17,17,595,775).contains(fitz.Rect(word[:4])),(source.name,word)
            records.append({'source':str(source),'source_pages':len(src),'output':str(target),'upright':True,'all_text_preserved':True,'source_sha256':before})
        assert hashlib.sha256(source.read_bytes()).hexdigest()==before
    assert records
    (output/'verification.json').write_text(json.dumps(records,indent=2),encoding='utf-8')
    print(f'Verified {len(records)} real records: upright, all text retained, originals unchanged.')


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--records',type=Path)
    parser.add_argument('--out',type=Path)
    args=parser.parse_args()
    result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(CondenseTests))
    if not result.wasSuccessful():return 1
    if args.records:
        if not args.out:parser.error('--out is required with --records')
        check_records(args.records,args.out)
    return 0


if __name__=='__main__':raise SystemExit(main())
