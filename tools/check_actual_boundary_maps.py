"""Exercise Create Boundary Map and Use This Map with four saved workbooks.

The JSON input maps wizard keys to workbook paths. Omit mbl to exercise the
installed instructional example. All outputs and session state stay in --out.
"""
from pathlib import Path
import argparse
import importlib
import json
import shutil
import time
import tkinter as tk
from contextlib import nullcontext
from unittest.mock import patch
from shapely.affinity import affine_transform, rotate
from shapely.geometry import Polygon
from pypdf import PdfReader

from core import session
from core.boundary_map import fingerprint, application_boundary, parcel_polygon, PublicMapClient
from core.workbooks import calculation_template_path


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cases',type=Path,required=True)
    parser.add_argument('--out',type=Path,required=True)
    parser.add_argument('--cache',type=Path)
    parser.add_argument('--offline',action='store_true',help='Require saved public data; fail any network lookup')
    args=parser.parse_args()
    cases=json.loads(args.cases.read_text(encoding='utf-8'))
    if isinstance(cases,dict):
        cases.setdefault('mbl',str(calculation_template_path('mbl')))
        cases=[{'id':key,'wizard':key,'workbook':path} for key,path in cases.items()]
    args.out.mkdir(parents=True,exist_ok=True)
    names={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
    root=tk.Tk(); root.withdraw()
    results={}
    def descendants(w):
        for child in w.winfo_children():
            yield child
            yield from descendants(child)
    def fail_dialog(title,message,**kwargs):
        raise AssertionError(f'{title}: {message}')
    try:
        with (patch.object(PublicMapClient,'request',side_effect=AssertionError('Uncached lookup during offline verification')) if args.offline else nullcontext()), patch.object(session,'SETTINGS_PATH',args.out/'test-session.json'), patch.object(session,'load',return_value={}), patch('tkinter.messagebox.showerror',fail_dialog):
            for entry in cases:
                key=entry['wizard']; name=names[key]; case_id=entry['id']
                owner=None
                try:
                    source=Path(entry['workbook']); before=fingerprint(source)
                    output=args.out/case_id; output.mkdir(exist_ok=True)
                    if args.cache:
                        shutil.copytree(args.cache,output/'.boundary_map_cache',dirs_exist_ok=True)
                    owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),name)(root)
                    owner.pack(fill='both',expand=True)
                    if hasattr(owner,'vars'):
                        owner.vars['excel'].set(str(source)); owner.vars['output_folder'].set(str(output))
                    else:
                        owner.files.update(workbook=str(source),out=str(output))
                    button=next(w for w in descendants(owner) if isinstance(w,tk.Button) and w.cget('text')=='Create Boundary Map')
                    button.invoke()
                    preview=owner._boundary_map_preview
                    preview.withdraw()
                    deadline=time.monotonic()+120
                    while preview.rendered is None and time.monotonic()<deadline:
                        root.update()
                        if preview.status.get().startswith(('Map not created','Could not draw')):
                            raise AssertionError(preview.review.get('1.0','end'))
                        time.sleep(.02)
                    if preview.rendered is None:
                        raise AssertionError('Timed out waiting for the actual map preview.')
                    if preview.rendered['blocking']:
                        raise AssertionError(preview.rendered['blocking'])
                    boundary=application_boundary(preview.data,preview.settings['include_street'])
                    if any(not boundary.buffer(.001).covers(parcel_polygon(m['feature'])) for m in preview.data['matched']):
                        raise AssertionError('A listed property, including EXC, is outside the application boundary.')
                    exc_count=sum(m['property']['land_use']=='EXC' for m in preview.data['matched'])
                    count=preview.rendered['selected_count']
                    labels=[l for l in preview.rendered['labels'] if l['key'].startswith('parcel:')]
                    addressed={f'parcel:{m["property"]["row"]}':m for m in preview.data['matched']}
                    if {l['key'] for l in labels}!=set(addressed):
                        raise AssertionError('Address labels do not match the addressed worksheet parcels.')
                    for label in labels:
                        matched=addressed[label['key']]
                        polygon=rotate(parcel_polygon(matched['feature']),preview.rendered['rotation'],origin=preview.rendered['rotation_origin'])
                        polygon=affine_transform(polygon,preview.rendered['page_transform'])
                        if label['text']!=(matched['property']['label'] or matched['property']['number'] or '0') or not polygon.covers(Polygon(label['corners'])):
                            raise AssertionError('An address is not entirely inside its own matched parcel.')
                    (output/'parcel_data.json').write_text(json.dumps(preview.data),encoding='utf-8')
                    downloaded=output/'Downloaded_Boundary_Map.pdf'
                    with patch('tkinter.filedialog.asksaveasfilename',return_value=str(downloaded)):
                        preview.download.invoke()
                    if downloaded.read_bytes()!=preview.rendered['pdf']:
                        raise AssertionError('Downloaded PDF differs from the reviewed preview.')
                    preview.use.invoke()
                    selected=owner.vars['map'].get() if hasattr(owner,'vars') else owner.files['map']
                    pdf=Path(selected)
                    if not pdf.is_file() or not pdf.read_bytes().startswith(b'%PDF'):
                        raise AssertionError('Use This Map did not save and attach a PDF.')
                    owner._check_boundary_map()
                    words=set(PdfReader(pdf).pages[0].extract_text().split())
                    if words.intersection({'SFR','MF','COM','VAC','EXC','Unaddressed'}):
                        raise AssertionError('A land-use label remains in the PDF.')
                    if fingerprint(source)!=before:
                        raise AssertionError('The source workbook changed during the check.')
                    if entry.get('expected_error'):
                        raise AssertionError('The known conflicting workbook was unexpectedly accepted.')
                    results[case_id]={'reference':preview.data['case']['reference'],'block':preview.data['case']['block'],'properties':count,'exc_inside_boundary':exc_count,'addresses_inside_own_lot':len(labels),'unaddressed_lots_zero':sum(l['text']=='0' for l in labels),'downloaded_pdf':str(downloaded),'warnings':preview.rendered['issues'],'pdf':str(pdf),'workbook_unchanged':True,'create_and_attach':'passed'}
                except Exception as e:
                    if entry.get('expected_error') and entry['expected_error'] in str(e) and fingerprint(source)==before:
                        results[case_id]={'input_error':str(e),'correctly_rejected':True,'workbook_unchanged':True}
                    else:
                        results[case_id]={'error':str(e)}
                finally:
                    if owner is not None:
                        preview=getattr(owner,'_boundary_map_preview',None)
                        if preview is not None and preview.winfo_exists(): preview.close()
                        owner.destroy()
                    root.update()
                print(case_id,json.dumps(results[case_id]),flush=True)
    finally:
        root.destroy()
    (args.out/'results.json').write_text(json.dumps(results,indent=2),encoding='utf-8')
    return int(any(r.get('error') for r in results.values()))


if __name__=='__main__': raise SystemExit(main())
