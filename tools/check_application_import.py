"""Audit actual intake cases without changing the source PDFs or templates."""
import argparse
import json
from pathlib import Path

from core.application_import import find_parcels,save_calculation,record_pdf_facts
from core.application_ocr import read_application


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--cases');parser.add_argument('--out',required=True)
    parser.add_argument('--read-pdf');parser.add_argument('--read-record')
    args=parser.parse_args();out=Path(args.out).resolve();out.mkdir(parents=True,exist_ok=True)
    if args.read_pdf or args.read_record:
        source=args.read_pdf or args.read_record
        reader=read_application if args.read_pdf else record_pdf_facts
        result=reader(source,lambda s:print(s,flush=True))
        (out/'reading.json').write_text(json.dumps(result,indent=2),encoding='utf8')
        print('Reading saved in',out/'reading.json',flush=True)
        return 0
    if not args.cases:parser.error('Provide --cases, --read-pdf or --read-record.')
    cases=json.loads(Path(args.cases).read_text(encoding='utf8'))
    report=[]
    for case in cases:
        print('Checking',case['name'],flush=True)
        application=json.loads(Path(case['application']).read_text(encoding='utf8'))
        data=find_parcels(case['fields'],out/'cache',application['property_hints'],lambda s:print(s,flush=True),use_cache=True)
        selected=[p for p in data['parcels'] if p['include']]
        (out/(case['name']+'_parcels.json')).write_text(json.dumps(data,indent=2),encoding='utf8')
        for p in selected:
            override=case.get('property_overrides',{}).get(p['account'],{})
            if override:p.update(override)
        path=save_calculation(out/(case['name']+'_Calculation.xlsx'),case['fields'],data,application,wizard=case.get('wizard','mls'))
        report.append({'name':case['name'],'selected':len(selected),'nearby':len(data['parcels']),
                       'accounts':[p['account'] for p in selected],'workbook':str(path),'warnings':data['warnings']})
        print(case['name'],len(selected),'selected',flush=True)
    (out/'intake_results.json').write_text(json.dumps(report,indent=2),encoding='utf8')
    return 0


if __name__=='__main__':raise SystemExit(main())
