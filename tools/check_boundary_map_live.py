"""Read-only public-service integration checks; save QA maps outside case folders.

Run from Wizard_Files. Optional --banks-workbook selects an existing real MLS
workbook; bundled MLS demo has intentionally inconsistent sample data.
"""
from pathlib import Path
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import sys
import traceback

WF = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WF))
from core.boundary_map import read_map_case, PublicMapClient, fetch_map_data
from core.boundary_map_render import default_settings, render_map, pdf_png


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--banks-workbook', type=Path)
    parser.add_argument('--output', type=Path, default=WF.parent / 'build' / 'map_qa')
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    cases = {key: next((WF / 'sample_input' / key).glob('*FILLED_SAMPLE.xlsx'))
             for key in ('mbl', 'mls_renewal', 'mbl_renewal')}
    if args.banks_workbook:
        cases['mls'] = args.banks_workbook
    def fetch(key, workbook):
        print(key, 'reading', flush=True)
        data_file = args.output / (key + '_data.json')
        case = read_map_case(workbook, key)
        if data_file.exists():
            data = json.loads(data_file.read_text())
            if data['case']['workbook_hash'] == case.workbook_hash:
                return key, data
        client = PublicMapClient(args.output / 'cache', progress=lambda msg: print(key, msg, flush=True))
        data = fetch_map_data(case, client)
        data_file.write_text(json.dumps(data), encoding='utf-8')
        return key, data
    results = {}
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(fetch, key, path): key for key, path in cases.items()}
        for future in as_completed(futures):
            key = futures[future]
            try:
                _, data = future.result()
                settings = default_settings(data)
                rendered = render_map(data, settings)
                (args.output / (key + '.pdf')).write_bytes(rendered['pdf'])
                (args.output / (key + '.png')).write_bytes(pdf_png(rendered['pdf'], 144))
                results[key] = {'properties': len(data['matched']), 'blocking': rendered['blocking'],
                                'warnings': rendered['issues'], 'settings': settings,
                                'footer_lines': rendered['footer_lines']}
                print(key, json.dumps(results[key]), flush=True)
            except Exception as e:
                results[key] = {'error': str(e)}
                traceback.print_exc()
    (args.output / 'results.json').write_text(json.dumps(results, indent=2), encoding='utf-8')
    return 1 if any(r.get('error') or r.get('blocking') for r in results.values()) else 0


if __name__ == '__main__':
    raise SystemExit(main())
