"""Opt-in native print round trip against synthetic input only."""
import argparse
from collections import Counter
import json
from pathlib import Path
import re
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from print_layout import preview, export, sha256


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--file', type=Path, default=Path(__file__).parent/'fixtures/A_B_AA_BB.mpp')
    parser.add_argument('--extended', action='store_true', help='Also validate A4 portrait/week-scale settings')
    args = parser.parse_args()
    # Source must be a disposable synthetic plan. Never pass business files here.
    before = sha256(args.file)
    root = Path(__file__).resolve().parents[1]/'work'/f'print-test-{time.time_ns()}'
    results = []
    profiles = [('A3', 'landscape', 'days'), ('A4', 'landscape', 'days')]
    if args.extended:
        profiles.append(('A4','portrait','weeks'))
    for paper, orientation, timescale in profiles:
        folder = root/f'{paper}-{orientation}-{timescale}'
        profile = {'paper':paper, 'header':'Synthetic Gantt print test', 'footer':'Test data only',
                   'inside_progress':True, 'orientation':orientation, 'timescale':timescale}
        proposal = preview(args.file,folder,profile)
        result = export(args.file,folder,profile,proposal['preview_sha256'])
        assert result['saved_data_unchanged'] and result['source_unchanged'], result
        assert result['saved_layout_text_unchanged'], result
        assert result['pdf']['basic_checks_passed'], result
        import fitz
        with fitz.open(result['pdf_path']) as pdf:
            # With default columns, each non-milestone percentage must occur once
            # in the table and once inside the bar (milestones only in the table).
            values = Counter(w[4] for page in pdf for w in page.get_text('words') if re.fullmatch(r'\d+%',w[4]))
            expected = Counter(result['table_percentages']) + Counter(result['inside_percentages'])
            assert values==expected, {'actual':dict(values),'expected':dict(expected)}
            for page in pdf:
                page.get_pixmap(matrix=fitz.Matrix(1,1)).save(folder/f'page-{page.number+1}.png')
        results.append(result)
    assert sha256(args.file)==before
    (root/'results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'ok':True,'output':str(root)},ensure_ascii=False))


if __name__=='__main__':
    main()
