"""Create a synthetic A/B, AA/BB fixture with neutral document metadata."""
import argparse
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from com_bridge import create_plan
from plan_builder import verify_new_mpp


def neutral_metadata(path):
    import pythoncom
    # Standard COM property-set storage; no user/application preferences change.
    storage = pythoncom.StgOpenStorageEx(str(path), 18, 5, 0, pythoncom.IID_IPropertySetStorage)
    properties = storage.Open(pythoncom.FMTID_SummaryInformation, 18)
    properties.WriteMultiple((2, 4, 8), ('Project MCP for codex - A/B', 'Project MCP', 'Project MCP'))
    properties.Commit(0)
    properties = storage = None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path, default=Path(__file__).parent / 'fixtures' / 'A_B_AA_BB.mpp')
    args = parser.parse_args()
    source = Path(__file__).with_name('basic_plan.xml')
    preview = create_plan(str(source), str(args.output))
    create_plan(str(source), str(args.output), preview=False, source_sha256=preview['source_sha256'])
    neutral_metadata(args.output)
    assert verify_new_mpp(source, args.output)['ok']
    print('Synthetic fixture verified; tasks A/B, resources AA/BB.')


if __name__ == '__main__':
    main()
