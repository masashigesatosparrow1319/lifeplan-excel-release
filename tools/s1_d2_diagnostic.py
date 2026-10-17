"""Inspect D2 only: unchanged 01_function_gate.xlsx, all 24 expected caches."""
import argparse
import hashlib
import json
import ntpath
from pathlib import Path

from openpyxl import load_workbook
from tools import s1_compatibility as gate

SCHEMA = 'S1-D2-function-gate-v1'
FILE = '01_function_gate.xlsx'
EXPECTED = {address: value for _, address, value in gate.expected_cells('gate')}


def compare_values(values):
    return {address: dict(expected=value, actual=values.get(address))
            for address, value in EXPECTED.items() if values.get(address) != value}


def inspect(inputs, results):
    inputs, results = Path(inputs), Path(results)
    assert inputs.resolve() != results.resolve(), '原本と結果は別フォルダーです'
    records = json.loads((inputs / 'manifest.json').read_text(encoding='utf-8'))
    entry = next(e for e in records if e['file'] == FILE)
    assert entry['formula_cells'] == 24
    assert hashlib.sha256((inputs / FILE).read_bytes()).hexdigest() == entry['sha256'], '原本変更を検出しました'
    log = json.loads((results / 'd2-session.json').read_text(encoding='utf-8-sig'))
    assert log['schema'] == SCHEMA and log['input_sha256'] == entry['sha256']
    trace = log['trace']
    assert trace.get('mode') == 'D2_Com', '別段階の記録です'
    assert trace['identity_verified'] is True
    facts = trace['connection']
    assert facts['excel_version'] == '14.0' and facts['file_major_version'] == 14
    assert facts['process_id'] == trace['launched_process_id'] > 0
    assert ntpath.normcase(facts['executable_path']) == ntpath.normcase(trace['expected_path'])
    observations = log['observations']
    report = dict(s1_acceptance=False, d2_gate_ready=False, manual_warning_check_required=True,
                  cache='not_saved', checked_formula_cells=0,
                  observations=[dict(phase=s['phase'], calculation=s['calculation'], state=s['state'],
                                     calculation_before_values=s.get('calculation_before_values'), state_before_values=s.get('state_before_values'),
                                     all_24_values_matched=not compare_values(s['values']),
                                     mismatched_cells=compare_values(s['values'])) for s in observations])
    saved = results / FILE
    if saved.exists():
        checked = gate.verify_file(inputs / FILE, saved, 'gate')
        assert checked['checked_formula_cells'] == checked['formula_cells'] == 24
        w = load_workbook(saved, read_only=True)
        try:
            assert w.calculation is None or w.calculation.calcMode in (None, 'auto'), '保存設定が自動ではありません'
        finally:
            w.close()
        report.update(cache='all_24_fixed_values_matched', checked_formula_cells=24)
    by_phase = {s['phase']: s for s in observations}
    before = by_phase.get(log.get('pre_save_phase'))
    reopened = by_phase.get('reopened')
    save_phase = log.get('pre_save_phase') or ''
    def ready(sample):
        return sample is not None and sample['state'] == 0 and sample['calculation'] == -4105 and not compare_values(sample['values'])
    report['d2_gate_ready'] = bool(log['observation_completed'] and log.get('full_rebuild_calls') == 1
        and log.get('first_error') is None and not log['cleanup_errors'] and not log['progress_log_errors']
        and save_phase.startswith('observed_')
        and trace.get('automatic_restore_verified') is True and ready(before) and ready(reopened)
        and report['cache'] == 'all_24_fixed_values_matched')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--inputs', default='output/s1_compatibility')
    parser.add_argument('--results', required=True)
    args = parser.parse_args()
    print(json.dumps(inspect(args.inputs, args.results), ensure_ascii=False, indent=2))
