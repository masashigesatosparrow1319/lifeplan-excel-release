"""No-Excel Phase A controller. A fresh CLI process stops after Phase A."""
import argparse
from contextlib import redirect_stderr, redirect_stdout
import io
import json
from pathlib import Path
import re
import stat

from . import windows_preflight as pre, no_excel_tests as tests

BASE_TESTS = (
    'tests/test_excel2024_windows_preflight.py',
    'tests/test_excel2024_preopen_input.py', 'tests/test_excel2024_preopen_diagnostic.py',
    'tests/test_excel2024_functions24.py', 'tests/test_excel2024_e1.py',
    'tests/test_excel2024_validation.py', 'tests/test_excel2024_calculation_wait.py',
    'tests/test_excel2024_doctor_aggregation.py', 'tests/test_excel2024_cleanup.py',
    'tests/test_p1.py', 'tests/test_p2a.py',
)
PHASE_TESTS = BASE_TESTS + (
    'tests/test_excel2024_saved_verifier.py', 'tests/test_excel2024_offline_tiny_e1.py',
)
COUNTS = dict(passed=947, failed=0, errors=0, skipped=35)
FLAGS = ('excel_started', 'com_connected', 'approval_id_used', 'gate2_acceptance', 'product_acceptance')


class Stop(Exception):
    def __init__(self, code, review=False):
        self.code, self.review = code, review


def require(value, code, review=False):
    if not value:
        raise Stop(code, review)


class Parser(argparse.ArgumentParser):
    def error(self, message):
        raise Stop('ARGUMENTS_INVALID')


class Discard(io.TextIOBase):
    """Do not buffer or expose exceptions/third-party output with private data."""
    def write(self, text):
        return len(text)

    def flush(self):
        pass


def flags_ok(value, *, legacy_tests=False):
    for key in FLAGS:
        # G2_NO_EXCEL_TESTS_V1 has no product_acceptance field. An explicit
        # true/None is never accepted; the controller itself always emits false.
        if legacy_tests and key == 'product_acceptance' and key not in value:
            continue
        require(value.get(key) is False, 'UNSAFE_STAGE_FLAGS')
    require('reason' not in value and 'error_type' not in value and 'report_write' not in value,
            'STAGE_ERROR_OR_REPORT_FAILURE')
    require(value.get('next_action') == 'STOP_PM_REVIEW', 'UNSAFE_NEXT_ACTION')


def execute(profile, preflight_scratch, tests_scratch, expected, confirmed=False):
    result = dict(schema='G2_PHASE_A_AUTO_V1', status='BLOCKED', reason='NOT_COMPLETED',
                  success=False, exit_code=2, preflight_status='NOT_RUN', tests_status='NOT_RUN',
                  test_files=13, counts=None, unknown=35, unverified='P1_NOT_VERIFIED',
                  guard_violation_count=None, junit_sha256=None, release_source_sha256=None,
                  next_action='STOP_PM_REVIEW', **{key: False for key in FLAGS})
    try:
        # Existing APIs are quiet today; suppress incidental future output too.
        with redirect_stdout(Discard()), redirect_stderr(Discard()):
            require(__debug__, 'ASSERTS_REQUIRED')
            require(confirmed is True, 'OWNER_CONFIRMATION_REQUIRED')
            require(type(expected) is str and re.fullmatch('[0-9a-f]{64}', expected), 'SOURCE_PIN_REQUIRED')
            require(type(tests.TESTS) is tuple and tests.TESTS == BASE_TESTS
                    and len(set(PHASE_TESTS)) == 13, 'TEST_LIST_CHANGED')
            profile = pre.path_check(profile, outside=True)
            info = profile.stat()
            require(stat.S_ISREG(info.st_mode) and info.st_nlink == 1, 'PROFILE_REGULAR_FILE_REQUIRED')
            first = pre.path_check(preflight_scratch, outside=True)
            second = pre.path_check(tests_scratch, outside=True)
            require(not first.is_relative_to(second) and not second.is_relative_to(first), 'SCRATCH_OVERLAP')
            # Validate BOTH empty directories before the first stage writes.
            pre.scratch_check(first, profile, confirmed)
            pre.scratch_check(second, profile, confirmed)
            result['preflight_status'] = 'BLOCKED'
            value = pre.execute(profile, first, expected, confirmed)
            require(type(value) is dict, 'PREFLIGHT_RESULT_INVALID')
            result['preflight_status'] = 'BLOCKED'
            require(value.get('status') == 'PREFLIGHT_CHECKED', 'PREFLIGHT_BLOCKED')
            flags_ok(value)
            require(value.get('schema') == 'G2_WINDOWS_NO_EXCEL_PREFLIGHT_V1', 'PREFLIGHT_SCHEMA_INVALID')
            source = value.get('checks', {}).get('source', {})
            require(source.get('release_source_sha256') == expected, 'PREFLIGHT_SOURCE_MISMATCH')
            # Existing environment function reports bitness but does not reject 32bit.
            environment = value.get('checks', {}).get('environment', {})
            require(type(environment.get('python_bits')) is int and environment['python_bits'] == 64,
                    'PYTHON_X64_REQUIRED')
            result.update(preflight_status='PREFLIGHT_CHECKED', release_source_sha256=expected)
            original = tests.TESTS
            try:
                tests.TESTS = PHASE_TESTS
                result['tests_status'] = 'BLOCKED'
                value = tests.execute(second, expected, confirmed)
            finally:
                tests.TESTS = original
            require(type(value) is dict, 'TEST_RESULT_INVALID')
            result['tests_status'] = 'BLOCKED'
            require(value.get('status') == 'NO_EXCEL_TESTS_COMPLETE', 'TESTS_BLOCKED')
            flags_ok(value, legacy_tests=True)
            require(value.get('schema') == 'G2_NO_EXCEL_TESTS_V1', 'TEST_SCHEMA_INVALID')
            require(value.get('tests') == list(PHASE_TESTS), 'TEST_LIST_RESULT_CHANGED', True)
            require(value.get('release_source_sha256') == expected, 'TEST_SOURCE_MISMATCH')
            require(type(value.get('pytest_exit_code')) is int and value['pytest_exit_code'] == 0,
                    'PYTEST_EXIT_NOT_ZERO')
            violations = value.get('guard_violations')
            require(type(violations) is list, 'GUARD_RESULT_MISSING')
            result['guard_violation_count'] = len(violations)
            require(not violations, 'GUARD_VIOLATION')
            counts = value.get('counts')
            require(type(counts) is dict and set(counts) == set(COUNTS)
                    and all(type(v) is int and v >= 0 for v in counts.values()), 'COUNTS_INVALID')
            result['counts'] = counts.copy()
            require(counts == COUNTS, 'EXACT_COUNTS_REQUIRE_REVIEW', True)
            require(value.get('junit') == 'tests.xml' and type(value.get('junit_sha256')) is str
                    and re.fullmatch('[0-9a-f]{64}', value['junit_sha256']), 'JUNIT_RESULT_INVALID')
            result.update(status='PHASE_A_COMPLETE_WITH_35_UNVERIFIED', reason='PHASE_A_ONLY',
                          success=True, exit_code=0, tests_status='NO_EXCEL_TESTS_COMPLETE',
                          junit_sha256=value['junit_sha256'])
    except Stop as error:
        result.update(status='BLOCKED_REVIEW' if error.review else 'BLOCKED', reason=error.code)
    except BaseException:
        # No str(error), traceback, arbitrary exception class or private paths.
        result.update(status='BLOCKED', reason='CONTROLLER_INTERRUPTED_OR_ERROR')
    return result


def main(argv=None):
    try:
        parser = Parser(description=__doc__, add_help=False)
        parser.add_argument('--profile', required=True, type=Path)
        parser.add_argument('--preflight-scratch', required=True, type=Path)
        parser.add_argument('--tests-scratch', required=True, type=Path)
        parser.add_argument('--expected-source-sha', required=True)
        parser.add_argument('--local-unsynced-confirmed', action='store_true')
        args = parser.parse_args(argv)
    except BaseException:
        # Use the same safe output shape even for parser failures.
        result = execute(None, None, None, None, False)
        result['reason'] = 'ARGUMENTS_INVALID'
    else:
        result = execute(args.profile, args.preflight_scratch, args.tests_scratch,
                         args.expected_source_sha, args.local_unsynced_confirmed)
    print(json.dumps(result, ensure_ascii=False))
    return result['exit_code']


if __name__ == '__main__':
    raise SystemExit(main())
