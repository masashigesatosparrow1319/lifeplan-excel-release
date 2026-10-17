"""SYNTHETIC_TEST_ONLY: Windows result shapes, never Windows/Excel/COM."""
import ast
import copy
import hashlib
import json
from pathlib import Path
import stat
from types import SimpleNamespace

import pytest

from tools.excel2024_validation import phase_a_auto as c

PIN = 'a' * 64
PRIVATE = 'PRIVATE_PROFILE_PATH_PID_OR_APPROVAL_VALUE'


@pytest.fixture
def env(tmp_path, monkeypatch):
    profile = tmp_path / 'readonly-profile.json'
    profile.write_text('{"private": "DO_NOT_OUTPUT"}', encoding='utf-8')
    first = tmp_path / 'pre'
    second = tmp_path / 'tests'
    first.mkdir()
    second.mkdir()
    # Only volume metadata is unavailable on Linux. Keep actual path/scratch
    # predicates, symlink detection and exclusive result writing in these tests.
    monkeypatch.setattr(c.pre, 'local_ntfs', lambda path: {'filesystem': 'NTFS'})
    pre_value = dict(schema='G2_WINDOWS_NO_EXCEL_PREFLIGHT_V1', status='PREFLIGHT_CHECKED',
                     checks={'source': {'release_source_sha256': PIN},
                             'environment': {'python_bits': 64}},
                     next_action='STOP_PM_REVIEW', **{key: False for key in c.FLAGS})
    test_value = dict(schema='G2_NO_EXCEL_TESTS_V1', status='NO_EXCEL_TESTS_COMPLETE',
                      tests=list(c.PHASE_TESTS), release_source_sha256=PIN,
                      pytest_exit_code=0, counts=c.COUNTS.copy(), guard_violations=[],
                      junit='tests.xml', junit_sha256='b' * 64,
                      next_action='STOP_PM_REVIEW',
                      **{key: False for key in c.FLAGS if key != 'product_acceptance'})
    calls = []

    def preflight(p, scratch, expected, confirmed):
        calls.append('preflight')
        assert p == profile and scratch == first and expected == PIN and confirmed is True
        c.pre.save_new(scratch / 'preflight.json', pre_value)
        return copy.deepcopy(pre_value)

    def guarded(scratch, expected, confirmed):
        calls.append('guarded13')
        assert calls == ['preflight', 'guarded13']
        assert scratch == second and expected == PIN and confirmed is True
        assert c.tests.TESTS == c.PHASE_TESTS
        c.pre.save_new(scratch / 'tests-summary.json', test_value)
        return copy.deepcopy(test_value)

    monkeypatch.setattr(c.pre, 'execute', preflight)
    monkeypatch.setattr(c.tests, 'execute', guarded)
    return SimpleNamespace(profile=profile, first=first, second=second,
                           pre=pre_value, test=test_value, calls=calls)


def run(e, **kwargs):
    values = dict(profile=e.profile, preflight_scratch=e.first, tests_scratch=e.second,
                  expected=PIN, confirmed=True)
    values.update(kwargs)
    return c.execute(**values)


def test_happy_path_order_exact_13_no_acceptance_and_readonly_profile(env):
    before = env.profile.read_bytes()
    base = c.tests.TESTS
    value = run(env)
    assert value['status'] == 'PHASE_A_COMPLETE_WITH_35_UNVERIFIED'
    assert value['success'] is True and value['exit_code'] == 0
    assert value['preflight_status'] == 'PREFLIGHT_CHECKED'
    assert value['tests_status'] == 'NO_EXCEL_TESTS_COMPLETE'
    assert value['counts'] == dict(passed=947, failed=0, errors=0, skipped=35)
    assert value['test_files'] == 13 and value['unknown'] == 35
    assert value['unverified'] == 'P1_NOT_VERIFIED'
    assert value['next_action'] == 'STOP_PM_REVIEW'
    assert all(value[key] is False for key in c.FLAGS)
    assert value['guard_violation_count'] == 0
    assert value['release_source_sha256'] == PIN and value['junit_sha256'] == 'b' * 64
    assert env.calls == ['preflight', 'guarded13']
    assert c.tests.TESTS is base and len(base) == 11
    assert len(set(c.PHASE_TESTS)) == 13
    assert 'tests/test_excel2024_phase_a_auto.py' not in c.PHASE_TESTS
    assert env.profile.read_bytes() == before
    assert sorted(p.name for p in env.first.iterdir()) == ['preflight.json']
    assert sorted(p.name for p in env.second.iterdir()) == ['tests-summary.json']


@pytest.mark.parametrize('stage', ['pre', 'test'])
@pytest.mark.parametrize('flag', c.FLAGS)
@pytest.mark.parametrize('bad', [True, None, 0, 'false'])
def test_unsafe_flags_rejected(env, stage, flag, bad):
    getattr(env, stage)[flag] = bad
    value = run(env)
    assert value['status'] == 'BLOCKED' and value['exit_code'] == 2
    assert value['reason'] == 'UNSAFE_STAGE_FLAGS'
    if stage == 'pre':
        assert env.calls == ['preflight']
    assert all(value[key] is False for key in c.FLAGS)


@pytest.mark.parametrize('stage', ['pre', 'test'])
@pytest.mark.parametrize('field', ['reason', 'error_type', 'report_write'])
def test_stage_error_or_report_save_failure_stops(env, stage, field):
    getattr(env, stage)[field] = PRIVATE
    value = run(env)
    assert value['reason'] == 'STAGE_ERROR_OR_REPORT_FAILURE'
    assert PRIVATE not in json.dumps(value)
    if stage == 'pre':
        assert env.calls == ['preflight'] and not list(env.second.iterdir())


@pytest.mark.parametrize('stage', ['pre', 'test'])
@pytest.mark.parametrize('key,bad', [('status','BLOCKED'), ('status','PASS'),
                                   ('schema','FAKE'), ('next_action','CONTINUE')])
def test_status_schema_and_next_action_fail_closed(env, stage, key, bad):
    getattr(env, stage)[key] = bad
    value = run(env)
    assert value['exit_code'] == 2 and value['success'] is False
    if stage == 'pre':
        assert env.calls == ['preflight']


@pytest.mark.parametrize('stage', ['pre','test'])
def test_source_mismatch_stops(env,stage):
    if stage == 'pre':
        env.pre['checks']['source']['release_source_sha256'] = 'c' * 64
    else:
        env.test['release_source_sha256'] = 'c' * 64
    value = run(env)
    assert value['reason'] == ('PREFLIGHT_SOURCE_MISMATCH' if stage == 'pre' else 'TEST_SOURCE_MISMATCH')
    assert value['exit_code'] == 2
    if stage == 'pre':
        assert env.calls == ['preflight']


@pytest.mark.parametrize('bits',[32, None, True, '64'])
def test_x64_required(env,bits):
    env.pre['checks']['environment']['python_bits'] = bits
    assert run(env)['reason'] == 'PYTHON_X64_REQUIRED'
    assert env.calls == ['preflight']


@pytest.mark.parametrize('field,bad', [('pytest_exit_code',1),('pytest_exit_code',False),
                                     ('guard_violations',['PRIVATE']),('guard_violations',None),
                                     ('junit','PRIVATE'),('junit_sha256','not-a-hash')])
def test_guard_exit_and_junit_result_checks(env,field,bad):
    env.test[field] = bad
    value = run(env)
    assert value['status'] == 'BLOCKED' and value['exit_code'] == 2
    assert 'PRIVATE' not in json.dumps(value)


@pytest.mark.parametrize('key,bad',[('passed',946),('passed',948),('failed',1),('errors',1),
                                  ('skipped',34),('skipped',36)])
def test_count_drift_requires_review(env,key,bad):
    env.test['counts'][key] = bad
    value = run(env)
    assert value['status'] == 'BLOCKED_REVIEW'
    assert value['reason'] == 'EXACT_COUNTS_REQUIRE_REVIEW'
    assert value['exit_code'] == 2 and value['gate2_acceptance'] is False


@pytest.mark.parametrize('bad',[None,{}, {'passed': True,'failed':0,'errors':0,'skipped':35},
                              dict(passed=947,failed=0,errors=0,skipped=35,extra=1)])
def test_malformed_counts_not_published(env,bad):
    env.test['counts'] = bad
    value = run(env)
    assert value['reason'] == 'COUNTS_INVALID' and value['counts'] is None


@pytest.mark.parametrize('mutation',['extra','reverse','missing','duplicate'])
def test_reported_test_list_change_requires_review(env,mutation):
    names = env.test['tests']
    if mutation == 'extra':names.append('unexpected_test.py')
    elif mutation == 'reverse':names.reverse()
    elif mutation == 'missing':names.pop()
    else:names[-1] = names[0]
    value = run(env)
    assert value['status'] == 'BLOCKED_REVIEW' and value['reason'] == 'TEST_LIST_RESULT_CHANGED'


@pytest.mark.parametrize('mutation',['list','reverse','extra'])
def test_original_tests_definition_cannot_be_modified(env,monkeypatch,mutation):
    names = list(c.BASE_TESTS)
    if mutation == 'reverse':names.reverse()
    if mutation == 'extra':names.append('x.py')
    monkeypatch.setattr(c.tests,'TESTS',names if mutation == 'list' else tuple(names))
    assert run(env)['reason'] == 'TEST_LIST_CHANGED'
    assert env.calls == []


@pytest.mark.parametrize('stage',['pre','test'])
@pytest.mark.parametrize('error',[RuntimeError(PRIVATE),KeyboardInterrupt(PRIVATE),SystemExit(PRIVATE)])
def test_exceptions_redacted_and_test_list_restored(env,monkeypatch,stage,error):
    base = c.tests.TESTS
    def explode(*args,**kwargs):
        print(PRIVATE)
        raise error
    monkeypatch.setattr(c.pre if stage == 'pre' else c.tests,'execute',explode)
    value = run(env)
    assert value['reason'] == 'CONTROLLER_INTERRUPTED_OR_ERROR'
    assert PRIVATE not in json.dumps(value) and value['exit_code'] == 2
    assert c.tests.TESTS is base
    assert value['preflight_status' if stage == 'pre' else 'tests_status'] == 'BLOCKED'


@pytest.mark.parametrize('stage',['pre','test'])
@pytest.mark.parametrize('bad',[None,[],False])
def test_non_dict_stage_results_block(env,monkeypatch,stage,bad):
    monkeypatch.setattr(c.pre if stage == 'pre' else c.tests,'execute',lambda *a,**k:bad)
    assert run(env)['exit_code'] == 2


@pytest.mark.parametrize('case',['same','nested','nonempty_pre','nonempty_tests','missing',
                                'repo','ledger','sync','traversal','profile_directory','profile_hardlink'])
def test_paths_stop_before_preflight_and_preserve_inputs(env,case):
    kwargs = {}
    if case == 'same':kwargs['tests_scratch'] = env.first
    elif case == 'nested':
        nested=env.first/'nested';nested.mkdir();kwargs['tests_scratch']=nested
    elif case.startswith('nonempty'):
        (env.first if case.endswith('pre') else env.second).joinpath('keep').write_bytes(b'KEEP')
    elif case == 'missing':kwargs['tests_scratch'] = env.second/'absent'
    elif case == 'repo':kwargs['preflight_scratch'] = c.pre.REPO
    elif case in ('ledger','sync'):
        path=env.first.parent/(c.pre.LEDGER if case == 'ledger' else 'OneDrive');path.mkdir()
        kwargs['tests_scratch']=path
    elif case == 'traversal':kwargs['tests_scratch'] = env.second/'..'/'tests'
    elif case == 'profile_directory':kwargs['profile'] = env.first.parent
    else:
        import os
        os.link(env.profile,env.profile.parent/'linked-profile')
    value = run(env,**kwargs)
    assert value['exit_code'] == 2 and env.calls == []
    assert env.profile.read_bytes() == b'{"private": "DO_NOT_OUTPUT"}'
    assert not (env.first/'preflight.json').exists()
    assert not (env.second/'tests-summary.json').exists()


@pytest.mark.parametrize('which',['profile','pre','tests'])
@pytest.mark.parametrize('kind',['symlink','reparse'])
def test_artificial_link_reparse_rejected_without_privileges(env,monkeypatch,which,kind):
    # Existing path validation remains real; only this fixture's lstat attributes
    # are artificial. This test does not create symlinks on Windows.
    target={'profile':env.profile,'pre':env.first,'tests':env.second}[which]
    real=Path.lstat
    def attributes(path,*a,**k):
        s=real(path,*a,**k)
        if path == target:
            return SimpleNamespace(st_mode=stat.S_IFLNK if kind=='symlink' else s.st_mode,
                                   st_file_attributes=0x400 if kind=='reparse' else 0)
        return s
    monkeypatch.setattr(Path,'lstat',attributes)
    value=run(env)
    assert value['exit_code'] == 2 and env.calls == []
    assert not (env.first/'preflight.json').exists()


@pytest.mark.parametrize('confirmed',[False,None,0,'yes'])
def test_owner_flag_required(env,confirmed):
    assert run(env,confirmed=confirmed)['reason'] == 'OWNER_CONFIRMATION_REQUIRED'
    assert env.calls == []


@pytest.mark.parametrize('pin',[None,'',PRIVATE,'A'*64,'a'*63])
def test_invalid_source_pin_not_echoed(env,pin):
    value=run(env,expected=pin)
    assert value['reason'] == 'SOURCE_PIN_REQUIRED' and env.calls == []
    assert PRIVATE not in json.dumps(value)


def test_cli_single_safe_json_on_success(env,capsys):
    code=c.main(['--profile',str(env.profile),'--preflight-scratch',str(env.first),
                 '--tests-scratch',str(env.second),'--expected-source-sha',PIN,'--local-unsynced-confirmed'])
    output=capsys.readouterr()
    value=json.loads(output.out)
    assert code == 0 and value['counts']['passed'] == 947
    assert len(output.out.splitlines()) == 1 and output.err == ''
    assert str(env.profile) not in output.out


@pytest.mark.parametrize('args',[[],['--profile',PRIVATE],['--unexpected',PRIVATE],['--help']])
def test_cli_parser_redacts_private_arguments(env,capsys,args):
    assert c.main(args) == 2
    captured=capsys.readouterr()
    value=json.loads(captured.out)
    assert value['reason'] == 'ARGUMENTS_INVALID' and value['success'] is False
    assert PRIVATE not in captured.out and captured.err == '' and env.calls == []


def test_cli_error_never_prints_exception_or_paths(env,monkeypatch,capsys):
    def explode(*a,**k):
        print(PRIVATE)
        raise RuntimeError(str(env.profile))
    monkeypatch.setattr(c.pre,'execute',explode)
    args=['--profile',str(env.profile),'--preflight-scratch',str(env.first),
          '--tests-scratch',str(env.second),'--expected-source-sha',PIN,'--local-unsynced-confirmed']
    assert c.main(args) == 2
    output=capsys.readouterr()
    assert json.loads(output.out)['reason'] == 'CONTROLLER_INTERRUPTED_OR_ERROR'
    assert PRIVATE not in output.out and str(env.profile) not in output.out and output.err == ''


def test_existing_reports_never_overwritten_or_retried(env):
    assert run(env)['exit_code'] == 0
    before={p:hashlib.sha256(p.read_bytes()).hexdigest() for p in (env.first/'preflight.json',env.second/'tests-summary.json')}
    assert run(env)['exit_code'] == 2
    assert env.calls == ['preflight','guarded13']
    assert all(hashlib.sha256(p.read_bytes()).hexdigest() == h for p,h in before.items())


def test_controller_has_only_two_approved_execution_calls_and_no_launcher():
    tree=ast.parse(Path(c.__file__).read_text(encoding='utf-8'))
    attributes={ast.unparse(n.func) for n in ast.walk(tree) if isinstance(n,ast.Call)}
    assert 'pre.execute' in attributes and 'tests.execute' in attributes
    assert not any(n in attributes for n in ('subprocess.run','subprocess.Popen','os.system',
                   'os.startfile','sys.addaudithook','run.execute','audit.execute','Session'))
    imports=[ast.unparse(n) for n in ast.walk(tree) if isinstance(n,(ast.Import,ast.ImportFrom))]
    assert not any(x in ' '.join(imports) for x in ('win32com','pythoncom','psutil','offline_tiny_e1','session','worker'))


def test_explicit_false_product_in_legacy_tests_is_allowed(env):
    env.test['product_acceptance']=False
    assert run(env)['exit_code'] == 0
