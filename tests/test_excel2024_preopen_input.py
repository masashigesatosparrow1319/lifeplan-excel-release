"""A-only bounded ZIP/proof tests. All generated XLSX live in pytest tmp paths."""
import copy
import io
from pathlib import Path
import stat
import struct
from types import SimpleNamespace
import zipfile

import pytest

from tools.excel2024_validation import A_DIAGNOSTIC, A_PROTOCOL, A_TRANSFORM
from tools.excel2024_validation import common, preopen_input as builder, preopen_verify as checker

APPROVAL = 'G2-20261016-999-WIN-01'  # mock ledger identifier; no Windows authorization


def prepared(tmp_path, *, source='MOCK_SOURCE'):
    directory, profiles = tmp_path / 'case', tmp_path / 'profiles'
    directory.mkdir(); profiles.mkdir()
    return directory, builder.prepare(directory, 'test', APPROVAL, profiles, source)


def repack(raw, change):
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(out, 'w') as dst:
        for info in src.infolist():
            info = copy.copy(info)
            content = src.read(info.filename)
            changed = change(info, content)
            if changed is not None:
                dst.writestr(*changed)
    return out.getvalue()


def test_exact_one_byte_independent_part_proof_and_no_cache(tmp_path, request):
    before = builder.ORIGINAL.read_bytes()
    directory, manifest = prepared(tmp_path)
    original = directory / 'provenance/original' / checker.FILE
    contrast = directory / 'inputs' / checker.FILE
    proof = checker.verify(original, contrast)
    assert builder.ORIGINAL.read_bytes() == before == original.read_bytes()
    assert proof['original_sha256'] == checker.ORIGINAL_SHA
    assert proof['input_sha256'] == manifest['sha256'] != checker.ORIGINAL_SHA
    assert proof['formula_count'] == 24 and proof['nonempty_caches'] == 0
    assert proof['preserved_attributes'] == dict(calcId='0', calcMode='auto', fullCalcOnLoad='1')
    a, _ = checker.package(before)
    b, _ = checker.package(contrast.read_bytes())
    assert [name for name in checker.PARTS if a[name] != b[name]] == [checker.WORKBOOK]
    assert sum(x != y for x, y in zip(a[checker.WORKBOOK], b[checker.WORKBOOK])) == 1
    assert len(a[checker.WORKBOOK]) == len(b[checker.WORKBOOK]) == 705
    assert sum(p['changed'] for p in proof['parts']) == 1 and len(proof['parts']) == 11
    assert proof['change']['before'] == '1' and proof['change']['after'] == '0'
    control = builder.validate_bundle(directory, manifest)
    assert control['verified'] is True
    assert contrast.stat().st_nlink == (directory / 'work' / checker.FILE).stat().st_nlink == 1
    assert contrast.read_bytes() == (directory / 'work' / checker.FILE).read_bytes()
    for key in ('original_sha256', 'input_sha256', 'workbook_before_sha256', 'workbook_after_sha256'):
        request.node.user_properties.append((key, proof[key]))


@pytest.mark.parametrize('fault', ['attribute', 'full_load', 'calc_mode', 'calc_id', 'namespace', 'duplicate_attr',
                                 'cell', 'formula', 'cache', 'meta', 'relationship', 'unknown_xml'])
def test_independent_checker_refuses_any_other_payload_change(tmp_path, fault):
    directory, manifest = prepared(tmp_path)
    path = directory / 'inputs' / manifest['file']
    def change(info, content):
        if info.filename == checker.WORKBOOK:
            substitutions = {'attribute': (b'forceFullCalc="0"', b'forceFullCalc="1"'),
                             'full_load': (b'fullCalcOnLoad="1"', b'fullCalcOnLoad="0"'),
                             'calc_mode': (b'calcMode="auto"', b'calcMode="manual"'),
                             'calc_id': (b'calcId="0"', b'calcId="1"'),
                             'namespace': (b'/2006/main', b'/2007/main'),
                             'duplicate_attr': (b'forceFullCalc="0"', b'forceFullCalc="0" forceFullCalc="0"')}
            if fault in substitutions:
                left, right = substitutions[fault]; content = content.replace(left, right)
        if info.filename == 'xl/worksheets/sheet1.xml' and fault in ('cell', 'formula', 'cache', 'unknown_xml'):
            if fault == 'cell': content = content.replace(b'>100<', b'>101<', 1)
            if fault == 'formula': content = content.replace(b'<f>', b'<f>1+', 1)
            if fault == 'cache': content = content.replace(b'<v />', b'<v>999</v>', 1)
            if fault == 'unknown_xml': content = b'<!DOCTYPE x [<!ENTITY bad "x">]>' + content
        if fault == 'meta' and info.filename == 'xl/worksheets/sheet3.xml': content += b' '
        if fault == 'relationship' and info.filename == 'xl/_rels/workbook.xml.rels': content += b' '
        return info, content
    altered = repack(path.read_bytes(), change)
    assert altered != path.read_bytes(), 'malicious fixture must actually change'
    path.write_bytes(altered)
    with pytest.raises((ValueError, zipfile.BadZipFile)):
        builder.validate_bundle(directory, manifest)


@pytest.mark.parametrize('fault', ['traversal', 'duplicate', 'missing', 'symlink', 'compression', 'extra', 'comment',
                                 'encrypted', 'crc', 'oversize', 'trailing', 'local_timestamp', 'local_name'])
def test_closed_zip_layout_header_crc_and_resources(tmp_path, fault):
    raw = builder.convert(builder.ORIGINAL.read_bytes())
    if fault in ('traversal', 'duplicate', 'missing', 'symlink', 'compression', 'extra', 'comment'):
        index = [0]
        def change(info, content):
            index[0] += 1
            if index[0] == 1:
                if fault == 'traversal': info.filename = '../app.xml'
                if fault == 'duplicate': info.filename = checker.PARTS[1]
                if fault == 'missing': return None
                if fault == 'symlink': info.external_attr = 0o120777 << 16
                if fault == 'compression': info.compress_type = zipfile.ZIP_STORED
                if fault == 'extra': info.extra = b'\x01\x00\x00\x00'
                if fault == 'comment': info.comment = b'unknown'
            return info, content
        raw = repack(raw, change)
    else:
        raw = bytearray(raw)
        central = raw.index(b'PK\x01\x02')
        if fault == 'encrypted':
            struct.pack_into('<H', raw, 6, 1); struct.pack_into('<H', raw, central + 8, 1)
        if fault == 'crc':
            struct.pack_into('<L', raw, 14, 1); struct.pack_into('<L', raw, central + 16, 1)
        if fault == 'oversize': struct.pack_into('<L', raw, central + 24, checker.MAX_EXPANDED + 1)
        if fault == 'trailing': raw += b'x'
        if fault == 'local_timestamp': raw[10] ^= 1
        if fault == 'local_name': raw[30] ^= 1
    with pytest.raises((ValueError, zipfile.BadZipFile, RuntimeError)):
        checker.package(bytes(raw))


def test_builder_and_checker_are_independent_and_original_is_pinned(tmp_path, monkeypatch):
    raw = builder.ORIGINAL.read_bytes()
    original, contrast = tmp_path / 'original.xlsx', tmp_path / 'contrast.xlsx'
    original.write_bytes(raw); contrast.write_bytes(builder.convert(raw))
    monkeypatch.setattr(builder, 'convert', lambda *a: pytest.fail('verifier must not call converter'))
    assert checker.verify(original, contrast)['formula_count'] == 24
    original.write_bytes(b'wrong original')
    with pytest.raises(ValueError, match='原本SHA'): checker.verify(original, contrast)


@pytest.mark.parametrize('fault', ['proof_missing', 'proof_sha', 'proof_forged', 'proof_source', 'timestamp',
                                 'case_sha', 'control_schema', 'variant', 'original', 'work', 'approval', 'extra_key'])
def test_proof_is_bound_to_rechecked_bytes_and_closed_ids(tmp_path, fault):
    directory, manifest = prepared(tmp_path)
    proof_path = directory / 'transform.json'
    proof = common.read_json(proof_path)
    if fault == 'proof_missing': proof_path.unlink()
    if fault == 'proof_sha': proof_path.write_bytes(b'{}')
    if fault in ('proof_forged', 'proof_source', 'timestamp'):
        if fault == 'proof_forged': proof['verification']['formula_count'] = 23
        if fault == 'proof_source': proof['checker_sha256'] = 'wrong'
        if fault == 'timestamp': proof['created_utc'] = 'no-time'
        common.write_json(proof_path, proof)
        manifest['control']['transform_sha256'] = common.sha(proof_path)
    if fault == 'case_sha': manifest['sha256'] = checker.ORIGINAL_SHA
    if fault == 'control_schema': manifest['control']['schema'] = 'other'
    if fault == 'variant': manifest['control']['variant'] = 'T0'
    if fault == 'original': manifest['control']['original_sha256'] = 'other'
    if fault == 'work': (directory / 'work' / checker.FILE).write_bytes(builder.ORIGINAL.read_bytes())
    if fault == 'approval': manifest['control']['approval_sha256'] = 'wrong'
    if fault == 'extra_key': manifest['control']['allow_any_input'] = True
    with pytest.raises((ValueError, FileNotFoundError)):
        builder.validate_bundle(directory, manifest)


@pytest.mark.parametrize('fault', ['run', 'diagnostic', 'protocol', 'source', 'manifest', 'control', 'case_id'])
def test_request_binding_cannot_be_reused_or_downgraded(tmp_path, fault):
    directory, manifest = prepared(tmp_path)
    req = dict(run_id='test', case_id='functions24', diagnostic=A_DIAGNOSTIC, protocol=A_PROTOCOL,
               source_sha256='MOCK_SOURCE', manifest_sha256=common.sha(directory / 'case.json'),
               control=copy.deepcopy(manifest['control']))
    if fault == 'run': req['run_id'] = 'other'
    if fault == 'diagnostic': req['diagnostic'] = 'FUNCTIONS24_FORCE_FALSE_V1'
    if fault == 'protocol': req['protocol'] = 'other'
    if fault == 'source': req['source_sha256'] = 'other'
    if fault == 'manifest': req['manifest_sha256'] = 'other'
    if fault == 'control': req['control']['variant'] = 'other'
    if fault == 'case_id': req['case_id'] = 'annual1'
    with pytest.raises(ValueError): builder.validate_bundle(directory, manifest, req)


def test_exclusive_publication_and_approval_one_use_across_runs(tmp_path):
    target = tmp_path / 'proof.json'
    builder.publish(target, b'first')
    with pytest.raises(FileExistsError): builder.publish(target, b'second')
    assert target.read_bytes() == b'first'
    directory, _ = prepared(tmp_path)
    second = tmp_path / 'case2'; second.mkdir()
    with pytest.raises(FileExistsError): builder.prepare(second, 'test2', APPROVAL, tmp_path / 'profiles', 'MOCK_SOURCE')
    assert not list(second.iterdir())
    with pytest.raises(FileExistsError): builder.prepare(directory, 'test', APPROVAL, tmp_path / 'profiles', 'MOCK_SOURCE')


def test_oversized_publication_refused_before_creating_any_file(tmp_path):
    with pytest.raises(ValueError): builder.publish(tmp_path / 'large', b'x' * (checker.MAX_ARCHIVE + 1))
    assert not list(tmp_path.iterdir())


def test_nul_encoded_xml_cannot_bypass_dtd_entity_filter():
    raw = builder.convert(builder.ORIGINAL.read_bytes())
    def change(info, content):
        if info.filename == 'docProps/app.xml':
            content = '<!DOCTYPE x [<!ENTITY e "x">]><x>&e;</x>'.encode('utf-16')
        return info, content
    with pytest.raises(ValueError, match='DTD/entity'): checker.package(repack(raw, change))


def test_failed_atomic_publication_keeps_staging_and_does_not_clobber(tmp_path, monkeypatch):
    target = tmp_path / 'proof.json'
    def collision(source, dest):
        dest.write_bytes(b'competing owner')
        raise FileExistsError(dest)
    monkeypatch.setattr(builder, '_move_new', collision)
    with pytest.raises(FileExistsError): builder.publish(target, b'candidate')
    assert target.read_bytes() == b'competing owner'
    assert [p.read_bytes() for p in tmp_path.glob('.staging-*')] == [b'candidate']


def test_native_linux_no_replace_refuses_existing_target(tmp_path):
    a, b = tmp_path / 'a', tmp_path / 'b'
    a.write_bytes(b'A'); b.write_bytes(b'B')
    with pytest.raises(FileExistsError): builder._move_new(a, b)
    assert a.read_bytes() == b'A' and b.read_bytes() == b'B'


def test_parent_traversal_rejected_even_when_lexically_outside_repo(tmp_path):
    with pytest.raises(ValueError): builder.safe_path(tmp_path / '..' / 'other', existing=False)


def test_builder_refuses_unpinned_original_without_transformation():
    with pytest.raises(ValueError): builder.convert(b'not the adopted original')


def test_windows_publication_uses_no_replace_flag_mock_only(tmp_path, monkeypatch):
    calls = []
    class Move:
        def __call__(self, source, target, flags): calls.append((source, target, flags)); return True
    move = Move()
    monkeypatch.setattr(builder, 'sys', SimpleNamespace(platform='win32'))
    monkeypatch.setattr(builder.ctypes, 'WinDLL', lambda *a, **k: SimpleNamespace(MoveFileExW=move), raising=False)
    builder._move_new(tmp_path / 'a', tmp_path / 'b')
    assert calls == [(str(tmp_path / 'a'), str(tmp_path / 'b'), 8)]  # not REPLACE_EXISTING=1


@pytest.mark.parametrize('fault', ['link', 'oversize', 'repo_path', 'missing_parent'])
def test_path_guards(tmp_path, monkeypatch, fault):
    path = tmp_path / 'file'; path.write_bytes(b'x')
    if fault == 'link':
        # Do not require Windows symlink privileges when preflight pytest is reused.
        original_lstat = Path.lstat
        monkeypatch.setattr(Path, 'lstat', lambda p: SimpleNamespace(st_mode=stat.S_IFLNK | 0o777)
                            if p == path else original_lstat(p))
    if fault == 'oversize': path.write_bytes(b'x' * (checker.MAX_ARCHIVE + 1))
    if fault == 'repo_path':
        with pytest.raises(ValueError): builder.safe_path(builder.ORIGINAL)
        return
    if fault == 'missing_parent': path = tmp_path / 'missing' / 'file'
    with pytest.raises((ValueError, FileNotFoundError)): checker.read_bounded(path)


def test_resource_limits_and_unsealed_failure_do_not_publish_input(tmp_path, monkeypatch):
    directory, profiles = tmp_path / 'case', tmp_path / 'profiles'
    directory.mkdir(); profiles.mkdir()
    monkeypatch.setattr(checker, 'verify', lambda *a: (_ for _ in ()).throw(ValueError('independent rejection')))
    with pytest.raises(ValueError, match='independent rejection'):
        builder.prepare(directory, 'test', APPROVAL, profiles, 'MOCK_SOURCE')
    assert not list((directory / 'inputs').iterdir())
    assert not (directory / 'case.json').exists()
    assert (profiles / builder.LEDGER / (APPROVAL + '.json')).exists()


def test_zip_inspection_has_finite_time_budget(monkeypatch):
    times = iter([0, 11])
    monkeypatch.setattr(checker.time, 'monotonic', lambda: next(times))
    with pytest.raises(common.GateError) as caught: checker.package(builder.ORIGINAL.read_bytes())
    assert caught.value.status == 'TIMEOUT'


@pytest.mark.parametrize('value', [None, '', 'other', '../WIN-01', 'G2-20261016-006'])
def test_windows_approval_is_required_not_implicitly_granted(value):
    with pytest.raises(common.GateError) as caught: builder.validate_approval_id(value)
    assert caught.value.status == 'BLOCKED'
