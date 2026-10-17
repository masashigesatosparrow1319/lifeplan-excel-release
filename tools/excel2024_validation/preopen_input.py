"""A-only immutable preparation and proof binding. No Excel/COM/cache writes."""
import copy
import ctypes
from datetime import datetime, timedelta
import io
import json
import os
from pathlib import Path
import re
import stat
import sys
import time
import uuid
import zipfile

from . import A_DIAGNOSTIC, A_PROTOCOL, A_TRANSFORM
from .cases import REPO, EXPECTED24, static_input
from .common import GateError, read_json, sha, utc
from . import preopen_verify as checker

VARIANT = 'F24_PREOPEN_FORCE0'
LEDGER = 'lifeplan-excel2024-preopen-approvals'
ORIGINAL = REPO / 'output/s1_compatibility' / checker.FILE


def safe_path(path, *, existing=True):
    if '..' in Path(path).parts:
        raise ValueError('親ディレクトリへのtraversalは禁止')
    path = Path(path).absolute()
    if path.is_relative_to(REPO) or str(path).startswith('\\\\'):
        raise ValueError('A証拠はリポジトリ外のローカル新規runだけです')
    for item in (path, *path.parents):
        try:
            info = item.lstat()
        except FileNotFoundError:
            if item == path and not existing:
                continue
            raise
        if stat.S_ISLNK(info.st_mode) or getattr(info, 'st_file_attributes', 0) & 0x400:
            raise ValueError('リンク/reparse pointは使用禁止')
    return path


def _move_new(source, target):
    """Native atomic no-replace; deliberately no hardlink/replace fallback."""
    if sys.platform == 'win32':
        from ctypes import wintypes
        library = ctypes.WinDLL('kernel32', use_last_error=True)
        move = library.MoveFileExW
        move.argtypes = (wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD)
        move.restype = wintypes.BOOL
        if not move(str(source), str(target), 8):  # WRITE_THROUGH, no REPLACE_EXISTING
            raise ctypes.WinError(ctypes.get_last_error())
    elif sys.platform.startswith('linux'):
        library = ctypes.CDLL(None, use_errno=True)
        move = getattr(library, 'renameat2', None)
        if move is None:
            raise GateError('BLOCKED', 'atomic no-replaceのrenameat2がありません')
        move.argtypes = (ctypes.c_int, ctypes.c_char_p, ctypes.c_int, ctypes.c_char_p, ctypes.c_uint)
        move.restype = ctypes.c_int
        if move(-100, os.fsencode(source), -100, os.fsencode(target), 1):  # RENAME_NOREPLACE
            code = ctypes.get_errno()
            raise OSError(code, os.strerror(code), str(target))
    else:
        raise GateError('BLOCKED', 'このFS/OSのatomic no-replaceは未対応')


def publish(path, content):
    if len(content) > checker.MAX_ARCHIVE:
        raise ValueError('公開ファイルのサイズ上限超過')
    path = safe_path(path, existing=False)
    if path.exists():
        raise FileExistsError(path)
    temporary = path.parent / ('.staging-' + uuid.uuid4().hex)
    with temporary.open('xb') as stream:
        stream.write(content)
        stream.flush()
        os.fsync(stream.fileno())
    # Keep failed staging files as evidence. Never remove/reuse an old run.
    _move_new(temporary, path)
    if checker.read_bounded(path) != content:
        raise ValueError('公開後のbyte不一致')


def publish_json(path, value):
    publish(path, json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2).encode('utf-8'))


def validate_approval_id(value):
    if not isinstance(value, str) or re.fullmatch(r'G2-\d{8}-\d{3}-WIN-\d{2}', value) is None:
        raise GateError('BLOCKED', 'A専用Windows実行には別承認の実行ID（G2-YYYYMMDD-NNN-WIN-NN）が必要です')


def convert(raw):
    """Builder uses a pinned lexical replacement; verifier never calls this."""
    if len(raw) > checker.MAX_ARCHIVE or checker.digest(raw) != checker.ORIGINAL_SHA:
        raise ValueError('変換元は固定原本SHAだけです')
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw)) as source, zipfile.ZipFile(output, 'w') as target:
        for info in source.infolist():
            content = source.read(info)
            if info.filename == checker.WORKBOOK:
                token = b'forceFullCalc="1"'
                if checker.digest(content) != checker.WORKBOOK_SHA or content.count(token) != 1:
                    raise ValueError('変換元属性の固定前提不一致')
                position = content.index(token) + len(b'forceFullCalc="')
                content = content[:position] + b'0' + content[position+1:]
            target.writestr(copy.copy(info), content)
    return output.getvalue()


def prepare(directory, run_id, approval_id, profile_directory, source_sha256):
    validate_approval_id(approval_id)
    directory = safe_path(directory)
    if any(directory.iterdir()):
        raise FileExistsError('A prepareは空の新規caseディレクトリだけです')
    profile_directory = safe_path(profile_directory)
    ledger = profile_directory / LEDGER
    ledger.mkdir(exist_ok=True)
    safe_path(ledger)
    marker = dict(schema=A_TRANSFORM, diagnostic=A_DIAGNOSTIC, run_id=run_id,
                  approval_id=approval_id, source_sha256=source_sha256)
    approval_path = ledger / (approval_id + '.json')
    publish_json(approval_path, marker)  # spent even on later failure; no auto reset
    started = time.monotonic()
    original_raw = checker.read_bounded(ORIGINAL)
    if checker.digest(original_raw) != checker.ORIGINAL_SHA:
        raise ValueError('原本SHA不一致')
    checker.package(original_raw)
    record = next(e for e in read_json(ORIGINAL.parent / 'manifest.json') if e['file'] == checker.FILE)
    if record['sha256'] != checker.ORIGINAL_SHA or record['formula_cells'] != 24:
        raise ValueError('既存原本manifest不一致')
    for name in ('provenance', 'provenance/original', 'inputs', 'work', 'results'):
        (directory / name).mkdir()
    original_copy = directory / 'provenance/original' / checker.FILE
    publish(original_copy, original_raw)
    candidate = convert(original_raw)
    # Verify BEFORE publishing eligible input, using a separate staging file.
    staging = directory / 'provenance' / 'candidate.xlsx'
    publish(staging, candidate)
    verified = checker.verify(original_copy, staging)
    publish(directory / 'inputs' / checker.FILE, candidate)
    publish(directory / 'work' / checker.FILE, candidate)
    if static_input(directory / 'inputs' / checker.FILE) != 24:
        raise ValueError('既存24式静的検査不一致')
    proof = dict(schema=A_TRANSFORM, diagnostic=A_DIAGNOSTIC, protocol=A_PROTOCOL, variant=VARIANT,
                 run_id=run_id, approval_id=approval_id, created_utc=utc(), verification=verified,
                 builder_sha256=sha(__file__), checker_sha256=sha(checker.__file__))
    publish_json(directory / 'transform.json', proof)
    control = dict(schema=A_TRANSFORM, diagnostic=A_DIAGNOSTIC, protocol=A_PROTOCOL, variant=VARIANT,
                   original_sha256=checker.ORIGINAL_SHA, input_sha256=verified['input_sha256'],
                   transform_sha256=sha(directory / 'transform.json'), approval_id=approval_id,
                   approval_path=str(approval_path), approval_sha256=sha(approval_path))
    manifest = dict(run_id=run_id, case_id='functions24', file=checker.FILE,
                    sha256=verified['input_sha256'], formula_count=24,
                    probes=[['確認', c, v] for c, v in EXPECTED24.items()],
                    edits=[['確認', 'B8', 100]], challenges=[['確認', 'B8', 0]], control=control)
    if time.monotonic() - started > 30:
        raise GateError('TIMEOUT', 'A入力前処理30秒超過')
    if checker.digest(checker.read_bounded(ORIGINAL)) != checker.ORIGINAL_SHA:
        raise ValueError('準備中の原本変更')
    publish_json(directory / 'case.json', manifest)
    return manifest


def validate_bundle(directory, manifest, request=None):
    from .functions24 import validate_manifest
    validate_manifest(manifest, diagnostic=A_DIAGNOSTIC)
    directory = safe_path(directory)
    control = manifest.get('control', {})
    expected_header = dict(schema=A_TRANSFORM, diagnostic=A_DIAGNOSTIC, protocol=A_PROTOCOL, variant=VARIANT)
    if (set(control) != set(expected_header) | {'original_sha256', 'input_sha256', 'transform_sha256',
                                             'approval_id', 'approval_path', 'approval_sha256'}
            or any(control.get(k) != v for k, v in expected_header.items())):
        raise ValueError('A controlのID/schema/variant不一致')
    validate_approval_id(control.get('approval_id'))
    proof_path = directory / 'transform.json'
    proof_raw = checker.read_bounded(proof_path)
    if checker.digest(proof_raw) != control.get('transform_sha256'):
        raise ValueError('transform SHA不一致')
    proof = json.loads(proof_raw.decode('utf-8'))
    expected_keys = set(expected_header) | {'run_id', 'approval_id', 'created_utc', 'verification', 'builder_sha256', 'checker_sha256'}
    if set(proof) != expected_keys or any(proof.get(k) != v for k, v in expected_header.items()):
        raise ValueError('A proofのschema/fields不一致')
    if (proof['run_id'] != manifest['run_id'] or proof['approval_id'] != control['approval_id']
            or proof['builder_sha256'] != sha(__file__) or proof['checker_sha256'] != sha(checker.__file__)
            or not isinstance(proof['created_utc'], str) or not proof['created_utc']):
        raise ValueError('A proofのrun/source識別不一致')
    if datetime.fromisoformat(proof['created_utc']).utcoffset() != timedelta(0):
        raise ValueError('A proofにはtimezone付きUTC時刻が必要です')
    original = directory / 'provenance/original' / checker.FILE
    inputs, work = (directory / name / checker.FILE for name in ('inputs', 'work'))
    evidence = checker.verify(original, inputs)
    if evidence != proof['verification']:
        raise ValueError('独立再検査とproof申告が不一致')
    if (manifest['sha256'] != evidence['input_sha256'] or control.get('input_sha256') != manifest['sha256']
            or control.get('original_sha256') != checker.ORIGINAL_SHA
            or checker.digest(checker.read_bounded(work)) != manifest['sha256']):
        raise ValueError('A元/入力/操作SHA不一致')
    if any(os.path.samefile(a, b) for a, b in ((original, inputs), (inputs, work), (original, work))):
        raise ValueError('原本/入力/操作は独立ファイルが必要です')
    approval_path = safe_path(control['approval_path'])
    approval_raw = checker.read_bounded(approval_path)
    if (approval_path.parent.name != LEDGER or approval_path.name != control['approval_id'] + '.json'
            or checker.digest(approval_raw) != control.get('approval_sha256')):
        raise ValueError('一回性承認印の識別不一致')
    approval = json.loads(approval_raw.decode('utf-8'))
    if (set(approval) != {'schema', 'diagnostic', 'run_id', 'approval_id', 'source_sha256'}
            or approval.get('schema') != A_TRANSFORM or approval.get('diagnostic') != A_DIAGNOSTIC
            or approval.get('run_id') != manifest['run_id'] or approval.get('approval_id') != control['approval_id']):
        raise ValueError('一回性承認印のrun不一致')
    if request is not None:
        if (request.get('diagnostic') != A_DIAGNOSTIC or request.get('protocol') != A_PROTOCOL
                or request.get('case_id') != 'functions24' or request.get('run_id') != manifest['run_id']
                or request.get('control') != control or request.get('manifest_sha256') != sha(directory / 'case.json')
                or request.get('source_sha256') != approval.get('source_sha256')):
            raise ValueError('A request/case/source/controlの束縛不一致')
    if static_input(inputs) != 24 or checker.digest(checker.read_bounded(ORIGINAL)) != checker.ORIGINAL_SHA:
        raise ValueError('A原本/静的検査不一致')
    return dict(transform_sha256=control['transform_sha256'], original_sha256=checker.ORIGINAL_SHA,
                input_sha256=manifest['sha256'], approval_id=control['approval_id'], verified=True)
