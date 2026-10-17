"""Read-only environment checks plus isolated dummy I/O. Never starts Office/COM."""
import argparse
import ast
import ctypes
import errno
from datetime import datetime, timezone
import hashlib
from importlib.metadata import version
import json
import ntpath
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import time

from . import SCHEMA, A_DIAGNOSTIC, A_PROTOCOL, A_TRANSFORM

REPO = Path(__file__).absolute().parents[2]
BRANCH = 'feature/excel2024-rebuild'
LEDGER = 'lifeplan-excel2024-preopen-approvals'
ORIGINAL = 'output/s1_compatibility/01_function_gate.xlsx'
ORIGINAL_SHA = 'ea158ca2052c1c2ba4d5b7bd30e4304846770e82392bea4229506d973f131e21'
DEPENDENCIES = {'pytest': '8.4.2', 'openpyxl': '3.1.5', 'psutil': '7.0.0', 'pywin32': '311'}
SUPPORTED_IDS = {'Excel2024Retail', 'Excel2024Volume', 'Home2024Retail', 'HomeBusiness2024Retail',
                 'ProPlus2024Volume', 'Standard2024Volume'}
COMPANION_IDS = {'OneNoteFreeRetail', 'LanguagePack', 'ProofingTools'}
SOURCE_EXTRAS = ('src/lifeplan.py', 'src/p2a.py', 'src/p2a_excel.py', 'src/p2a_s1.py',
                 'tools/s1_tiny_diagnostic.py', 'tools/s1_compatibility.py',
                 'tools/p2a_s1_cases.py', 'tools/verify_p2a_s1.py', 'tests/p2a_cases.py')


class Blocked(Exception):
    """Only fixed, non-sensitive reason codes may cross the CLI boundary."""


class SafeParser(argparse.ArgumentParser):
    def error(self, message):
        self.exit(2, 'BLOCKED: ARGUMENTS_INVALID\n')  # never echo private argument values


def require(condition, code):
    if not condition:
        raise Blocked(code)


def path_check(path, *, outside=False):
    require('..' not in Path(path).parts, 'PARENT_TRAVERSAL')
    path = Path(path).absolute()
    require(not str(path).startswith(('\\\\', '//')), 'UNC_NOT_ALLOWED')
    require(LEDGER not in [p.casefold() for p in path.parts], 'APPROVAL_LEDGER_NOT_ALLOWED')
    if outside:
        require(not path.is_relative_to(REPO), 'REPOSITORY_WRITE_NOT_ALLOWED')
    for item in (path, *path.parents):
        info = item.lstat()
        require(not stat.S_ISLNK(info.st_mode) and not getattr(info, 'st_file_attributes', 0) & 0x400,
                'LINK_OR_REPARSE_POINT')
    require(not any(p.casefold().startswith(('onedrive', 'dropbox')) for p in path.parts), 'SYNC_LOCATION')
    for name in ('OneDrive', 'OneDriveConsumer', 'OneDriveCommercial'):
        root = os.environ.get(name)
        if root:
            require(not path.is_relative_to(Path(root).absolute()), 'SYNC_LOCATION')
    return path


def read_file(path, limit=131072):
    path = path_check(path)
    info = path.stat()
    require(stat.S_ISREG(info.st_mode) and info.st_nlink == 1 and info.st_size <= limit, 'UNSAFE_FILE')
    with path.open('rb') as stream:
        raw = stream.read(limit + 1)
    require(len(raw) <= limit, 'FILE_SIZE_LIMIT')
    return raw


def digest(raw):
    return hashlib.sha256(raw).hexdigest()


def source_fingerprint():
    """Runtime digest plus a portable release pin including tests/declarations."""
    paths = sorted((REPO / 'tools/excel2024_validation').glob('*.py'))
    paths += [REPO / p for p in SOURCE_EXTRAS]
    native = {str(p.relative_to(REPO)): digest(read_file(p)) for p in paths}
    portable = {k.replace('\\', '/'): v for k, v in native.items()}
    tests = sorted((REPO / 'tests').glob('test_excel2024_*.py'))
    tests += [REPO / p for p in ('tests/test_p1.py', 'tests/test_p2a.py',
                                 'requirements-dev.txt', 'requirements-excel2024.txt')]
    extra = {str(p.relative_to(REPO)).replace('\\', '/'): digest(read_file(p)) for p in tests}
    return dict(source_sha256=digest(json.dumps(native, sort_keys=True).encode()),
                release_source_sha256=digest(json.dumps({**portable, **extra}, sort_keys=True).encode()),
                files=portable, test_and_dependency_files=extra)


def check_source(expected):
    require(isinstance(expected, str) and re.fullmatch('[0-9a-f]{64}', expected), 'EXPECTED_SOURCE_REQUIRED')
    result = source_fingerprint()
    require(result['release_source_sha256'] == expected, 'SOURCE_SHA_MISMATCH')
    require(digest(read_file(REPO / ORIGINAL)) == ORIGINAL_SHA, 'ORIGINAL_SHA_MISMATCH')
    tree = ast.parse(read_file(REPO / 'tools/excel2024_validation/__init__.py').decode('utf-8'))
    literals = {n.targets[0].id: ast.literal_eval(n.value) for n in tree.body
                if isinstance(n, ast.Assign) and isinstance(n.targets[0], ast.Name)}
    require(all(literals.get(k) == v for k, v in {'A_DIAGNOSTIC': A_DIAGNOSTIC, 'A_PROTOCOL': A_PROTOCOL,
                                                 'A_TRANSFORM': A_TRANSFORM}.items()), 'A_IDS_MISMATCH')
    result['original_sha256'] = ORIGINAL_SHA
    return result


def git_identity():
    if not (REPO / '.git').exists():
        return dict(status='NOT_VERIFIABLE_ZIP', commit=None, branch=None)
    def read(args):
        r = subprocess.run(['git', 'rev-parse', *args], cwd=REPO, stdin=subprocess.DEVNULL,
                           capture_output=True, text=True, encoding='utf-8', timeout=5, check=False)
        require(r.returncode == 0, 'GIT_IDENTITY_UNREADABLE')
        return r.stdout.strip()
    branch, commit = read(['--abbrev-ref', 'HEAD']), read(['HEAD'])
    require(branch == BRANCH and re.fullmatch('[0-9a-f]{40}', commit), 'GIT_BRANCH_OR_COMMIT_MISMATCH')
    return dict(status='READ_ONLY_GIT', branch=branch, commit=commit)


def windows_environment():
    require(sys.platform == 'win32', 'WINDOWS_NOT_AVAILABLE')
    w = sys.getwindowsversion()
    require(w.major == 10 and w.build >= 22000 and w.product_type == 1, 'WINDOWS_11_REQUIRED')
    require(sys.version_info[:2] == (3, 12) and __debug__, 'PYTHON_312_WITH_ASSERTS_REQUIRED')
    found = {n: version(n) for n in DEPENDENCIES}  # metadata only; no COM extension import
    require(found == DEPENDENCIES, 'DEPENDENCY_VERSION_MISMATCH')
    return dict(os='Windows', build=w.build, python='.'.join(map(str, sys.version_info[:3])),
                python_bits=64 if sys.maxsize > 2**32 else 32, dependencies=found,
                native_module_import='NOT_CHECKED')


def local_ntfs(path):
    """Read volume metadata, not a process/Office API. No registry writes."""
    from ctypes import wintypes
    drive = ntpath.splitdrive(str(path))[0] + '\\'
    require(re.fullmatch(r'[A-Za-z]:\\', drive), 'LOCAL_DRIVE_REQUIRED')
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.GetDriveTypeW.argtypes, kernel.GetDriveTypeW.restype = (wintypes.LPCWSTR,), wintypes.UINT
    require(kernel.GetDriveTypeW(drive) == 3, 'FIXED_LOCAL_DRIVE_REQUIRED')
    fs = ctypes.create_unicode_buffer(32)
    kernel.GetVolumeInformationW.argtypes = (wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD,
                                            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
                                            wintypes.LPWSTR, wintypes.DWORD)
    kernel.GetVolumeInformationW.restype = wintypes.BOOL
    require(kernel.GetVolumeInformationW(drive, None, 0, None, None, None, fs, len(fs)), 'VOLUME_UNREADABLE')
    require(fs.value == 'NTFS', 'NTFS_REQUIRED')
    return dict(drive_type='FIXED_LOCAL', filesystem='NTFS')


def read_installation(image):
    """Two read-only HKLM views; return private data only to the comparator."""
    import winreg
    records = []
    for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
                                 0, winreg.KEY_READ | view)
        except FileNotFoundError:
            continue
        with key:
            root = ntpath.normcase(ntpath.normpath(winreg.QueryValueEx(key, 'InstallationPath')[0]))
            require(ntpath.commonpath([root, ntpath.normcase(image)]) == root, 'REGISTRY_PATH_MISMATCH')
            raw = winreg.QueryValueEx(key, 'ProductReleaseIds')[0]
            parts = [raw] if isinstance(raw, str) else raw
            require(isinstance(parts, (list, tuple)) and all(isinstance(p, str) for p in parts), 'PRODUCT_IDS_TYPE')
            known = {p.casefold(): p for p in SUPPORTED_IDS | COMPANION_IDS}
            ids = {known.get(p.casefold(), '') for value in parts for p in re.split(r'[,;\s]+', value.strip()) if p}
            require('' not in ids and bool(ids & SUPPORTED_IDS), 'PRODUCT_IDS_AMBIGUOUS')
            record = dict(installation_path=root, product_release_ids=sorted(ids, key=str.casefold),
                          version_to_report=winreg.QueryValueEx(key, 'VersionToReport')[0],
                          platform=winreg.QueryValueEx(key, 'Platform')[0])
            if record not in records:
                records.append(record)
    require(len(records) == 1, 'REGISTRY_NOT_UNIQUE')
    return records[0]


def image_digest(path):
    path = path_check(path)
    require(stat.S_ISREG(path.stat().st_mode) and path.stat().st_size <= 268435456, 'IMAGE_FILE_LIMIT')
    result, total, started = hashlib.sha256(), 0, time.monotonic()
    with path.open('rb') as stream:
        while chunk := stream.read(1048576):
            result.update(chunk); total += len(chunk)
            require(total <= 268435456 and time.monotonic() - started <= 30, 'IMAGE_READ_BUDGET')
    return result.hexdigest()


def profile_check(path):
    raw = read_file(path)
    p = json.loads(raw.decode('utf-8-sig'))
    require(p.get('schema') == SCHEMA and p.get('licensed') is True
            and p.get('ui_confirmed_on_owned_pid') is True, 'PROFILE_ATTESTATION_MISSING')
    text = re.sub(r'\s+', '', str(p.get('product'))).lower()
    require('2024' in text and any(s in text for s in ('office', 'excel'))
            and not any(s in text for s in ('365', '2010', '2013', '2016', '2019', '2021')), 'PROFILE_PRODUCT')
    facts = p['facts']
    require(facts.get('version') == '16.0' and type(facts.get('pid')) is int and facts['pid'] > 0
            and facts.get('file_version') and facts.get('build'), 'PROFILE_FACTS_INCOMPLETE')
    image = facts['path']
    require(ntpath.isabs(image) and ntpath.basename(image).casefold() == 'excel.exe', 'PROFILE_IMAGE_PATH')
    local_ntfs(Path(image))
    installation = read_installation(image)
    require(installation == facts.get('installation'), 'PROFILE_INSTALLATION_CHANGED')
    require(image_digest(image) == facts.get('exe_sha256'), 'PROFILE_IMAGE_CHANGED')
    require(os.access(Path(path).parent, os.W_OK), 'PROFILE_PARENT_WRITE_ADVISORY_DENIED')
    ledger = Path(path).parent / LEDGER
    ledger_exists = os.path.lexists(ledger)
    if ledger_exists:
        # Inspect only the directory itself/ancestors, never its approval files.
        for item in (ledger, *ledger.parents):
            info = item.lstat()
            require(not stat.S_ISLNK(info.st_mode) and not getattr(info, 'st_file_attributes', 0) & 0x400,
                    'LEDGER_LINK_OR_REPARSE')
        require(ledger.is_dir() and os.access(ledger, os.W_OK), 'LEDGER_ACCESS_ADVISORY_DENIED')
    return dict(profile_sha256=digest(raw), schema_matches=True, installation_matches=True, image_sha_matches=True,
                product_ids=installation['product_release_ids'], saved_ui_attestation=True,
                saved_pid_present=True, next_pid_must_equal_saved_pid=False,
                profile_parent_write_advisory=True, ledger_exists=ledger_exists,
                ledger_write_test='NOT_RUN', current_license='NOT_CHECKED', live_pid='NOT_CHECKED',
                com_version_build='NOT_CHECKED', native_module_loading='NOT_CHECKED')


def scratch_check(path, profile, confirmed):
    require(confirmed, 'OWNER_LOCAL_UNSYNCED_CONFIRMATION_REQUIRED')
    path = path_check(path, outside=True)
    profile = path_check(profile, outside=True)
    require(path.is_dir() and not any(path.iterdir()), 'SCRATCH_MUST_BE_EXISTING_EMPTY')
    require(path != profile.parent and not path.is_relative_to(profile.parent / LEDGER), 'SCRATCH_IS_PROFILE_OR_LEDGER')
    require(len(str(path)) <= 180 and len(str(profile)) <= 240, 'PATH_LENGTH_BUDGET')
    require(os.access(path, os.W_OK), 'SCRATCH_NOT_WRITABLE')
    local_ntfs(path); local_ntfs(profile.parent)
    return path


def dummy_probe(directory):
    # Lazy, safe import: this module has no COM/Session/run dependency.
    from .preopen_input import _move_new, publish
    from .preopen_verify import read_bounded
    source, destination = directory / 'dummy-source.bin', directory / 'dummy-existing.bin'
    for path, content in ((source, b'SOURCE'), (destination, b'KEEP')):
        with path.open('xb') as stream:
            stream.write(content); stream.flush(); os.fsync(stream.fileno())
    try:
        _move_new(source, destination)
    except OSError as error:
        require(error.errno == errno.EEXIST or getattr(error, 'winerror', None) in (80, 183),
                'COLLISION_ACCESS_OR_OTHER_ERROR')
        collision = dict(type=type(error).__name__, errno=error.errno, winerror=getattr(error, 'winerror', None))
        require(read_bounded(source) == b'SOURCE' and read_bounded(destination) == b'KEEP', 'COLLISION_CHANGED_CONTENT')
    else:
        raise Blocked('COLLISION_NOT_REJECTED')
    fresh = directory / 'dummy-new.bin'
    _move_new(source, fresh)
    require(not source.exists() and read_bounded(fresh) == b'SOURCE', 'NEW_MOVE_FAILED')
    published = directory / 'dummy-published.bin'
    publish(published, b'PUBLISH')
    try:
        publish(published, b'NO_REPLACE')
    except FileExistsError:
        require(read_bounded(published) == b'PUBLISH', 'PUBLISH_CHANGED_CONTENT')
    else:
        raise Blocked('PUBLISH_OVERWRITE_NOT_REJECTED')
    require(all(p.stat().st_nlink == 1 for p in (destination, fresh, published)), 'DUMMY_HARDLINK')
    return dict(existing_rejected=True, existing_unchanged=True, new_move_succeeded=True,
                publish_no_clobber=True, independent_regular_files=True,
                collision_error=collision,
                dummy_files=['dummy-existing.bin', 'dummy-new.bin', 'dummy-published.bin'], automatic_cleanup=False)


def save_new(path, value):
    with Path(path).open('x', encoding='utf-8') as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)
        stream.flush(); os.fsync(stream.fileno())


def execute(profile, scratch, expected, confirmed=False):
    value = dict(schema='G2_WINDOWS_NO_EXCEL_PREFLIGHT_V1', status='BLOCKED',
                 created_utc=datetime.now(timezone.utc).isoformat(), checks={},
                 excel_started=False, com_connected=False, approval_id_used=False,
                 gate2_acceptance=False, product_acceptance=False, next_action='STOP_PM_REVIEW')
    safe_scratch = None
    try:
        value['phase'] = 'environment'
        value['checks']['environment'] = windows_environment()
        value['phase'] = 'scratch'
        safe_scratch = scratch_check(scratch, profile, confirmed)
        value['phase'] = 'source'
        value['checks']['source'] = check_source(expected)
        value['checks']['git'] = git_identity()
        value['phase'] = 'profile_readonly'
        value['checks']['profile'] = profile_check(profile)
        value['phase'] = 'dummy_filesystem'
        value['checks']['filesystem'] = local_ntfs(safe_scratch)
        value['checks']['dummy_publication'] = dummy_probe(safe_scratch)
        value['status'] = 'PREFLIGHT_CHECKED'
        value['phase'] = 'complete'
    except BaseException as error:
        value['reason'] = str(error) if type(error) is Blocked else 'CHECK_INTERRUPTED' if isinstance(error, KeyboardInterrupt) else 'CHECK_FAILED'
        value['error_type'] = type(error).__name__
        # Never copy exception text/tracebacks or profile/absolute paths into logs.
    if safe_scratch is not None:
        try:
            save_new(safe_scratch / 'preflight.json', value)
        except BaseException:
            value.update(status='BLOCKED', report_write='FAILED')
    return value


def main():
    parser = SafeParser(description=__doc__)
    parser.add_argument('--profile', required=True, type=Path)
    parser.add_argument('--scratch', required=True, type=Path)
    parser.add_argument('--expected-source-sha', required=True)
    parser.add_argument('--local-unsynced-confirmed', action='store_true')
    args = parser.parse_args()
    result = execute(args.profile, args.scratch, args.expected_source_sha, args.local_unsynced_confirmed)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result['status'] == 'PREFLIGHT_CHECKED' else 2


if __name__ == '__main__':
    raise SystemExit(main())
