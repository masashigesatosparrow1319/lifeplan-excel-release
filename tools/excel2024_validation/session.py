"""Windows-only identity gate. No registry writes, fallback, or process killing."""
import gc
import ntpath
import re
import sys
import time

from .common import GateError, error_info, sha


LIMITATION = ('製品/認証は同じ専用Excel画面での利用者確認とClick-to-Run製品IDに基づく。'
              'COM Version/Build自体は2024のライセンスを証明しない。')

# Exact supported SKUs, not a substring/year or build-number heuristic.
EXCEL_2024_IDS = {s.casefold(): s for s in (
    'Excel2024Retail', 'Excel2024Volume', 'Home2024Retail',
    'HomeBusiness2024Retail', 'ProPlus2024Volume', 'Standard2024Volume',
)}
# These packages do not supply Excel. Unknown companion IDs remain blocked.
NON_EXCEL_IDS = {s.casefold(): s for s in ('OneNoteFreeRetail', 'LanguagePack', 'ProofingTools')}
ID_POLICY = 'excel2024-product-id-policy-v2'
REGISTRY_KEY = r'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'


def parse_product_ids(raw):
    """REG_SZ or REG_MULTI_SZ; trim separators, never infer a product name."""
    parts = [raw] if isinstance(raw, str) else raw if isinstance(raw, (list, tuple)) else None
    if parts is None or not all(isinstance(p, str) for p in parts):
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: ProductReleaseIdsの型が不明です',
                        details=dict(code='INVALID_PRODUCT_IDS_TYPE', detected_value=raw,
                                     detected_type=type(raw).__name__, policy=ID_POLICY))
    known = {**EXCEL_2024_IDS, **NON_EXCEL_IDS}
    ids = {}
    for part in parts:
        for token in re.split(r'[,;\s]+', part.strip()):
            if token:
                ids.setdefault(token.casefold(), known.get(token.casefold(), token))
    return sorted(ids.values(), key=str.casefold)


def product_id_decision(metadata):
    ids = parse_product_ids(metadata.get('product_release_ids', []))
    result = dict(policy=ID_POLICY, detected_product_ids=ids, excel_2024_ids=[],
                  non_excel_companion_ids=[], subscription_ids=[], older_excel_ids=[], unknown_ids=[],
                  basis='Exact supported Excel-2024 SKU + limited non-Excel companions; UI confirmation still required')
    for token in ids:
        name = token.casefold()
        if name in EXCEL_2024_IDS:
            result['excel_2024_ids'].append(token)
        elif name in NON_EXCEL_IDS:
            result['non_excel_companion_ids'].append(token)
        elif name.startswith(('o365', 'm365')) or '365' in name:
            result['subscription_ids'].append(token)
        elif re.fullmatch(r'(excel|home|homebusiness|personal|professional|proplus|standard)(2010|2013|2016|2019|2021)(retail|volume)', name):
            result['older_excel_ids'].append(token)
        else:
            result['unknown_ids'].append(token)
    reasons = []
    if not result['excel_2024_ids']:
        reasons.append('NO_SUPPORTED_EXCEL_2024_ID')
    for key, code in [('subscription_ids', 'SUBSCRIPTION_ID_PRESENT'), ('older_excel_ids', 'OLDER_EXCEL_ID_PRESENT'),
                      ('unknown_ids', 'UNKNOWN_PRODUCT_ID_PRESENT')]:
        if result[key]:
            reasons.append(code)
    result.update(accepted=not reasons, rejection_reasons=reasons)
    return result


def check_installation_ids(metadata, *, progress=None):
    decision = product_id_decision(metadata)
    if progress:
        progress.emit('product_ids.decision', diagnostics=decision)
    if not decision['accepted']:
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: ' + ', '.join(decision['rejection_reasons']), details=decision)
    return decision


def check_product(product, licensed, metadata, *, progress=None):
    decision = product_id_decision(metadata)
    text = re.sub(r'\s+', '', str(product)).lower()
    if (licensed is not True or '2024' not in text or not any(s in text for s in ('excel', 'office'))
            or any(s in text for s in ('365', '2021', '2019', '2016', '2013', '2010'))):
        decision['rejection_reasons'].append('UI_PRODUCT_OR_LICENSE_NOT_CONFIRMED')
        decision['accepted'] = False
    decision.update(ui_product=product, ui_license_attested=licensed is True)
    if progress:
        progress.emit('product_ui.decision', diagnostics=decision)
    if not decision['accepted']:
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: ' + ', '.join(decision['rejection_reasons']), details=decision)
    return decision


def check_ownership(facts, existing_pids, started, expected_path=None):
    if facts['pid'] in existing_pids or facts['created'] < started - 1 or facts['pid'] <= 0:
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: 既存PIDまたは起動時刻不一致。Quit禁止')
    if ntpath.basename(facts['path']).lower() != 'excel.exe':
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: 実行イメージがExcelではありません')
    if expected_path and ntpath.normcase(facts['path']) != ntpath.normcase(expected_path):
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: COM起動先のパス不一致。Quit禁止')


def check_profile(profile, facts, *, progress=None):
    if facts['version'] != '16.0':
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: Excel 2024候補の16.0ではありません',
                        details=dict(code='COM_VERSION_MISMATCH', observed=facts['version'], expected='16.0'))
    check_product(profile.get('product'), profile.get('licensed'), facts['installation'], progress=progress)
    for key in ('path', 'exe_sha256', 'file_version', 'build', 'installation'):
        if facts[key] != profile['facts'].get(key):
            raise GateError('BLOCKED', 'BLOCKED_IDENTITY: 確認時から環境が変化: ' + key + '。doctorを再実行',
                            details=dict(code='PROFILE_ENVIRONMENT_CHANGED', field=key,
                                         expected=profile['facts'].get(key), observed=facts[key]))
    if profile.get('ui_confirmed_on_owned_pid') is not True:
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: 同一専用PIDのUI製品確認がありません')


def installation(path, *, progress=None):
    import winreg
    candidates = []
    observations = []
    for label, view in (('64bit', winreg.KEY_WOW64_64KEY), ('32bit', winreg.KEY_WOW64_32KEY)):
        observation = dict(hive='HKEY_LOCAL_MACHINE', key=REGISTRY_KEY, view=label,
                           process_path=path, values={})
        observations.append(observation)
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REGISTRY_KEY,
                                0, winreg.KEY_READ | view) as key:
                def read(name):
                    observation['reading_value'] = name
                    value, kind = winreg.QueryValueEx(key, name)
                    observation['values'][name] = dict(value=value, registry_type=kind)
                    return value
                root = ntpath.normcase(ntpath.normpath(read('InstallationPath')))
                image = ntpath.normcase(ntpath.normpath(path))
                if ntpath.commonpath([root, image]) != root:
                    observation.update(status='PATH_MISMATCH', installation_path=root)
                    continue
                observation['path_matched'] = True
                ids = parse_product_ids(read('ProductReleaseIds'))
                candidate = dict(installation_path=root, product_release_ids=ids,
                                 version_to_report=read('VersionToReport'), platform=read('Platform'))
                observation.update(status='PATH_MATCHED', candidate=candidate)
                if candidate not in candidates:
                    candidates.append(candidate)
        except (OSError, ValueError, TypeError, GateError) as error:
            observation.update(status='READ_FAILED', error=error_info(error))
    diagnostics = dict(policy=ID_POLICY, process_path=path, registry_views=observations,
                       matched_candidate_count=len(candidates))
    if progress:
        progress.emit('installation.observed', diagnostics=diagnostics)
    # Missing alternate keys are normal. Other unreadable/incomplete views
    # cannot silently establish that no conflicting Excel SKU exists.
    broken_view = any(o['status'] == 'READ_FAILED' and
                      ('reading_value' in o or o['error']['type'] != 'FileNotFoundError') for o in observations)
    if len(candidates) != 1:
        diagnostics['rejection_reasons'] = ['NO_UNIQUE_PATH_MATCHED_INSTALLATION']
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: 実行パスに結びつく一意のClick-to-Run情報が必要', details=diagnostics)
    if broken_view:
        diagnostics['rejection_reasons'] = ['UNREADABLE_OR_INCOMPLETE_REGISTRY_VIEW']
        raise GateError('BLOCKED', 'BLOCKED_IDENTITY: 読取不能/不完全なレジストリviewがあります', details=diagnostics)
    return candidates[0]


class Session:
    """One STA worker, one new COM application, one verified process identity."""
    def __init__(self, progress, profile=None):
        self.log, self.profile = progress, profile
        self.excel = None
        self.owned = False
        self.initialized = False
        self.facts = None
        self.observed_facts = None

    def connect(self):
        if sys.platform != 'win32':
            raise GateError('BLOCKED', 'BLOCKED_EXCEL_UNAVAILABLE: Windows Excel 2024実機なし')
        try:
            import psutil
            import pythoncom
            import win32com.client
            import win32process
            import win32api
        except ImportError as error:
            raise GateError('BLOCKED', 'Windows追加依存をインストールしてください: ' + str(error)) from error
        self.psutil, self.pythoncom = psutil, pythoncom
        self.win32process, self.win32api = win32process, win32api
        # Include all existing PIDs, not only processes whose name is readable.
        self.existing = set(psutil.pids())
        self.started = time.time()
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        self.initialized = True
        self.excel = self.log.call('connect', lambda: win32com.client.DispatchEx('Excel.Application'), 30)
        def identify_owner():
            pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)[1]
            proc = psutil.Process(pid)
            return dict(pid=pid, path=proc.exe(), created=proc.create_time())
        owner = self.log.call('ownership', identify_owner, 30)
        self.owner = owner
        self.log.emit('ownership.observed', identity=owner)
        check_ownership(owner, self.existing, self.started,
                        self.profile['facts']['path'] if self.profile else None)
        self.owned = True
        self.log.emit('ownership.verified', identity=owner)
        facts = self.log.call('identity', self.read_facts, 30)
        self.observed_facts = facts
        self.log.emit('identity.observed', identity=facts)
        if self.profile:
            check_profile(self.profile, facts, progress=self.log)
        else:
            if facts['version'] != '16.0':
                raise GateError('BLOCKED', 'BLOCKED_IDENTITY: 16.0以外のExcel',
                                details=dict(code='COM_VERSION_MISMATCH', observed=facts['version'], expected='16.0'))
            # Check SKU candidates only. Do not invent a confirmed UI product.
            check_installation_ids(facts['installation'], progress=self.log)
        self.facts = facts
        self.log.emit('identity.verified', identity=facts,
                      basis='UI profile + process fingerprint + ClickToRun' if self.profile else 'doctor: UI confirmation pending',
                      limitation=LIMITATION)
        # No mutation/open before process and installation identity checks.
        def settings():
            if self.excel.Workbooks.Count:
                raise GateError('BLOCKED', '専用インスタンスに予期しないブックがあります')
            self.excel.Visible = True
            self.excel.DisplayAlerts = True
            self.excel.AutomationSecurity = 3
        self.log.call('settings', settings)
        return facts

    def read_facts(self):
        self.confirm_owner()
        v = self.win32api.GetFileVersionInfo(self.owner['path'], '\\')
        hi, lo = v['FileVersionMS'], v['FileVersionLS']
        return dict(**self.owner, version=str(self.excel.Version), build=str(self.excel.Build),
                    exe_sha256=sha(self.owner['path']), file_version=f'{hi>>16}.{hi&65535}.{lo>>16}.{lo&65535}',
                    installation=installation(self.owner['path'], progress=self.log))

    def confirm_owner(self):
        proc = self.psutil.Process(self.owner['pid'])
        pid = self.win32process.GetWindowThreadProcessId(self.excel.Hwnd)[1]
        if (pid != self.owner['pid'] or proc.create_time() != self.owner['created']
                or ntpath.normcase(proc.exe()) != ntpath.normcase(self.owner['path'])):
            self.owned = False
            raise GateError('BLOCKED', '専用PID/パス/生成時刻が変わりました。Quit禁止')

    def close(self):
        if self.owned:
            # Read-only observations, not proof that no modal dialog exists.
            self.log.call('cleanup.diagnostics.identity', self.confirm_owner, cleanup=True)
            for name in ('Ready', 'Interactive', 'DisplayAlerts'):
                try:
                    value = self.log.call('cleanup.observe.' + name,
                                          lambda: getattr(self.excel, name), cleanup=True)
                    self.log.emit('cleanup.observed', property=name, value=value)
                except Exception as error:
                    self.log.emit('cleanup.observation_unavailable', property=name, error=error_info(error))
            def quit_owned():
                self.confirm_owner()
                self.excel.Quit()
            self.log.call('cleanup.quit', quit_owned, cleanup=True)

    def release(self):
        self.excel = None
        # pywin32 wrappers can participate in Python reference cycles. Drop
        # application/child wrappers before uninitializing this worker's STA.
        collected = gc.collect()
        try:
            self.log.emit('cleanup.references_collected', collected=collected)
        finally:
            if self.initialized:
                self.pythoncom.CoUninitialize()
                self.initialized = False
