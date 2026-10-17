"""Durable records and cooperative cancellation; never terminates processes."""
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import time

from . import SCHEMA


class GateError(Exception):
    def __init__(self, status, reason, *, details=None):
        self.status = status
        self.reason = reason
        self.details = details
        super().__init__(reason)


def utc():
    return datetime.now(timezone.utc).isoformat()


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write_json(path, value):
    path = Path(path)
    temporary = path.with_suffix(path.suffix + '.tmp')
    with temporary.open('w', encoding='utf-8') as f:
        json.dump(value, f, ensure_ascii=False, indent=2, default=str)
        f.flush()
        os.fsync(f.fileno())
    temporary.replace(path)


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def error_info(error):
    result = dict(type=type(error).__name__, message=str(error),
                  hresult=getattr(error, 'hresult', None))
    if getattr(error, 'details', None) is not None:
        result['details'] = error.details
    return result


def create_run(root):
    """Only a new directory outside every tracked asset tree is allowed."""
    root = Path(root).resolve()
    repo = Path(__file__).resolve().parents[2]
    if root == repo or root.is_relative_to(repo):
        raise ValueError('結果はリポジトリ外の新規フォルダーにしてください')
    root.mkdir(parents=True, exist_ok=False)
    return root


class Progress:
    def __init__(self, directory, run_id, case_id):
        self.directory = Path(directory)
        self.path = self.directory / 'progress.jsonl'
        self.stop = self.directory / 'STOP'
        self.run_id, self.case_id = run_id, case_id
        self.sequence = 0
        self.errors = []

    def check_stop(self):
        if self.stop.exists():
            raise GateError('INTERRUPTED', '親プロセスから停止要求。新たな編集/計算/保存を禁止')

    def emit(self, event, **details):
        self.sequence += 1
        record = dict(schema=SCHEMA, run_id=self.run_id, case_id=self.case_id,
                      sequence=self.sequence, utc=utc(), worker_pid=os.getpid(),
                      event=event, **details)
        with self.path.open('a', encoding='utf-8') as f:
            f.write(json.dumps(record, ensure_ascii=False, default=str) + '\n')
            f.flush()
            os.fsync(f.fileno())

    def call(self, name, action, timeout=60, *, cleanup=False):
        if not cleanup:
            self.check_stop()
        self.emit(name + '.before_call', timeout_seconds=timeout)
        started = time.monotonic()
        try:
            value = action()
        except BaseException as error:
            try:
                self.emit(name + '.failed', elapsed_seconds=time.monotonic() - started,
                          error=error_info(error))
            except OSError as log_error:
                self.errors.append(error_info(log_error))
            raise
        self.emit(name + '.after_call', elapsed_seconds=time.monotonic() - started)
        if not cleanup:
            self.check_stop()  # a late returning synchronous call cannot continue
            if time.monotonic() - started > timeout:
                raise GateError('TIMEOUT', name + ' exceeded deadline')
        return value
