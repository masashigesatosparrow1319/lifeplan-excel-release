"""Pre-COM identity handshake for the fresh tiny3 one-shot diagnostic."""
import argparse
import json
import os
from pathlib import Path
import sys
import time

from .common import sha
from .fresh_tiny3 import FRESH_SCHEMA, _write_new_json, require, validate_control


def main(request_path, *, process_identity=None, worker_main=None, pause=time.sleep):
    from . import functions24, worker
    process_identity = process_identity or functions24.process_identity
    worker_main = worker_main or worker.main
    request_path = Path(request_path).resolve()
    require(request_path.is_file() and request_path.stat().st_size <= 131072, 'REQUEST_FILE_INVALID')
    request = json.loads(request_path.read_text(encoding='utf-8-sig'))
    directory = request_path.parent
    actual = validate_control(request, directory, process_identity=process_identity)
    actual = dict(actual, sys_executable=sys.executable)
    proof = dict(schema=FRESH_SCHEMA, request_sha256=sha(request_path), process=actual)
    proof_path = directory / 'python-worker.json'
    _write_new_json(proof_path, proof)
    deadline = time.monotonic() + 15
    authorization = directory / 'launch-authorization.json'
    while not authorization.is_file():
        require(not (directory / 'STOP').exists(), 'PARENT_STOP_BEFORE_COM')
        require(time.monotonic() < deadline, 'LAUNCH_AUTHORIZATION_TIMEOUT')
        pause(.05)
    value = json.loads(authorization.read_text(encoding='utf-8-sig'))
    require(value == dict(schema=FRESH_SCHEMA, request_sha256=proof['request_sha256'],
                          worker_sha256=sha(proof_path)), 'LAUNCH_AUTHORIZATION_MISMATCH')
    return worker_main(request_path)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--request', required=True)
    try:
        code = main(parser.parse_args().request)
    except BaseException:
        # No Excel has been started if the wrapper gate fails.
        code = 2
    raise SystemExit(code)
