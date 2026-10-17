"""Pre-COM identity handshake for the fresh tiny3 one-shot diagnostic."""
import argparse
import json
import os
from pathlib import Path
import sys
import time

from .common import sha
from .fresh_tiny3 import (FRESH_SCHEMA, Blocked, _write_new_json, require,
                          validate_control, write_precom_failure)


def main(request_path, *, process_identity=None, worker_main=None, pause=time.sleep,
         write_new=_write_new_json):
    from . import functions24
    process_identity = process_identity or functions24.process_identity
    request_path = Path(request_path).resolve()
    directory = request_path.parent
    stage = 'WORKER_REQUEST_PARSE'

    def reason_for(stage_value, error):
        fixed = str(error) if isinstance(error, Blocked) else ''
        if stage_value == 'WORKER_REQUEST_PARSE':
            return 'REQUEST_FILE_INVALID' if fixed == 'REQUEST_FILE_INVALID' else 'REQUEST_JSON_INVALID'
        if isinstance(error, RuntimeError) and not fixed \
                and stage_value != 'WORKER_PROCESS_IDENTITY':
            return 'PRECOM_UNKNOWN'
        if stage_value == 'WORKER_CONTROL_VALIDATE': return 'CONTROL_BINDING_INVALID'
        if stage_value == 'WORKER_MANIFEST_VALIDATE': return 'MANIFEST_INVALID'
        if stage_value == 'WORKER_MARKER_VALIDATE': return 'APPROVAL_MARKER_INVALID'
        if stage_value == 'WORKER_PROFILE_VALIDATE': return 'PROFILE_BINDING_INVALID'
        if stage_value == 'WORKER_PROCESS_IDENTITY':
            return ('PARENT_IDENTITY_MISMATCH' if fixed == 'ACTUAL_PARENT_MISMATCH'
                    else 'PROCESS_IDENTITY_UNAVAILABLE')
        if stage_value == 'WORKER_PROOF_CREATE': return 'PROOF_CREATE_FAILED'
        if stage_value == 'WORKER_AUTHORIZATION_WAIT':
            return 'PARENT_STOPPED' if fixed == 'PARENT_STOP_BEFORE_COM' else 'AUTHORIZATION_TIMEOUT'
        if stage_value == 'WORKER_AUTHORIZATION_VALIDATE': return 'AUTHORIZATION_INVALID'
        if stage_value == 'WORKER_HANDOFF_CHECK':
            return ('PARENT_STOPPED' if fixed == 'PARENT_STOP_BEFORE_COM'
                    else 'PRECOM_FAILURE_BEFORE_HANDOFF')
        return 'PRECOM_UNKNOWN'

    try:
        require(request_path.is_file() and request_path.stat().st_size <= 131072,
                'REQUEST_FILE_INVALID')
        request = json.loads(request_path.read_text(encoding='utf-8-sig'))
        def observed(value):
            nonlocal stage
            stage = value
        stage = 'WORKER_CONTROL_VALIDATE'
        actual = validate_control(request, directory, process_identity=process_identity,
                                  observe_stage=observed)
        actual = dict(actual, sys_executable=sys.executable)
        proof = dict(schema=FRESH_SCHEMA, request_sha256=sha(request_path), process=actual)
        proof_path = directory / 'python-worker.json'
        stage = 'WORKER_PROOF_CREATE'
        write_new(proof_path, proof)
        deadline = time.monotonic() + 15
        authorization = directory / 'launch-authorization.json'
        stage = 'WORKER_AUTHORIZATION_WAIT'
        while not authorization.is_file():
            require(not (directory / 'STOP').exists(), 'PARENT_STOP_BEFORE_COM')
            require(time.monotonic() < deadline, 'LAUNCH_AUTHORIZATION_TIMEOUT')
            pause(.05)
        stage = 'WORKER_AUTHORIZATION_VALIDATE'
        value = json.loads(authorization.read_text(encoding='utf-8-sig'))
        require(value == dict(schema=FRESH_SCHEMA, request_sha256=proof['request_sha256'],
                              worker_sha256=sha(proof_path)), 'LAUNCH_AUTHORIZATION_MISMATCH')
        stage = 'WORKER_HANDOFF_CHECK'
        require(not (directory / 'STOP').exists(), 'PARENT_STOP_BEFORE_COM')
        require(not (directory / 'precom-failure.json').exists(),
                'PRECOM_FAILURE_BEFORE_HANDOFF')
    except BaseException as error:
        write_precom_failure(directory, stage, reason_for(stage, error), write_new=write_new)
        raise
    if worker_main is None:
        from . import worker
        worker_main = worker.main
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
