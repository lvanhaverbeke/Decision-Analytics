"""Preserve workbook checkpoints while enforcing a command's wall-clock limit."""

import os
from pathlib import Path
import signal
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import datetime, timedelta, timezone


def valid_workbook(path):
    try:
        with zipfile.ZipFile(path) as archive:
            return (
                {'[Content_Types].xml', 'xl/workbook.xml'} <= set(archive.namelist())
                and archive.testzip() is None
            )
    except (OSError, zipfile.BadZipFile, EOFError):
        return False


def snapshot(source, destination):
    """Validate a private copy before atomically replacing the checkpoint."""
    if not source.is_file():
        return False
    with tempfile.TemporaryDirectory(dir=destination.parent, prefix='.checkpoint-') as tmp:
        candidate = Path(tmp) / 'model.xlsx'
        try:
            candidate.write_bytes(source.read_bytes())
            if not valid_workbook(candidate):
                return False
            os.replace(candidate, destination)
        except OSError:
            return False
    return True


def main():
    folder = Path(sys.argv[1]).resolve()
    save_after, limit = map(float, sys.argv[2:4])
    if not 0 < save_after < limit:
        raise ValueError('Save deadline must precede the hard timeout.')
    model = folder / 'model.xlsx'
    latest = folder / '.model_latest.xlsx'
    checkpoint = folder / 'model_6m30.xlsx'
    request = folder / 'SAVE_NOW.txt'
    for old in (request, latest, checkpoint):
        old.unlink(missing_ok=True)
    prompt = sys.stdin.buffer.read()
    deadline = datetime.now(timezone.utc) + timedelta(seconds=save_after)
    prompt += (
        f'\n\nSave deadline (UTC): {deadline.isoformat()}. '
        'Save model.xlsx before this deadline; save a basic version immediately '
        'and update it incrementally. Check SAVE_NOW.txt after each step.\n'
    ).encode()
    start = time.monotonic()
    requested = False
    with tempfile.TemporaryFile() as stdin:
        stdin.write(prompt)
        stdin.seek(0)
        process = subprocess.Popen(sys.argv[4:], stdin=stdin, start_new_session=True)
        try:
            while True:
                snapshot(model, latest)
                elapsed = time.monotonic() - start
                if not requested and elapsed >= save_after:
                    request.write_text(
                        'De save-deadline is bereikt. Sla model.xlsx NU op en rond af.\n'
                    )
                    requested = True
                    if snapshot(latest, checkpoint):
                        print('SAVE-CHECKPOINT: model_6m30.xlsx opgeslagen.', flush=True)
                    else:
                        print('SAVE-CHECKPOINT: nog geen geldige spreadsheet op schijf.', flush=True)
                status = process.poll()
                if status is not None:
                    break
                if elapsed >= limit:
                    os.killpg(process.pid, signal.SIGKILL)
                    process.wait()
                    status = 124
                    print('TIJDSLIMIET: run gestopt na 7 minuten.', flush=True)
                    break
                time.sleep(min(1, limit - elapsed, max(0.01, save_after - elapsed) if not requested else 1))
        except KeyboardInterrupt:
            os.killpg(process.pid, signal.SIGKILL)
            process.wait()
            return 130
    # Keep a usable on-disk result even if the final save was interrupted.
    if not valid_workbook(model) and snapshot(latest, model):
        print('HERSTEL: laatste geldige spreadsheet bewaard als model.xlsx.', flush=True)
    return status if status >= 0 else 128 - status


if __name__ == '__main__':
    sys.exit(main())
