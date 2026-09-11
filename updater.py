import argparse
import json
import os
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path


def msg(title, text, error=False):
    try:
        import tkinter as tk
        from tkinter import messagebox
        root=tk.Tk(); root.withdraw()
        (messagebox.showerror if error else messagebox.showinfo)(title, text, parent=root)
        root.destroy()
    except Exception:
        pass


def wait_pid(pid, timeout=120):
    if not pid:
        return
    end=time.time()+timeout
    while time.time()<end:
        try:
            if os.name=='nt':
                import ctypes
                PROCESS_QUERY_LIMITED_INFORMATION=0x1000
                handle=ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
                if not handle:
                    return
                ctypes.windll.kernel32.CloseHandle(handle)
            else:
                os.kill(int(pid),0)
        except Exception:
            return
        time.sleep(0.5)


def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--zip', required=True)
    ap.add_argument('--dir', required=True)
    ap.add_argument('--pid', type=int, default=0)
    args=ap.parse_args()
    zpath=Path(args.zip).resolve(); install=Path(args.dir).resolve()
    if not zpath.exists():
        msg('PhotoArchiveCatalog','Файл обновления не найден.',True); return 2
    wait_pid(args.pid)
    temp=Path(tempfile.mkdtemp(prefix='PhotoArchive_update_'))
    backup=temp/'backup'; unpack=temp/'unpack'; backup.mkdir(); unpack.mkdir()
    try:
        with zipfile.ZipFile(str(zpath),'r') as z:
            z.extractall(str(unpack))
        manifest=unpack/'update_manifest.json'
        if manifest.exists():
            data=json.loads(manifest.read_text(encoding='utf-8'))
            if data.get('app') not in (None,'PhotoArchiveCatalog'):
                raise RuntimeError('Этот пакет обновления предназначен для другой программы.')
        files=[p for p in unpack.rglob('*') if p.is_file() and p.name!='update_manifest.json']
        if not files:
            raise RuntimeError('В ZIP-пакете нет файлов обновления.')
        for src in files:
            rel=src.relative_to(unpack)
            dst=install/rel
            dst.parent.mkdir(parents=True,exist_ok=True)
            if dst.exists():
                b=backup/rel; b.parent.mkdir(parents=True,exist_ok=True); shutil.copy2(str(dst),str(b))
            shutil.copy2(str(src),str(dst))
        exe=install/'PhotoArchiveCatalog.exe'
        if exe.exists():
            os.startfile(str(exe)) if os.name=='nt' else None
        msg('PhotoArchiveCatalog','Обновление успешно установлено.')
        return 0
    except Exception as exc:
        try:
            for src in backup.rglob('*'):
                if src.is_file():
                    dst=install/src.relative_to(backup); dst.parent.mkdir(parents=True,exist_ok=True); shutil.copy2(str(src),str(dst))
        except Exception:
            pass
        msg('PhotoArchiveCatalog',f'Не удалось установить обновление:\n{exc}',True)
        return 1
    finally:
        try: shutil.rmtree(str(temp),ignore_errors=True)
        except Exception: pass


if __name__=='__main__':
    raise SystemExit(main())
