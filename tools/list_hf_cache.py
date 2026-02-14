import os
from pathlib import Path

hub = Path.home() / '.cache' / 'huggingface' / 'hub'
if not hub.exists():
    print(f'No huggingface hub cache at: {hub}')
    raise SystemExit(0)

entries = []
for child in hub.iterdir():
    if child.is_dir():
        total = 0
        for f in child.rglob('*'):
            if f.is_file():
                try:
                    total += f.stat().st_size
                except Exception:
                    pass
        entries.append((child.name, str(child), total))

entries.sort(key=lambda x: x[2], reverse=True)
print(f"Found {len(entries)} entries in {hub}\n")
for name, path, size in entries:
    mb = round(size / (1024*1024), 2)
    print(f"{mb:8.2f} MB  | {name}  | {path}")
