from pathlib import Path
import sys
p = Path(r"C:\Users\suraj.yadav\.cache\torch\sentence_transformers")
if not p.exists():
    print(f'Path not found: {p}')
    sys.exit(0)

files = []
for f in p.rglob('*'):
    if f.is_file():
        try:
            size = f.stat().st_size
        except Exception:
            size = 0
        files.append((f, size))

files.sort(key=lambda x: x[1], reverse=True)
print(f'Found {len(files)} files under {p}\n')
for f, size in files:
    print(f"{size/1024/1024:8.2f} MB  | {f}")
