import shutil
from pathlib import Path
import sys

cache_dir = Path(r"C:\Users\suraj.yadav\.cache\torch\sentence_transformers")
if not cache_dir.exists():
    print(f'Cache directory not found: {cache_dir}')
    sys.exit(1)

# Keep any directories whose name contains these substrings
keep_substrings = ['all-mpnet-base-v2', 'all-MiniLM-L6-v2']

deleted = []
failed = []
for child in cache_dir.iterdir():
    if not child.is_dir():
        continue
    name = child.name
    if any(sub in name for sub in keep_substrings):
        print(f'Keeping: {name}')
        continue
    try:
        print(f'Deleting: {child}')
        shutil.rmtree(child)
        deleted.append(name)
    except Exception as e:
        print(f'Failed to delete {child}: {e}')
        failed.append((name, str(e)))

print('\nDeleted directories:')
for d in deleted:
    print('-', d)
if failed:
    print('\nFailed to delete:')
    for name, err in failed:
        print(name, err)
print('\nDone.')
