import shutil
from pathlib import Path
import sys

hub = Path.home() / '.cache' / 'huggingface' / 'hub'
keep = [
    'models--sentence-transformers--all-mpnet-base-v2',
    'models--sentence-transformers--all-MiniLM-L6-v2'
]

# Step 1: Download the fast model into cache
try:
    from sentence_transformers import SentenceTransformer
    print('Downloading fast model all-MiniLM-L6-v2...')
    SentenceTransformer('all-MiniLM-L6-v2')
    print('Download complete.')
except Exception as e:
    print('Failed to download model: ', e)
    print('You may need to activate the virtualenv or install sentence-transformers in this environment.')

# Step 2: Delete everything under hub except keep list
if not hub.exists():
    print(f'No hub cache found at {hub}. Nothing to prune.')
    sys.exit(0)

deleted = []
for child in hub.iterdir():
    if child.is_dir() and child.name not in keep:
        try:
            print(f'Deleting {child}...')
            shutil.rmtree(child)
            deleted.append(child.name)
        except Exception as e:
            print(f'Failed to delete {child}: {e}')

print('Deleted folders:')
for d in deleted:
    print('-', d)
print('Done.')
