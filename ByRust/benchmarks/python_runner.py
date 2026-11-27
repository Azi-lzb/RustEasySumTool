"""Machine-readable Python reference runner used only by the benchmark."""
import argparse
import dataclasses
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from summary_tool.engine import aggregate

parser = argparse.ArgumentParser()
parser.add_argument('--template', type=Path, required=True)
parser.add_argument('--sources', type=Path, nargs='+', required=True)
parser.add_argument('--output', type=Path, required=True)
parser.add_argument('--config', type=Path, required=True)
args = parser.parse_args()
result = aggregate(args.template, args.sources, args.output, config_path=args.config)
print(json.dumps({'ok': True, 'result': dataclasses.asdict(result)}, default=str))
