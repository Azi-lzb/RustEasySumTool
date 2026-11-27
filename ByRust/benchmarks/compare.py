"""Reproducible end-to-end comparison with an isolated, synthetic XLSX workload.

Run with the project's development Python. Timed processes run sequentially,
alternate order, and publish a different output on every run. Fixture generation,
correctness verification and cleanup are outside the timing interval.
"""
from __future__ import annotations
import argparse
from datetime import datetime
import hashlib
import importlib.metadata
import json
import math
from pathlib import Path
import platform
import statistics
import subprocess
import sys
import time
import tomllib

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from summary_tool.demo import create_demo_assets


def synthetic(directory, file_count, rule_count, source_rows):
    directory.mkdir(parents=True)
    template = directory / '模板.xlsx'
    book = Workbook()
    sheet = book.active
    sheet.title = '业务明细'
    sheet.merge_cells('B2:M2')
    sheet['B2'] = '独立性能测试模板'
    algorithms = ['SUM', 'AVG', 'WAVG', 'MAX', 'MIN', 'COUNT', 'SUMPRODUCT', 'RATIO']
    aliases = ['FUN_1', 'FUN_3', 'FUN_2', 'FUN_4', 'FUN_5', 'FUN_6', 'FUN_7', 'FUN_8']
    for index in range(rule_count):
        row = 1 + (index * 7) % (source_rows - 3)
        arguments = [f'A{row}:A{row+2},B{row}:B{row+2}', f'A{row}',
                     f'A{row}:A{row+2},B{row}:B{row+2}', f'C{row}', f'D{row}',
                     f'E{row}', f'F{row},G{row}', f'H{row},I{row}']
        name = aliases[index % 8] if (index // 8) % 2 else algorithms[index % 8]
        cell = sheet.cell(index + 5, 12, f'{name}({arguments[index % 8]})')
        cell.number_format = '0.0000'
        cell.font = Font(name='微软雅黑', size=11, color='FF17365D')
        cell.fill = PatternFill('solid', fgColor='FFF3F6FA' if index % 2 else 'FFFFFFFF')
        cell.alignment = Alignment(horizontal='right')
    sheet.column_dimensions['L'].width = 30
    sheet.print_title_rows = '1:4'
    sheet.freeze_panes = 'C5'
    book.create_sheet('附页')['B2'] = '固定附页'
    book.save(template)
    book.close()
    sources = []
    for source_index in range(file_count):
        book = Workbook(write_only=True)
        sheet = book.create_sheet('业务明细')
        for row in range(1, source_rows + 1):
            values = [
                f'{(row % 7) + 1}.5%' if row % 17 == 0 else ((row % 7) + 1) / 100,
                100 + row + source_index,
                row * 10 + source_index,
                row * 20 + source_index,
                None if row % 11 == 0 else (0 if row % 13 == 0 else row),
                row + source_index + 1, 1.25 + source_index / 100,
                row + source_index + 2, row * 2 + source_index + 10,
            ]
            values += [row * column + source_index for column in range(10, 21)]
            sheet.append(values)
        source = directory / f'来源_{source_index+1:03}.xlsx'
        book.save(source)
        book.close()
        sources.append(source)
    return {'template': template, 'sources': sources, 'rule_count': rule_count,
            'source_rows': source_rows, 'source_columns': 20}


def business_values(path):
    book = load_workbook(path, read_only=True, data_only=False)
    try:
        return {sheet.title: {cell.coordinate: cell.value for row in sheet for cell in row
                              if cell.value is not None}
                for sheet in book if sheet.title != '运行日志'}
    finally:
        book.close()


def equivalent(left, right):
    assert left.keys() == right.keys()
    for sheet, cells in left.items():
        assert cells.keys() == right[sheet].keys()
        for coordinate, value in cells.items():
            actual = right[sheet][coordinate]
            if isinstance(value, (float, int)) and not isinstance(value, bool):
                assert math.isclose(value, actual, rel_tol=2e-15, abs_tol=0), (sheet, coordinate, value, actual)
            else:
                assert value == actual, (sheet, coordinate, value, actual)


def stats(values):
    return {'median_seconds': statistics.median(values), 'min_seconds': min(values),
            'max_seconds': max(values), 'samples_seconds': values}


def engine_versions(python_exe, rust_exe):
    packages = tomllib.loads((ROOT / 'ByRust/Cargo.lock').read_text(encoding='utf-8'))['package']
    return {
        'python_source': sys.version.split()[0],
        'python_openpyxl': importlib.metadata.version('openpyxl'),
        'python_calamine': importlib.metadata.version('python-calamine'),
        'python_packaged': subprocess.check_output([str(python_exe), '--version'], encoding='utf-8').strip(),
        'rust_packaged': subprocess.check_output([str(rust_exe), '--version'], encoding='utf-8').strip(),
        'rust_crates': {package['name']: package['version'] for package in packages
                        if package['name'] in {'calamine', 'umya-spreadsheet', 'bigdecimal'}},
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--repeats', type=int, default=5)
    parser.add_argument('--workdir', type=Path, default=ROOT / '.build/benchmarks')
    parser.add_argument('--report', type=Path, default=ROOT / 'ByRust/benchmarks/results.json')
    parser.add_argument('--python-exe', type=Path, default=ROOT / 'dist/Excel 模板汇总工具_v0.6.1.exe')
    parser.add_argument('--rust-exe', type=Path, default=ROOT / 'ByRust/target/release/excel-summary-rust.exe')
    args = parser.parse_args()
    if args.repeats < 3:
        parser.error('Use at least three measured repetitions')
    for executable in [args.python_exe, args.rust_exe]:
        if not executable.is_file():
            parser.error(f'Missing release executable: {executable}')
    session = args.workdir / datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    session.mkdir(parents=True)
    config = ROOT / 'config.xlsx'
    canonical = create_demo_assets(session / 'small')
    cases = [dict(name='small_demo', template=canonical.template_path,
                  sources=list(canonical.source_paths), rule_count=61,
                  source_rows=None, source_columns=None)]
    for name, count, rules, rows in [('medium', 30, 400, 200), ('large', 100, 1000, 500)]:
        print(f'Generating {name}: {count} files, {rules} rules, {rows}x20 source cells', flush=True)
        cases.append(dict(name=name, **synthetic(session / name, count, rules, rows)))
    runtimes = {
        'python_source': [sys.executable, str(Path(__file__).with_name('python_runner.py'))],
        'rust_release_json': [str(args.rust_exe)],
        'python_packaged_cli': [str(args.python_exe)],
        'rust_packaged_cli': [str(args.rust_exe)],
    }
    hashes = {str(path): hashlib.sha256(path.read_bytes()).hexdigest()
              for path in [config, *[path for case in cases for path in [case['template'], *case['sources']]]]}
    report = {'timestamp_local': datetime.now().isoformat(), 'timezone': 'Asia/Shanghai',
              'platform': platform.platform(), 'python': sys.version, 'repeats': args.repeats,
              'warmup_runs_per_runtime_per_case': 1,
              'method': 'Sequential fresh processes; one warmup, alternating order; warm OS file cache; unique output; perf_counter wall time and engine duration_seconds; fixture generation/validation excluded.',
              'binary_sha256': {key: hashlib.sha256(path.read_bytes()).hexdigest()
                                for key, path in [('python', args.python_exe), ('rust', args.rust_exe)]},
              'cases': []}
    report['versions'] = engine_versions(args.python_exe, args.rust_exe)
    machine = ROOT / '.build/benchmark-machine.json'
    if machine.is_file():
        report['machine'] = json.loads(machine.read_text(encoding='utf-8-sig'))
    for case in cases:
        print(f'Benchmarking {case["name"]}', flush=True)
        baseline = None
        case_results = {name: {'wall': [], 'engine': []} for name in runtimes}
        for repetition in range(args.repeats + 1):
            order = list(runtimes)
            if repetition % 2:
                order.reverse()
            for name in order:
                output = session / case['name'] / f'result_{name}_{repetition}.xlsx'
                command = [*runtimes[name], '--template', str(case['template']), '--sources',
                           *map(str, case['sources']), '--output', str(output), '--config', str(config)]
                if name == 'rust_release_json':
                    command.append('--json')
                start = time.perf_counter()
                completed = subprocess.run(command, cwd=ROOT, capture_output=True,
                                           text=True, encoding='utf-8', timeout=300)
                elapsed = time.perf_counter() - start
                if completed.returncode:
                    raise RuntimeError(f'{name} failed: {completed.stdout}\n{completed.stderr}')
                if not output.is_file():
                    raise RuntimeError(f'{name} did not publish an output')
                if name in {'python_packaged_cli', 'rust_packaged_cli'}:
                    terminal = completed.stdout + completed.stderr
                    assert '利率空白' not in terminal
                    assert '权重仍进入分母' not in terminal
                engine = None
                if name in {'python_source', 'rust_release_json'}:
                    payload = json.loads(completed.stdout)
                    assert payload['ok']
                    assert payload['result']['source_count'] == len(case['sources'])
                    assert payload['result']['rule_count'] == case['rule_count']
                    engine = payload['result']['duration_seconds']
                values = business_values(output)
                if baseline is None:
                    baseline = values
                else:
                    equivalent(baseline, values)
                output.unlink()
                if repetition:
                    case_results[name]['wall'].append(elapsed)
                    if engine is not None:
                        case_results[name]['engine'].append(engine)
                print(f'  {name} {"warmup" if repetition==0 else str(repetition)+"/"+str(args.repeats)}: wall={elapsed:.4f}s' + (f' engine={engine:.4f}s' if engine is not None else ''), flush=True)
        summaries = {name: {'wall': stats(values['wall']),
                            'engine': stats(values['engine']) if values['engine'] else None}
                     for name, values in case_results.items()}
        entry = {'name': case['name'], 'source_count': len(case['sources']),
                 'rule_count': case['rule_count'], 'source_rows': case['source_rows'],
                 'source_columns': case['source_columns'],
                 'source_bytes': sum(path.stat().st_size for path in case['sources']),
                 'results_verified': True, 'runtimes': summaries}
        entry['python_over_rust_engine_ratio'] = summaries['python_source']['engine']['median_seconds'] / summaries['rust_release_json']['engine']['median_seconds']
        entry['python_over_rust_packaged_ratio'] = summaries['python_packaged_cli']['wall']['median_seconds'] / summaries['rust_packaged_cli']['wall']['median_seconds']
        report['cases'].append(entry)
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
        print(f'  medians {case["name"]}: Python/Rust engine={entry["python_over_rust_engine_ratio"]:.2f}x, packaged={entry["python_over_rust_packaged_ratio"]:.2f}x', flush=True)
    assert all(hashlib.sha256(Path(path).read_bytes()).hexdigest() == digest for path, digest in hashes.items())
    report['completed'] = True
    report['input_hashes_verified'] = True
    report['all_outputs_verified'] = True
    report['ordinary_cli_warnings_absent'] = True
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print(f'All input hashes unchanged; all outputs equivalent; report: {args.report}', flush=True)


if __name__ == '__main__':
    main()
