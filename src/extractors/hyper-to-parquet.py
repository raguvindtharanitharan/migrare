"""
hyper-to-parquet.py
Extract all tables from a Tableau .hyper file and write one .parquet per table.

Usage:
    python hyper-to-parquet.py <hyper_path> <output_dir>

Output (stdout):
    JSON: { "<schema>.<table>": "<output_dir>/<slug>.parquet", ... }

Exit codes:
    0  success
    1  missing arguments
    2  import error (tableauhyperapi or pyarrow not installed)
    3  runtime error
"""

import sys
import os
import json
import re

def main():
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: hyper-to-parquet.py <hyper_path> <output_dir> [slug_prefix]"}), file=sys.stderr)
        sys.exit(1)

    hyper_path = sys.argv[1]
    output_dir = sys.argv[2]
    slug_prefix = sys.argv[3] if len(sys.argv) > 3 else None

    try:
        from tableauhyperapi import HyperProcess, Connection, Telemetry, TableName
    except ImportError:
        print(json.dumps({"error": "tableauhyperapi not installed. Run: pip install tableauhyperapi"}), file=sys.stderr)
        sys.exit(2)

    try:
        import pyarrow as pa
        import pyarrow.parquet as pq
    except ImportError:
        print(json.dumps({"error": "pyarrow not installed. Run: pip install pyarrow"}), file=sys.stderr)
        sys.exit(2)

    os.makedirs(output_dir, exist_ok=True)
    result = {}

    try:
        with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
            with Connection(endpoint=hyper.endpoint, database=hyper_path) as conn:
                catalog = conn.catalog
                schemas = catalog.get_schema_names()

                for schema_name in schemas:
                    if schema_name.name.unescaped in ('pg_catalog', 'information_schema'):
                        continue

                    tables = catalog.get_table_names(schema_name)
                    for table_name in tables:
                        fq = f'{schema_name.name.unescaped}.{table_name.name.unescaped}'

                        table_def = catalog.get_table_definition(table_name)
                        columns = [col.name.unescaped for col in table_def.columns]
                        col_types = [col.type for col in table_def.columns]

                        rows = conn.execute_list_query(f'SELECT * FROM {table_name}')

                        if not rows:
                            continue

                        # Build PyArrow arrays per column
                        arrays = []
                        for i, (col, hyper_type) in enumerate(zip(columns, col_types)):
                            values = [row[i] for row in rows]
                            arr = to_arrow_array(values, hyper_type)
                            arrays.append(arr)

                        schema = pa.schema([
                            pa.field(col, arr.type)
                            for col, arr in zip(columns, arrays)
                        ])
                        table = pa.table(dict(zip(columns, arrays)), schema=schema)

                        slug = slugify(table_name.name.unescaped)
                        final_slug = f'{slug_prefix}-{slug}' if slug_prefix else slug
                        out_path = os.path.join(output_dir, f'{final_slug}.parquet')
                        pq.write_table(table, out_path, compression='snappy')
                        result[fq] = out_path

    except Exception as e:
        print(json.dumps({"error": str(e)}), file=sys.stderr)
        sys.exit(3)

    print(json.dumps(result))


def to_arrow_array(values, hyper_type):
    """Convert a list of Python values from tableauhyperapi to a PyArrow array."""
    import pyarrow as pa

    type_name = str(hyper_type).lower()

    if 'int' in type_name or 'smallint' in type_name or 'bigint' in type_name:
        return pa.array(values, type=pa.int64())
    if 'double' in type_name or 'float' in type_name or 'numeric' in type_name or 'decimal' in type_name:
        return pa.array(values, type=pa.float64())
    if 'bool' in type_name:
        return pa.array(values, type=pa.bool_())
    if 'date' in type_name and 'time' not in type_name:
        # tableauhyperapi returns datetime.date objects
        str_vals = [str(v) if v is not None else None for v in values]
        return pa.array(str_vals, type=pa.string())
    if 'timestamp' in type_name or 'datetime' in type_name:
        str_vals = [str(v) if v is not None else None for v in values]
        return pa.array(str_vals, type=pa.string())
    # Default: string
    str_vals = [str(v) if v is not None else None for v in values]
    return pa.array(str_vals, type=pa.string())


def slugify(name: str) -> str:
    return re.sub(r'[^a-z0-9]+', '-', name.lower()).strip('-')


if __name__ == '__main__':
    main()
