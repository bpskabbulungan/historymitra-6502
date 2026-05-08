from __future__ import annotations

import argparse
import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

from .api import MitraApiClient, MitraApiConfig, MitraApiError
from .parsers import (
    load_history_documents_from_file,
    load_mitra_detail_from_file,
    load_mitra_list_from_file,
)
from .reporting import (
    ReportBundle,
    build_rekap_mitra,
    build_rekap_survey,
    enrich_mitra_row_with_detail,
    export_excel,
    format_terminal_summary,
    normalize_history_documents,
    normalize_mitra_row,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="historymitra",
        description="Fetcher dan exporter histori survei mitra BPS Bulungan.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    run_parser = subparsers.add_parser("run", help="Jalankan pipeline fetch/parse dan export.")
    run_parser.add_argument("--source", choices=["sample", "live"], required=True)
    run_parser.add_argument("--year", type=int, default=2026)
    run_parser.add_argument("--prov", default="65")
    run_parser.add_argument("--kab", default="02")
    run_parser.add_argument("--output-dir", default="output")
    run_parser.add_argument("--excel-name", default="mitra_history_report.xlsx")
    run_parser.add_argument("--limit", type=int, default=0, help="Batasi jumlah mitra saat mode live.")
    run_parser.add_argument("--workers", type=int, default=8)
    run_parser.add_argument("--table-file", help="Path ke file dump daftar mitra.")
    run_parser.add_argument("--detail-file", help="Path ke file dump detail satu mitra.")
    run_parser.add_argument("--history-file", help="Path ke file dump histori mitra.")
    return parser


def ensure_output_dirs(output_dir: Path) -> Path:
    raw_dir = output_dir / "raw"
    raw_dir.mkdir(parents=True, exist_ok=True)
    return raw_dir


def write_json(path: Path, payload: Any) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def run_sample_mode(args: argparse.Namespace) -> ReportBundle:
    if not args.table_file:
        raise SystemExit("--table-file wajib diisi saat --source sample.")
    if not args.history_file:
        raise SystemExit("--history-file wajib diisi saat --source sample.")

    output_dir = Path(args.output_dir)
    raw_dir = ensure_output_dirs(output_dir)

    mitra_list = load_mitra_list_from_file(args.table_file)
    mitra_rows = [normalize_mitra_row(item) for item in mitra_list]

    detail_payload = load_mitra_detail_from_file(args.detail_file) if args.detail_file else {}
    history_documents = load_history_documents_from_file(args.history_file)

    detail_id = str(detail_payload.get("idmitra", "")).strip() if detail_payload else ""
    detail_by_id = {
        row["id_mitra"]: row for row in mitra_rows
    }
    if detail_id and detail_id in detail_by_id:
        detail_by_id[detail_id] = enrich_mitra_row_with_detail(detail_by_id[detail_id], detail_payload)

    if detail_id and detail_id in detail_by_id:
        target_mitra = detail_by_id[detail_id]
    elif mitra_rows:
        target_mitra = mitra_rows[0]
    else:
        raise SystemExit("Data mitra dari sample kosong.")

    history_rows = normalize_history_documents(
        id_mitra=target_mitra["id_mitra"],
        mitra_name=target_mitra["nama_lengkap"],
        sobat_id=target_mitra["sobat_id"],
        history_documents=history_documents,
    )

    mitra_rows = [
        detail_by_id.get(row["id_mitra"], row)
        for row in mitra_rows
    ]
    rekap_mitra_rows = build_rekap_mitra(history_rows)
    rekap_survey_rows = build_rekap_survey(history_rows)

    write_json(raw_dir / "mitra_list.json", mitra_list)
    write_json(raw_dir / "mitra_histories.json", history_documents)

    return ReportBundle(
        mitra_rows=mitra_rows,
        history_rows=history_rows,
        rekap_mitra_rows=rekap_mitra_rows,
        rekap_survey_rows=rekap_survey_rows,
        source_label=f"sample:{args.table_file}",
        output_dir=output_dir,
    )


def fetch_history_rows_live(
    client: MitraApiClient,
    mitra_rows: list[dict[str, str]],
    *,
    year: int,
    workers: int,
) -> tuple[list[dict[str, str]], list[str]]:
    history_rows: list[dict[str, str]] = []
    errors: list[str] = []
    mitra_by_id = {row["id_mitra"]: row for row in mitra_rows}

    with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        future_map = {
            executor.submit(client.fetch_mitra_history_documents, id_mitra=row["id_mitra"], year=year): row["id_mitra"]
            for row in mitra_rows
            if row["id_mitra"]
        }
        for future in as_completed(future_map):
            id_mitra = future_map[future]
            mitra = mitra_by_id[id_mitra]
            try:
                documents = future.result()
            except Exception as exc:
                errors.append(f"{id_mitra} - {mitra['nama_lengkap']}: {exc}")
                continue
            history_rows.extend(
                normalize_history_documents(
                    id_mitra=id_mitra,
                    mitra_name=mitra["nama_lengkap"],
                    sobat_id=mitra["sobat_id"],
                    history_documents=documents,
                )
            )
    history_rows.sort(key=lambda row: (row["nama_lengkap"], row["id_mitra"], row["history_doc_index"], row["history_order"]))
    return history_rows, errors


def run_live_mode(args: argparse.Namespace) -> ReportBundle:
    output_dir = Path(args.output_dir)
    raw_dir = ensure_output_dirs(output_dir)
    client = MitraApiClient(MitraApiConfig.from_env())
    mitra_list = client.fetch_mitra_list(year=args.year, prov=args.prov, kab=args.kab)
    if args.limit and args.limit > 0:
        mitra_list = mitra_list[: args.limit]

    mitra_rows = [normalize_mitra_row(item) for item in mitra_list]
    history_rows, errors = fetch_history_rows_live(
        client,
        mitra_rows,
        year=args.year,
        workers=args.workers,
    )

    rekap_mitra_rows = build_rekap_mitra(history_rows)
    rekap_survey_rows = build_rekap_survey(history_rows)

    write_json(raw_dir / "mitra_list.json", mitra_list)
    write_json(raw_dir / "mitra_histories.json", history_rows)
    if errors:
        write_json(raw_dir / "errors.json", errors)

    return ReportBundle(
        mitra_rows=mitra_rows,
        history_rows=history_rows,
        rekap_mitra_rows=rekap_mitra_rows,
        rekap_survey_rows=rekap_survey_rows,
        source_label=f"live:{args.year}/{args.prov}/{args.kab}",
        output_dir=output_dir,
    )


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "run":
            if args.source == "sample":
                bundle = run_sample_mode(args)
            else:
                bundle = run_live_mode(args)
            excel_path = export_excel(bundle, filename=args.excel_name)
            print(format_terminal_summary(bundle))
            print(f"\nWorkbook Excel      : {excel_path}")
            print(f"Raw JSON output     : {bundle.output_dir / 'raw'}")
            return 0
    except MitraApiError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
