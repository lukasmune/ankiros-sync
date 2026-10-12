import argparse
import re
import sqlite3
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

import config


FILTERS = {"all", "direct", "co-exhibitors"}
PARTICIPANT_ENDPOINT = "api/v1.0/participant/getAllParticipantList"
DETAIL_ENDPOINT = "api/v1.0/participant/getCompanyDetailBySlug"
EXCEL_ILLEGAL_CHARACTERS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
RECORD_FIELDS = (
    "company_name",
    "slug",
    "relationship",
    "parent_company",
    "address",
    "stand_no",
    "phone",
    "email",
    "description",
    "products",
)


def request_json(session, endpoint, payload):
    url = config.API_BASE_URL + endpoint
    for attempt in range(config.REQUEST_RETRIES + 1):
        response = session.post(
            url,
            json=payload,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json",
                "User-Agent": "ankiros-sync/1.0",
            },
            timeout=config.REQUEST_TIMEOUT,
        )
        if response.status_code != 429:
            response.raise_for_status()
            body = response.json()
            if body.get("hasError"):
                raise RuntimeError(
                    body.get("message") or f"ANKIROS API error: {endpoint}"
                )
            return body.get("result") or []
        if attempt == config.REQUEST_RETRIES:
            response.raise_for_status()
        retry_after = response.headers.get("Retry-After")
        try:
            delay = max(float(retry_after), config.REQUEST_DELAY)
        except (TypeError, ValueError):
            delay = config.REQUEST_DELAY * (2 ** attempt)
        time.sleep(delay)
    raise RuntimeError("ANKIROS request retry loop completed unexpectedly.")


def fetch_exhibitors(session=None):
    """Fetch direct exhibitors and expand their co-exhibitor details."""

    session = session or requests.Session()
    direct = request_json(
        session,
        PARTICIPANT_ENDPOINT,
        {"expoId": config.EXPO_ID},
    )
    records = []
    for item in direct:
        parent = _make_record(item, "direct")
        records.append(parent)
        if not item.get("slug"):
            continue
        detail = request_json(
            session,
            DETAIL_ENDPOINT,
            {"expoId": config.EXPO_ID, "slug": item["slug"]},
        )
        parent.update(_detail_fields(detail))
        for co_exhibitor in detail.get("subcoCompany", []) or []:
            if not isinstance(co_exhibitor, dict):
                continue
            co_record = _make_record(
                co_exhibitor,
                "co-exhibitor",
                parent["company_name"],
            )
            if co_record["slug"]:
                try:
                    co_detail = request_json(
                        session,
                        DETAIL_ENDPOINT,
                        {"expoId": config.EXPO_ID, "slug": co_record["slug"]},
                    )
                except RuntimeError as error:
                    print(
                        f"Warning: could not enrich co-exhibitor "
                        f"{co_record['company_name']!r}: {error}",
                        file=sys.stderr,
                    )
                else:
                    co_record.update(_detail_fields(co_detail))
            records.append(co_record)
    return records


def _make_record(item, relationship, parent_company=None):
    return {
        "company_name": (
            item.get("companyName")
            or item.get("subcoCompanyName")
            or item.get("name")
            or ""
        ).strip(),
        "slug": item.get("slug"),
        "relationship": relationship,
        "parent_company": parent_company,
        "address": None,
        "stand_no": None,
        "phone": None,
        "email": None,
        "description": None,
        "products": None,
    }


def _detail_fields(detail):
    products = [
        item.get("titleEng", "").strip()
        for item in detail.get("exhibitList", []) or []
        if isinstance(item, dict) and item.get("titleEng")
    ]
    return {
        "address": detail.get("address"),
        "stand_no": detail.get("standNo"),
        "phone": detail.get("phone"),
        "email": detail.get("email"),
        "description": detail.get("detailEng") or detail.get("detailTr"),
        "products": "; ".join(products) or None,
    }


def filter_exhibitors(records, relationship_filter="all"):
    """Filter direct and co-exhibitor records without changing their data."""

    if relationship_filter not in FILTERS:
        raise ValueError(
            f"Unknown filter {relationship_filter!r}; choose from "
            f"{', '.join(sorted(FILTERS))}"
        )
    if relationship_filter == "all":
        return list(records)
    relationship = (
        "direct" if relationship_filter == "direct" else "co-exhibitor"
    )
    return [
        record for record in records
        if record.get("relationship") == relationship
    ]


def initialize_database(path=config.DATABASE_PATH):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(path) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS exhibitors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_name TEXT NOT NULL,
                slug TEXT,
                relationship TEXT NOT NULL,
                parent_company TEXT,
                address TEXT,
                stand_no TEXT,
                phone TEXT,
                email TEXT,
                description TEXT,
                products TEXT,
                edition INTEGER NOT NULL,
                first_seen_at TEXT NOT NULL,
                UNIQUE(company_name, relationship, parent_company, edition)
            )
            """
        )
        columns = {
            row[1] for row in connection.execute("PRAGMA table_info(exhibitors)")
        }
        for column in RECORD_FIELDS[4:]:
            if column not in columns:
                connection.execute(f"ALTER TABLE exhibitors ADD COLUMN {column} TEXT")


def sync_records(records, path=config.DATABASE_PATH):
    initialize_database(path)
    now = datetime.now(timezone.utc).isoformat()
    with sqlite3.connect(path) as connection:
        for record in records:
            connection.execute(
                """
                INSERT INTO exhibitors (
                    company_name, slug, relationship, parent_company,
                    address, stand_no, phone, email, description, products,
                    edition, first_seen_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(company_name, relationship, parent_company, edition)
                DO UPDATE SET
                    slug = excluded.slug,
                    address = excluded.address,
                    stand_no = excluded.stand_no,
                    phone = excluded.phone,
                    email = excluded.email,
                    description = excluded.description,
                    products = excluded.products
                """,
                (
                    record["company_name"],
                    record.get("slug"),
                    record["relationship"],
                    record.get("parent_company"),
                    record.get("address"),
                    record.get("stand_no"),
                    record.get("phone"),
                    record.get("email"),
                    record.get("description"),
                    record.get("products"),
                    config.EDITION,
                    now,
                ),
            )


def export_records(records, output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    summary.append(["Field", "Value"])
    summary.append(["Event", "ANKIROS"])
    summary.append(["Edition", config.EDITION])
    summary.append(["Total Exhibitors", len(records)])
    summary.append([
        "Direct Exhibitors",
        sum(record["relationship"] == "direct" for record in records),
    ])
    summary.append([
        "Co-exhibitors",
        sum(record["relationship"] == "co-exhibitor" for record in records),
    ])
    for cell in summary[1]:
        cell.font = Font(bold=True)

    companies = workbook.create_sheet("Company Master")
    companies.append(["Company Name", "Relationship", "Parent Company", "Slug"])
    for record in records:
        companies.append(_excel_row([
            record["company_name"],
            record["relationship"],
            record.get("parent_company"),
            record.get("slug"),
        ]))

    worksheet = workbook.create_sheet("Sales Leads")
    worksheet.append([
        "Company Name", "Relationship", "Parent Company", "Address",
        "Stand", "Phone", "Email", "Description", "Product Categories",
        "Slug", "Edition",
    ])
    for record in records:
        worksheet.append(_excel_row([
            record["company_name"],
            record["relationship"],
            record.get("parent_company"),
            record.get("address"),
            record.get("stand_no"),
            record.get("phone"),
            record.get("email"),
            record.get("description"),
            record.get("products"),
            record.get("slug"),
            config.EDITION,
        ]))
    worksheet.freeze_panes = "A2"
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    if worksheet.max_row > 1:
        table = Table(displayName="AnkirosSalesLeads", ref=worksheet.dimensions)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)
    for column in worksheet.columns:
        worksheet.column_dimensions[column[0].column_letter].width = min(
            max(len(str(cell.value or "")) for cell in column) + 2,
            50,
        )
    workbook.save(output_path)
    return output_path


def _excel_row(values):
    return [
        EXCEL_ILLEGAL_CHARACTERS.sub("", value)
        if isinstance(value, str)
        else value
        for value in values
    ]


def main(argv=None):
    parser = argparse.ArgumentParser(description="Sync ANKIROS exhibitors")
    subparsers = parser.add_subparsers(dest="command", required=True)
    for command in ("sync", "export"):
        subparser = subparsers.add_parser(command)
        subparser.add_argument("--filter", choices=sorted(FILTERS), default="all")
    args = parser.parse_args(argv)
    records = filter_exhibitors(
        fetch_exhibitors(),
        args.filter,
    )
    if args.command == "sync":
        sync_records(records)
    output = config.EXPORTS_DIR / f"ANKIROS_{config.EDITION}.xlsx"
    export_records(records, output)
    print(f"{len(records)} {args.filter} exhibitors processed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
