from __future__ import annotations

import csv
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook


WOS_DIR = Path(r"D:\数据\WOS数据（数据不超过10w条的学校）")
OUTPUT_DIR = Path("outputs")
YEAR_MIN = 2000
YEAR_MAX = 2022


SCHOOLS = [
    {
        "school_cn": "北京理工大学",
        "wos_file": "Beijing Institute of Technology.xlsx",
        "treated": 1,
        "relo_year": 2007,
        "sample_role": "original_treat",
        "note": "良乡校区；同城新建多校区",
    },
    {
        "school_cn": "中央财经大学",
        "wos_file": "Central University of Finance & Economics.xlsx",
        "treated": 1,
        "relo_year": 2009,
        "sample_role": "original_treat",
        "note": "沙河校区；同城新建多校区",
    },
    {
        "school_cn": "暨南大学",
        "wos_file": "Jinan University.xlsx",
        "treated": 1,
        "relo_year": 2014,
        "sample_role": "original_treat",
        "note": "番禺校区；同城新建多校区；不可误用 University of Jinan",
    },
    {
        "school_cn": "华东理工大学",
        "wos_file": "East China University of Science & Technology.xlsx",
        "treated": 1,
        "relo_year": 2007,
        "sample_role": "original_treat",
        "note": "奉贤校区；同城新建多校区",
    },
    {
        "school_cn": "东华大学",
        "wos_file": "Donghua University.xlsx",
        "treated": 1,
        "relo_year": 2003,
        "sample_role": "replacement_treat",
        "note": "替代南京大学；松江校区；同城新建多校区",
    },
    {
        "school_cn": "北京交通大学",
        "wos_file": "Beijing Jiaotong University.xlsx",
        "treated": 0,
        "relo_year": None,
        "sample_role": "original_control",
        "note": "威海校区为异地，不属于同城处理",
    },
    {
        "school_cn": "对外经济贸易大学",
        "wos_file": "University of International Business & Economics.xlsx",
        "treated": 0,
        "relo_year": None,
        "sample_role": "original_control",
        "note": "青岛国际校区为异地/建设信息，不属于同城处理",
    },
    {
        "school_cn": "北京工业大学",
        "wos_file": "Beijing University of Technology.xlsx",
        "treated": 0,
        "relo_year": None,
        "sample_role": "original_control",
        "note": "通州校区早于窗口期且本地表未列为同城处理",
    },
    {
        "school_cn": "中国传媒大学",
        "wos_file": "Communication University of China.xlsx",
        "treated": 0,
        "relo_year": None,
        "sample_role": "replacement_control",
        "note": "替代上海财经大学；本地搬迁表无多校区记录，官网概况列北京定福庄地址",
    },
    {
        "school_cn": "中国政法大学",
        "wos_file": "China University of Political Science & Law.xlsx",
        "treated": 0,
        "relo_year": None,
        "sample_role": "replacement_control",
        "note": "替代上海大学；昌平校区 1987 年，早于窗口期",
    },
]

EXCLUDED = [
    {
        "school_cn": "上海财经大学",
        "expected_wos_file": "Shanghai University of Finance & Economics.xlsx",
        "reason": "用户要求剔除；对照组在窗口期和之前均不得有多校区",
        "must_be_absent": 0,
    },
    {
        "school_cn": "南京大学",
        "expected_wos_file": "Nanjing University.xlsx",
        "reason": "当前 WOS 文件夹未找到对应文件，不能虚填",
        "must_be_absent": 1,
    },
    {
        "school_cn": "上海大学",
        "expected_wos_file": "Shanghai University.xlsx",
        "reason": "当前 WOS 文件夹未找到对应文件，不能虚填",
        "must_be_absent": 1,
    },
]


PAPER_FIELDS = [
    "school_id",
    "school_cn",
    "school_code",
    "wos_english_name",
    "treated",
    "relo_year",
    "post",
    "year",
    "person_id",
    "person_id_type",
    "author_full_name",
    "matched_orcid",
    "ut",
    "doi",
    "article_title",
    "source_title",
    "document_type",
    "publication_type",
    "cites_wos_core",
    "cites_all_db",
]


PANEL_FIELDS = [
    "school_id",
    "school_cn",
    "treated",
    "relo_year",
    "year",
    "post",
    "did",
    "person_id",
    "person_id_type",
    "author_full_name",
    "pub_count",
    "cites_wos_core_sum",
    "cites_all_db_sum",
]


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\r", " ").replace("\n", " ").strip()


def parse_int(value: object) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(str(value).strip()))
    except ValueError:
        return None


def parse_float(value: object) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(str(value).strip())
    except ValueError:
        return 0.0


def split_semicolon(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    return [part.strip() for part in text.split(";") if part.strip()]


def norm_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def parse_orcids(value: object) -> dict[str, str]:
    pairs: dict[str, str] = {}
    for part in split_semicolon(value):
        if "/" not in part:
            continue
        name, orcid = part.rsplit("/", 1)
        orcid = orcid.strip()
        if re.fullmatch(r"\d{4}-\d{4}-\d{4}-[\dX]{4}", orcid):
            pairs[norm_name(name)] = orcid
    return pairs


def require_columns(headers: list[str], required: list[str], file_name: str) -> dict[str, int]:
    index = {header: i for i, header in enumerate(headers)}
    missing = [col for col in required if col not in index]
    if missing:
        raise ValueError(f"{file_name} missing required columns: {missing}")
    return index


def validate_sample() -> None:
    missing_files = []
    for school in SCHOOLS:
        path = WOS_DIR / school["wos_file"]
        if not path.exists():
            missing_files.append(str(path))
    if missing_files:
        raise FileNotFoundError("Missing WOS files:\n" + "\n".join(missing_files))

    for excluded in EXCLUDED:
        path = WOS_DIR / excluded["expected_wos_file"]
        if int(excluded["must_be_absent"]) == 1 and path.exists():
            raise RuntimeError(f"Excluded file unexpectedly exists; review sample: {path}")

    treated_count = sum(int(s["treated"]) for s in SCHOOLS)
    control_count = len(SCHOOLS) - treated_count
    if treated_count != 5 or control_count != 5:
        raise RuntimeError(f"Expected 5:5 sample, got treated={treated_count}, control={control_count}")

    if (WOS_DIR / "University of Jinan.xlsx").exists() and not (WOS_DIR / "Jinan University.xlsx").exists():
        raise RuntimeError("Jinan University file is missing; do not use University of Jinan.")


def write_metadata() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    with (OUTPUT_DIR / "sample_school_mapping.csv").open("w", newline="", encoding="utf-8-sig") as f:
        fieldnames = ["school_id", "school_cn", "wos_file", "treated", "relo_year", "sample_role", "note"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for i, school in enumerate(SCHOOLS, start=1):
            writer.writerow({"school_id": i, **school})

    with (OUTPUT_DIR / "excluded_schools.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["school_cn", "expected_wos_file", "reason", "must_be_absent"])
        writer.writeheader()
        writer.writerows(EXCLUDED)


def main() -> None:
    validate_sample()
    write_metadata()

    paper_path = OUTPUT_DIR / "wos_author_paper_sample.csv"
    panel_path = OUTPUT_DIR / "wos_author_year_panel_sample.csv"
    validation_path = OUTPUT_DIR / "validation_summary.txt"

    aggregates: dict[tuple[int, str, int], dict[str, float]] = defaultdict(
        lambda: {"pub_count": 0.0, "cites_wos_core_sum": 0.0, "cites_all_db_sum": 0.0}
    )
    persons: dict[tuple[int, str], dict[str, object]] = {}
    school_stats: list[dict[str, object]] = []

    with paper_path.open("w", newline="", encoding="utf-8-sig") as f_paper:
        paper_writer = csv.DictWriter(f_paper, fieldnames=PAPER_FIELDS)
        paper_writer.writeheader()

        for school_id, school in enumerate(SCHOOLS, start=1):
            file_path = WOS_DIR / school["wos_file"]
            wb = load_workbook(file_path, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            headers = [clean_text(cell.value) for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            idx = require_columns(
                headers,
                [
                    "code",
                    "EnglishName",
                    "Publication Type",
                    "Author Full Names",
                    "Article Title",
                    "Source Title",
                    "Document Type",
                    "Publication Year",
                    "DOI",
                    "Times Cited, WoS Core",
                    "Times Cited, All Databases",
                    "ORCIDs",
                    "UT (Unique WOS ID)",
                ],
                school["wos_file"],
            )

            source_rows = 0
            in_window_rows = 0
            author_rows = 0
            orcid_matches = 0

            for row in ws.iter_rows(min_row=2, values_only=True):
                source_rows += 1
                year = parse_int(row[idx["Publication Year"]])
                if year is None or year < YEAR_MIN or year > YEAR_MAX:
                    continue

                authors = split_semicolon(row[idx["Author Full Names"]])
                if not authors:
                    continue

                in_window_rows += 1
                orcid_by_name = parse_orcids(row[idx["ORCIDs"]])
                cites_wos_core = parse_float(row[idx["Times Cited, WoS Core"]])
                cites_all_db = parse_float(row[idx["Times Cited, All Databases"]])
                relo_year = school["relo_year"]
                treated = int(school["treated"])
                post = int(treated == 1 and relo_year is not None and year >= int(relo_year))

                for author in authors:
                    normalized = norm_name(author)
                    matched_orcid = orcid_by_name.get(normalized, "")
                    if matched_orcid:
                        person_id = f"orcid:{matched_orcid}"
                        person_id_type = "orcid"
                        orcid_matches += 1
                    else:
                        person_id = f"school_name:{school_id}:{normalized}"
                        person_id_type = "school_name"

                    author_rows += 1
                    key = (school_id, person_id, year)
                    aggregates[key]["pub_count"] += 1
                    aggregates[key]["cites_wos_core_sum"] += cites_wos_core
                    aggregates[key]["cites_all_db_sum"] += cites_all_db
                    persons[(school_id, person_id)] = {
                        "school_id": school_id,
                        "school_cn": school["school_cn"],
                        "treated": treated,
                        "relo_year": "" if relo_year is None else relo_year,
                        "person_id": person_id,
                        "person_id_type": person_id_type,
                        "author_full_name": author,
                    }

                    paper_writer.writerow(
                        {
                            "school_id": school_id,
                            "school_cn": school["school_cn"],
                            "school_code": clean_text(row[idx["code"]]),
                            "wos_english_name": clean_text(row[idx["EnglishName"]]),
                            "treated": treated,
                            "relo_year": "" if relo_year is None else relo_year,
                            "post": post,
                            "year": year,
                            "person_id": person_id,
                            "person_id_type": person_id_type,
                            "author_full_name": author,
                            "matched_orcid": matched_orcid,
                            "ut": clean_text(row[idx["UT (Unique WOS ID)"]]),
                            "doi": clean_text(row[idx["DOI"]]),
                            "article_title": clean_text(row[idx["Article Title"]]),
                            "source_title": clean_text(row[idx["Source Title"]]),
                            "document_type": clean_text(row[idx["Document Type"]]),
                            "publication_type": clean_text(row[idx["Publication Type"]]),
                            "cites_wos_core": cites_wos_core,
                            "cites_all_db": cites_all_db,
                        }
                    )

            school_stats.append(
                {
                    "school_cn": school["school_cn"],
                    "wos_file": school["wos_file"],
                    "source_rows": source_rows,
                    "in_window_paper_rows": in_window_rows,
                    "author_paper_rows": author_rows,
                    "orcid_matches": orcid_matches,
                }
            )
            wb.close()

    with panel_path.open("w", newline="", encoding="utf-8-sig") as f_panel:
        panel_writer = csv.DictWriter(f_panel, fieldnames=PANEL_FIELDS)
        panel_writer.writeheader()

        for person_key in sorted(persons):
            person = persons[person_key]
            treated = int(person["treated"])
            relo_year_value = person["relo_year"]
            relo_year = int(relo_year_value) if relo_year_value != "" else None
            for year in range(YEAR_MIN, YEAR_MAX + 1):
                post = int(treated == 1 and relo_year is not None and year >= relo_year)
                did = treated * post
                vals = aggregates.get((int(person["school_id"]), str(person["person_id"]), year))
                panel_writer.writerow(
                    {
                        **person,
                        "year": year,
                        "post": post,
                        "did": did,
                        "pub_count": int(vals["pub_count"]) if vals else 0,
                        "cites_wos_core_sum": vals["cites_wos_core_sum"] if vals else 0,
                        "cites_all_db_sum": vals["cites_all_db_sum"] if vals else 0,
                    }
                )

    with validation_path.open("w", encoding="utf-8-sig") as f:
        f.write("Sample validation summary\n")
        f.write("=========================\n")
        f.write(f"Window: {YEAR_MIN}-{YEAR_MAX}\n")
        f.write("Sample balance: treated=5, control=5\n")
        f.write("Panel construction: balanced person-year rows for every observed person over 2000-2022.\n")
        f.write("Excluded schools: 南京大学 and 上海大学 because expected WOS files are absent; 上海财经大学 because it fails the strict control rule.\n")
        f.write("Control definition: no multicampus before or during the study window.\n")
        f.write("Shanghai University of Finance & Economics is excluded under the strict control rule.\n\n")
        for stat in school_stats:
            f.write(
                "{school_cn}\t{wos_file}\tsource_rows={source_rows}\t"
                "in_window_paper_rows={in_window_paper_rows}\tauthor_paper_rows={author_paper_rows}\t"
                "orcid_matches={orcid_matches}\n".format(**stat)
            )
        f.write(f"\nUnique persons: {len(persons)}\n")
        f.write(f"Person-year rows: {len(persons) * (YEAR_MAX - YEAR_MIN + 1)}\n")
        f.write(f"Author-paper CSV: {paper_path}\n")
        f.write(f"Person-year CSV: {panel_path}\n")

    print(f"Wrote {paper_path}")
    print(f"Wrote {panel_path}")
    print(f"Wrote {validation_path}")


if __name__ == "__main__":
    main()
