#!/usr/bin/env python3
"""Build the public SRCD evidence page from the local master workbook.

Only bibliography, evidence labels/findings, and stable row numbers are exported.
The translated abstracts in column D, private workbook contents, and local PDF paths
are never copied into the public page.
"""
from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

REPO_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = Path(__file__).resolve().parents[3]
DEFAULT_WORKBOOK = PROJECT_ROOT / "SRCD_文獻總表.xlsx"
TEMPLATE = REPO_ROOT / "site-src" / "index.template.html"
OUTPUT = REPO_ROOT / "public" / "index.html"
DOI_MAP_PATH = REPO_ROOT / "site-src" / "doi-map.json"
PUBLICATION_MAP_PATH = REPO_ROOT / "site-src" / "publication-map.json"
LOCAL_MIRROR = PROJECT_ROOT / "tmp" / "working" / "population-and-mechanisms" / "research-sequence-network.html"
SHEET_ID = "1BwHrlJhnYIRecen6l0tnWpZPVhIc7X_zW_3rRHFekZo"
SHEET_GID = "738173269"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

STUDY_LABELS = {
    1: "閱讀表現型 × 通用／閱讀情境 EF",
    2: "遊戲原型 × 歷程操弄確認",
    3: "介入 × 近遠遷移 × 中介",
}


def col_name(cell_ref: str) -> str:
    return re.sub(r"\d+$", "", cell_ref)


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.findall(f".//{{{NS_MAIN}}}t"))
            for item in root.findall(f"{{{NS_MAIN}}}si")]


def resolve_sheet_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(f"{{{NS_REL_PKG}}}Relationship")}
    for sheet in workbook.find(f"{{{NS_MAIN}}}sheets"):
        if sheet.attrib["name"] == sheet_name:
            target = rel_map[sheet.attrib[f"{{{NS_REL_DOC}}}id"]]
            if target.startswith("/"):
                return target.lstrip("/")
            return posixpath.normpath(posixpath.join("xl", target))
    raise ValueError(f"找不到工作表：{sheet_name}")


def read_sheet_rows(
    workbook_path: Path, sheet_name: str, columns: set[str] | None = None
) -> dict[int, dict[str, str]]:
    selected_columns = columns or {"A", "B", "C"}
    with zipfile.ZipFile(workbook_path) as archive:
        shared = read_shared_strings(archive)
        sheet_path = resolve_sheet_path(archive, sheet_name)
        root = ET.fromstring(archive.read(sheet_path))
    rows: dict[int, dict[str, str]] = {}
    for row_node in root.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row"):
        number = int(row_node.attrib["r"])
        values: dict[str, str] = {}
        for cell in row_node.findall(f"{{{NS_MAIN}}}c"):
            column = col_name(cell.attrib["r"])
            if column not in selected_columns:
                continue
            kind = cell.attrib.get("t", "")
            value = cell.find(f"{{{NS_MAIN}}}v")
            if kind == "s" and value is not None:
                text = shared[int(value.text or "0")]
            elif kind == "inlineStr":
                text = "".join(node.text or "" for node in cell.findall(f".//{{{NS_MAIN}}}t"))
            else:
                text = value.text if value is not None and value.text else ""
            values[column] = text
        rows[number] = values
    return rows


def author_surnames(author_text: str) -> list[str]:
    """Extract surnames from the master workbook's mixed author separators."""
    text = author_text.strip()
    if not text:
        return []
    # The workbook uses semicolons in newer rows, while older rows concatenate
    # names such as "Cartwright, Kelly B. Coppage, Elizabeth A.".
    parts = re.split(r";\s*|\s+(?=[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'’.-]+,\s)", text)
    surnames: list[str] = []
    for part in parts:
        match = re.match(r"\s*([^,]+),", part)
        surname = match.group(1).strip() if match else part.strip().split()[0]
        if surname and surname.lower() not in {value.lower() for value in surnames}:
            surnames.append(surname)
    return surnames


def apa_in_text(author_text: str, year: str, curated_surnames: list[str] | None = None) -> str:
    surnames = curated_surnames or author_surnames(author_text)
    if not surnames:
        return f"（{year}）" if year else "作者待核"
    if len(surnames) == 1:
        authors = surnames[0]
    elif len(surnames) == 2:
        # Narrative APA in-text citations spell out "and"; ampersands are for
        # parenthetical citations, which this linked line is not.
        authors = f"{surnames[0]} and {surnames[1]}"
    else:
        authors = f"{surnames[0]} et al."
    return f"{authors} ({year})" if year else authors


def readable_publication_name(name: str) -> str:
    """Convert WoS all-caps source titles to readable title case."""
    if not name:
        return ""
    words = name.lower().split()
    minor_words = {"a", "an", "and", "as", "at", "by", "for", "from", "in", "of", "on", "or", "the", "to", "with"}
    result = []
    for index, word in enumerate(words):
        if index and word in minor_words:
            result.append(word)
        else:
            result.append("-".join(piece[:1].upper() + piece[1:] for piece in word.split("-")))
    readable = " ".join(result)
    if readable.startswith("Jmir "):
        readable = "JMIR " + readable[5:]
    return readable


def classify(label: str, study: int) -> str:
    text = label.strip()
    if "理論背景" in text or "外部理論補充" in text:
        return "context"
    if "反例" in text or "反向" in text or "不一致" in text or "未獲支持" in text:
        return "not-supported"
    if "多研究背景" in text or "多項研究支持" in text:
        return "supported"
    return "partial"


def parse_records(
    rows: dict[int, dict[str, str]],
    doi_by_key: dict[str, str],
    publication_by_id: dict[str, dict[str, object]],
    bibliography_by_id: dict[str, dict[str, str]],
) -> list[dict[str, object]]:
    result: list[dict[str, object]] = []
    current_study: int | None = None
    for row_number in sorted(rows):
        cells = rows[row_number]
        heading = cells.get("A", "").strip()
        heading_match = re.match(r"研究([一二三])\s*[｜|]", heading)
        if heading_match:
            current_study = {"一": 1, "二": 2, "三": 3}[heading_match.group(1)]
        if current_study is None or row_number < 5:
            continue
        bibliographic = cells.get("B", "").strip()
        first_line, _, title = bibliographic.partition("\n")
        id_match = re.match(r"(SRCD-\d+|MIYAKE-\d{4}|WOS:\d+)\s*[｜|]\s*(.*)", first_line, re.S)
        if not id_match or not title.strip():
            continue
        literature_id, authors_year = id_match.groups()
        year_match = re.search(r"（(\d{4})）", authors_year)
        if not year_match:
            year_match = re.search(r"\b(19|20)\d{2}\b", authors_year)
            year = year_match.group(0) if year_match else ""
        else:
            year = year_match.group(1)
        authors = re.sub(r"\s*（\d{4}）\s*", "", authors_year).strip()
        bibliography = bibliography_by_id.get(literature_id, {})
        supplement = publication_by_id.get(literature_id, {})
        source_authors = bibliography.get("authors", authors)
        year = bibliography.get("year", year) or str(supplement.get("year", year))
        publication = bibliography.get("publication", "") or str(supplement.get("publication", ""))
        publication = readable_publication_name(publication)
        curated_surnames = supplement.get("surnames")
        label_parts = [part.strip() for part in heading.split("｜") if part.strip()]
        # Literature rows use column A for the scope and relationship label; section
        # headings carry a study prefix and no bibliographic ID, so they were skipped.
        if label_parts and label_parts[0].startswith("研究"):
            continue
        scope = label_parts[0] if label_parts else "補充來源"
        finding_label = "｜".join(label_parts[1:]) if len(label_parts) > 1 else ""
        detail = cells.get("C", "").strip()
        detail_lines = [line.strip() for line in detail.splitlines() if line.strip()]
        detail_lines = [line for line in detail_lines
                        if not line.startswith("判讀標籤：") and not line.startswith("本研究用途：")]
        finding = "\n".join(detail_lines)
        status = classify(heading, current_study)
        doi_suffix = doi_by_key.get(f"{current_study}:{literature_id}")
        summary_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit#gid={SHEET_GID}&range=D{row_number}"
        result.append({
            "row": row_number,
            "study": current_study,
            "studyLabel": STUDY_LABELS[current_study],
            "id": literature_id,
            "year": year,
            "authors": source_authors,
            "title": title.strip(),
            "shortCitation": f"{literature_id}（{year}）" if year else literature_id,
            "apaCitation": apa_in_text(source_authors, year, curated_surnames if isinstance(curated_surnames, list) else None),
            "publication": publication,
            "indexScope": scope,
            "findingLabel": finding_label,
            "finding": finding,
            "status": status,
            "summaryUrl": summary_url,
            "doi": f"https://doi.org/{doi_suffix}" if doi_suffix else "",
        })
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK,
                        help="本機主本 SRCD_文獻總表.xlsx 路徑")
    args = parser.parse_args()
    workbook_path = args.workbook.expanduser().resolve()
    if not workbook_path.exists():
        print(f"找不到本機主本：{workbook_path}", file=sys.stderr)
        return 2
    if not TEMPLATE.exists():
        print(f"找不到網站樣板：{TEMPLATE}", file=sys.stderr)
        return 2
    if not DOI_MAP_PATH.exists():
        print(f"找不到 DOI 對照表：{DOI_MAP_PATH}", file=sys.stderr)
        return 2

    if not PUBLICATION_MAP_PATH.exists():
        print(f"找不到書目出版資訊對照表：{PUBLICATION_MAP_PATH}", file=sys.stderr)
        return 2

    doi_by_key = json.loads(DOI_MAP_PATH.read_text(encoding="utf-8"))
    publication_by_id = json.loads(PUBLICATION_MAP_PATH.read_text(encoding="utf-8"))
    bibliography_rows = read_sheet_rows(workbook_path, "文獻總覽", {"C", "H", "J", "L"})
    bibliography_by_id = {
        cells["H"].strip(): {
            "authors": cells.get("C", "").strip(),
            "year": cells.get("J", "").strip(),
            "publication": cells.get("L", "").strip(),
        }
        for cells in bibliography_rows.values()
        if cells.get("H", "").strip()
    }
    records = parse_records(
        read_sheet_rows(workbook_path, "三研究導覽"), doi_by_key,
        publication_by_id, bibliography_by_id,
    )
    if not records:
        print("未從「三研究導覽」讀到文獻列，未輸出網站。", file=sys.stderr)
        return 2
    for study in (1, 2, 3):
        if not any(record["study"] == study for record in records):
            print(f"研究{study} 沒有文獻列，請檢查活頁簿分段標題。", file=sys.stderr)
            return 2

    data = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    data = data.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    page = TEMPLATE.read_text(encoding="utf-8").replace("__REFERENCE_DATA__", data)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(page, encoding="utf-8")
    if LOCAL_MIRROR.parent.exists():
        LOCAL_MIRROR.write_text(page, encoding="utf-8")

    print(f"已從 {workbook_path.name} 產生 {len(records)} 筆關係紀錄。")
    print(f"研究一／二／三：{', '.join(str(sum(1 for r in records if r['study'] == study)) for study in (1,2,3))} 筆")
    missing_publication = [record["id"] for record in records if not record["publication"]]
    missing_doi = sorted({record["id"] for record in records if not record["doi"]})
    if missing_publication:
        print(f"缺少期刊／書名的文獻 ID：{', '.join(sorted(set(missing_publication)))}", file=sys.stderr)
    if missing_doi:
        print(f"缺少 DOI／文章入口的文獻 ID：{', '.join(missing_doi)}", file=sys.stderr)
    print(f"GitHub Pages 檔案：{OUTPUT}")
    print(f"本機圖表鏡像：{LOCAL_MIRROR if LOCAL_MIRROR.parent.exists() else '未更新（找不到本機 tmp 路徑）'}")
    print("匯出欄位不含繁中摘要欄 D、私有活頁簿內容或本機全文路徑。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
