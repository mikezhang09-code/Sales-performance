"""基于 Mike-prepared.xlsx 构建区域销售评价得分并输出排名。

不依赖 openpyxl，直接解析 xlsx(xml)。
"""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile

XLSX_PATH = "Mike-prepared.xlsx"
NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def load_sheet_records(xlsx_path: str) -> list[dict[str, float | str]]:
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = []
        ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        for si in ss_root.findall("a:si", NS):
            txt = "".join((t.text or "") for t in si.findall(".//a:t", NS))
            shared_strings.append(txt)

        sheet_root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

    def cell_value(cell):
        cell_type = cell.attrib.get("t")
        v = cell.find("a:v", NS)
        if cell_type == "s" and v is not None:
            return shared_strings[int(v.text)]
        if v is not None:
            raw = v.text
            try:
                return float(raw)
            except (TypeError, ValueError):
                return raw
        inline = cell.find("a:is", NS)
        if inline is not None:
            return "".join((t.text or "") for t in inline.findall(".//a:t", NS))
        return None

    rows = sheet_root.findall(".//a:sheetData/a:row", NS)
    records: list[dict[str, float | str]] = []
    for row in rows[1:]:  # 跳过表头
        rec: dict[str, float | str] = {}
        for c in row.findall("a:c", NS):
            col = re.match(r"[A-Z]+", c.attrib["r"]).group(0)
            rec[col] = cell_value(c)
        if rec.get("A"):
            records.append(rec)
    return records


def minmax_norm(values: list[float]) -> dict[float, float]:
    mn, mx = min(values), max(values)
    if mx == mn:
        return {v: 0.0 for v in values}
    return {v: (v - mn) / (mx - mn) for v in values}


def main() -> None:
    data = load_sheet_records(XLSX_PATH)

    for r in data:
        r["penetration"] = 1 - float(r["F"])  # 客户渗透率（越高越好）
        r["y_per_mgr"] = float(r["Y"]) / float(r["B"]) if float(r["B"]) else 0.0
        t = float(r["T"])
        r["pilot_conv"] = float(r["U"]) / t if t else 0.0

    weights = {
        "P": 0.25,  # 净增收入完成率
        "O": 0.15,  # 收入增幅
        "I": 0.15,  # 人均产值
        "H": 0.10,  # 单位客户收入
        "penetration": 0.10,  # 客户渗透率
        "S": 0.10,  # 实际签约率
        "y_per_mgr": 0.10,  # 人均数字化签约金额
        "pilot_conv": 0.05,  # 试点企业转化率
    }

    norm_maps = {
        m: minmax_norm([float(r[m]) for r in data])
        for m in weights
    }

    for r in data:
        score = 0.0
        for m, w in weights.items():
            score += norm_maps[m][float(r[m])] * w
        r["score"] = score * 100

    ranked = sorted(data, key=lambda x: float(x["score"]), reverse=True)

    print("排名,区域,综合得分")
    for i, r in enumerate(ranked, start=1):
        print(f"{i},{r['A']},{float(r['score']):.2f}")


if __name__ == "__main__":
    main()
