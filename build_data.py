# -*- coding: utf-8 -*-
"""
產生 data.js：醫師平均用量網頁的資料來源
來源檔案:
  權限表.xlsx          -> 登入權限 (google帳號 -> 可看的業代碼)
  客戶清單明細表2.xls   -> 醫院/醫師 (科別含「精神」, 職稱限 主治醫師/總醫師/部/科主任)
  GM-業績 (3).xlsx DATA -> Lote 100mg / Fute 5mg 2026年1~5月業績數量
計算:
  月平均 = 1~5月業績數量加總 / 5
  病患數 = ceil(月平均/1.5/28)  (兩產品各算後相加)
  醫師平均處方病人數 = 病患數 / 醫師數
"""
import json
import math
import sys
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

MONTHS = [202601, 202602, 202603, 202604, 202605]
TITLES = ["主治醫師", "總醫師", "部/科主任"]

# ---------- 權限表 ----------
perm = pd.read_excel("權限表.xlsx")
permissions = {}
for _, r in perm.iterrows():
    email = str(r["google登入帳號"]).strip().lower()
    scope = str(r["權限可看資料"]).strip()
    codes = "ALL" if scope == "全部" else [c.strip() for c in scope.split(",")]
    permissions[email] = {"name": str(r["呈現名稱"]).strip(), "codes": codes}

# ---------- 業績 DATA ----------
d = pd.read_excel("GM-業績 (3).xlsx", sheet_name="DATA")
rep_names = dict(d[["業代碼", "業代名稱"]].dropna().drop_duplicates().values)

d = d[d["年月"].isin(MONTHS) & d["業績數量"].notna()]
def product_key(name):
    s = str(name)
    if s.startswith("Lote F.C. 100mg"):
        return "lote"
    if s.startswith("Fute 5mg"):
        return "fute"
    return None
d = d.assign(prod=d["產品名稱"].map(product_key))
d = d[d["prod"].notna()]

qty = (
    d.groupby(["業代碼", "客戶名稱", "prod"])["業績數量"]
    .sum()
    .unstack("prod")
    .reindex(columns=["lote", "fute"])
    .fillna(0)
    .reset_index()
)

# ---------- 客戶清單 (精神科醫師) ----------
c = pd.read_excel("客戶清單明細表2.xls")
psy = c[
    c["科室名稱"].fillna("").str.contains("精神")
    & c["職稱"].isin(TITLES)
    # 名稱開頭為 XX/xx 的醫師不計算
    & ~c["聯絡人姓名"].fillna("").str.upper().str.startswith("XX")
]
# 醫師依 (員工姓名, 醫院) 分組
doctors_by = {}
for (emp, hosp), g in psy.groupby(["員工姓名", "客戶名稱"]):
    doctors_by[(emp, hosp)] = sorted(g["聯絡人姓名"].dropna().unique().tolist())

# ---------- 組表 (每個產品一列) ----------
PRODUCT_LABELS = {"lote": "Lote 100mg", "fute": "Fute 5mg"}

# 指定排除的醫院 (依業代碼，醫院名稱用開頭比對)
EXCLUDE = {
    "GM11": ["部玉里", "玉里榮"],
    "GM33": ["部台東", "台東榮", "台東聖母", "台東基督教"],
}

def is_excluded(code, hosp):
    return any(str(hosp).startswith(p) for p in EXCLUDE.get(code, []))

rows = []
for _, r in qty.iterrows():
    code = r["業代碼"]
    hosp = r["客戶名稱"]
    if is_excluded(code, hosp):
        continue
    rep = rep_names.get(code, "")
    docs = doctors_by.get((rep, hosp), [])
    for prod in ["lote", "fute"]:
        avg = r[prod] / len(MONTHS)
        if avg == 0:
            continue
        patients = math.ceil(avg / 1.5 / 28)
        rows.append({
            "code": code,
            "rep": rep,
            "hospital": hosp,
            "product": PRODUCT_LABELS[prod],
            "doctors": docs,
            "avg": round(avg, 1),
            "patients": patients,
            "perDoctor": round(patients / len(docs), 1) if docs else None,
        })

rows.sort(key=lambda x: (x["code"], x["hospital"], x["product"]))

data = {
    "generatedAt": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
    "months": "2026年1~5月",
    "permissions": permissions,
    "rows": rows,
}

with open("data.js", "w", encoding="utf-8") as f:
    f.write("const APP_DATA = ")
    f.write(json.dumps(data, ensure_ascii=False, indent=1))
    f.write(";\n")

# 將資料直接內嵌進 index.html (單一檔案即可部署)
import re
html = open("index.html", encoding="utf-8").read()
inline = ("/*APP_DATA_START*/\nwindow.APP_DATA = "
          + json.dumps(data, ensure_ascii=False) + ";\n/*APP_DATA_END*/")
html, n = re.subn(r"/\*APP_DATA_START\*/.*?/\*APP_DATA_END\*/", lambda m: inline, html, flags=re.S)
assert n == 1, "index.html 中找不到 APP_DATA 標記"
open("index.html", "w", encoding="utf-8").write(html)
print("index.html 已內嵌資料 (單一檔案可直接部署)")

# Apps Script 版資料檔 (伺服器端，使用者看不到)
import os
os.makedirs("apps-script", exist_ok=True)
with open("apps-script/Data.gs", "w", encoding="utf-8") as f:
    f.write("// 此檔由 build_data.py 自動產生，Excel 更新後重跑腳本並重新貼上\n")
    f.write("const PERMISSIONS = " + json.dumps(permissions, ensure_ascii=False, indent=1) + ";\n\n")
    f.write("const DATA_META = " + json.dumps(
        {"generatedAt": data["generatedAt"], "months": data["months"]},
        ensure_ascii=False) + ";\n\n")
    f.write("const DATA_ROWS = " + json.dumps(rows, ensure_ascii=False, indent=1) + ";\n")

print(f"rows: {len(rows)}")
for code in sorted({r['code'] for r in rows}):
    sub = [r for r in rows if r['code'] == code]
    hosps = {r['hospital'] for r in sub}
    nodoc = {r['hospital'] for r in sub if not r['doctors']}
    print(f"  {code} {rep_names.get(code,'')}: {len(sub)} 列 / {len(hosps)} 家醫院 (其中 {len(nodoc)} 家無符合條件醫師)")
