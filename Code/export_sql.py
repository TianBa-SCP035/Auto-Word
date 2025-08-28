# export_sql.py
import pandas as pd
from sqlalchemy import create_engine, text
from pathlib import Path
from datetime import datetime
from urllib.parse import quote_plus  # ← 新增：安全编码密码

# --- DB1：项目/给药（原来这套） ---
DB_URL = (
    "mysql+pymysql://BioLIMS:y4kU*QrkCULGsR6S"
    "@rm-2ze3785rm2409mk1jfo.mysql.rds.aliyuncs.com:3306/"
    "xdida_platform_biocytogen?charset=utf8mb4"
)

# --- DB2：受试品信息（你新给的这套）---
# TODO: 把数据库名改成实际的（示例：'bbctg_sms'）
SMS_DB_NAME = "bbctg_sms"
_SMS_USER = "bbctgsmsview"
_SMS_PASS = quote_plus("BbctgsmsView2025#@")  # 编码 '#@'
DB_URL_SMS = (
    f"mysql+pymysql://{_SMS_USER}:{_SMS_PASS}@172.16.1.1:3306/"
    f"{SMS_DB_NAME}?charset=utf8mb4"
)

# ① 项目信息（按项目编号前缀）
SQL_PROJECT = """
SELECT
    p.id AS `项目ID`, 
    LEFT(p.snum, LENGTH(p.snum) - 2) AS `项目编号`,
    p.snum AS `实验编号`, 
    p.sname AS `项目名称`,
    CASE 
        WHEN ppe.project_purpose IS NULL 
             OR CHAR_LENGTH(TRIM(ppe.project_purpose)) < 3 
        THEN p.sname 
        ELSE ppe.project_purpose
    END AS `项目目的`,
    oe.sname AS `负责人`, 
    oe.email AS `负责人邮箱`,
    p.customer_name AS `客户名称`,
    up.entrust_people AS `委托单位负责人`,
    DATE_FORMAT(p.start_date, '%Y年%m月%d日') AS `开始日期`,
    DATE_FORMAT(p.end_date, '%Y年%m月%d日') AS `结束日期`,
    c.cell_names AS `细胞名称`,           
    an.sname AS `动物名称`, 
    st.sname AS `动物品系`, 
    CASE 
        WHEN pea.mouse_age IS NOT NULL AND pea.mouse_age <> 0 
            THEN pea.mouse_age
        ELSE CONCAT(pea.rat_age_low, '-', pea.rat_age_top, '周')
    END AS `鼠龄`,
    wt.sname AS `体重范围`, 
    CASE pea.sex
        WHEN 0 THEN '雌鼠'
        WHEN 1 THEN '雄鼠'
        ELSE '未知'
    END AS `性别`,
    pea.order_number AS `订购数量`, 
    ppe.group_number AS `入组数量`,
    ppe.groups AS `组数`,
    ppe.animal_number AS `每组数量`,
    ppe.group_condition AS `分组条件`
FROM project p
LEFT JOIN project_entry_effect_animal pea
    ON p.id = pea.project_id
LEFT JOIN org_tag AS an
    ON pea.animal_name = an.id
LEFT JOIN org_tag AS st
    ON pea.animal_strain = st.id
LEFT JOIN org_tag AS wt
    ON pea.weight_range = wt.id
LEFT JOIN united_project up
    ON p.snum LIKE CONCAT('%', up.snum, '%')
LEFT JOIN org_emp oe
    ON up.sd = oe.id
LEFT JOIN project_entry_pharmacological_effect ppe
    ON p.id = ppe.project_id
LEFT JOIN (
    SELECT
        pec.project_id,
        GROUP_CONCAT(DISTINCT oc.sname ORDER BY oc.sname SEPARATOR ',') AS cell_names
    FROM project_entry_effect_cell AS pec
    JOIN target_tag AS tt
        ON tt.project_effect_cell_id = pec.id
    JOIN org_tag AS oc
        ON oc.id = tt.tag_id
       AND oc.tag_type = 94
    GROUP BY pec.project_id
) AS c
    ON c.project_id = p.id
WHERE p.snum LIKE CONCAT(:project_code, '%');
"""



# ② 给药方案（按项目ID）
SQL_DOSE = """
SELECT
    project_entry_effect_drug.group_category AS `组别`, 
    project_entry_effect_drug.treatment_method AS `受试品`, 
    project_entry_effect_drug.animals_number AS `动物只数`, 
    project_entry_effect_drug.dose AS `剂量`, 
    route.sname AS `给药途径`, 
    freq.sname AS `给药频率`, 
    project_entry_effect_drug.dose_times AS `给药次数`
FROM project_entry_effect_drug
LEFT JOIN org_tag AS route
    ON project_entry_effect_drug.dose_mode LIKE CONCAT('[', route.id, ']')
LEFT JOIN org_tag AS freq
    ON project_entry_effect_drug.dose_frequency LIKE CONCAT('[', freq.id, ']')
WHERE project_entry_effect_drug.project_id = :project_id;
"""

# ③ 受试品信息（按 项目ID → LIKE 匹配 实验号）
SQL_SUPPLIES = """
SELECT
  /* —— 基本信息（统一做无效值归一化） —— */
  CASE
    WHEN rsp.NAME IS NULL
      OR TRIM(rsp.NAME) = ''
      OR TRIM(rsp.NAME) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.NAME
  END AS `名称`,
  CASE
    WHEN rsp.simplename IS NULL
      OR TRIM(rsp.simplename) = ''
      OR TRIM(rsp.simplename) IN ('0', 'NA', '/', '\\\\')
    THEN 
      CASE
        WHEN rsp.name IS NULL
          OR TRIM(rsp.name) = ''
          OR TRIM(rsp.name) IN ('0', 'NA', '/', '\\\\') THEN '-'
        ELSE rsp.name
      END
    ELSE rsp.simplename
  END AS `代号`,
  CASE
    WHEN rsp.suppliesType IS NULL
      OR TRIM(rsp.suppliesType) = ''
      OR TRIM(rsp.suppliesType) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.suppliesType
  END AS `来源`,
  CASE
    WHEN rsp.projectno IS NULL
      OR TRIM(rsp.projectno) = ''
      OR TRIM(rsp.projectno) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.projectno
  END AS `项目编号`,
  CASE
    WHEN rsp.testno IS NULL
      OR TRIM(rsp.testno) = ''
      OR TRIM(rsp.testno) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.testno
  END AS `实验号`,
  CASE
    WHEN rsp.properties IS NULL
      OR TRIM(rsp.properties) = ''
      OR TRIM(rsp.properties) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.properties
  END AS `性状`,
  /* 纯度：有效时加 %，否则 '-' */
  CASE
    WHEN rsp.purity IS NULL
      OR TRIM(rsp.purity) = ''
      OR TRIM(rsp.purity) IN ('0', 'NA', '/', '\\\\')
      OR rsp.purity = 0 THEN
      '-'
    ELSE
      CONCAT(rsp.purity, '%')
  END AS `纯度`,
  CASE
    WHEN rsp.lot_number IS NULL
      OR TRIM(rsp.lot_number) = ''
      OR TRIM(rsp.lot_number) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.lot_number
  END AS `批号`,
  CASE
    WHEN rsp.material_lot_number IS NULL
      OR TRIM(rsp.material_lot_number) = ''
      OR TRIM(rsp.material_lot_number) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.material_lot_number
  END AS `货号`,
  /* 规格：models/modelsunit 任一有效则拼接；都无效返回 '-' */
  CASE
    WHEN (
        rsp.models IS NULL
        OR TRIM(rsp.models) = ''
        OR TRIM(rsp.models) IN ('0', 'NA', '/', '\\\\')
      )
      AND (
        rsp.modelsunit IS NULL
        OR TRIM(rsp.modelsunit) = ''
        OR TRIM(rsp.modelsunit) IN ('0', 'NA', '/', '\\\\')
      ) THEN
      '-'
    ELSE
      CONCAT(
        CASE
          WHEN rsp.models IS NULL
            OR TRIM(rsp.models) = ''
            OR TRIM(rsp.models) IN ('0', 'NA', '/', '\\\\') THEN
            ''
          ELSE
            rsp.models
        END,
        CASE
          WHEN rsp.modelsunit IS NULL
            OR TRIM(rsp.modelsunit) = ''
            OR TRIM(rsp.modelsunit) IN ('0', 'NA', '/', '\\\\') THEN
            ''
          ELSE
            rsp.modelsunit
        END
      )
  END AS `规格`,
  /* 浓度：粉末 → '-'；否则拼 potency+unit（无效则 '-'） */
  CASE
    WHEN rsp.properties = '粉末' THEN
      '-'
    WHEN (
        rsp.potency IS NULL
        OR TRIM(rsp.potency) = ''
        OR TRIM(rsp.potency) IN ('0', 'NA', '/', '\\\\')
        OR rsp.potency = 0
      )
      AND (
        rsp.potencyunit IS NULL
        OR TRIM(rsp.potencyunit) = ''
        OR TRIM(rsp.potencyunit) IN ('0', 'NA', '/', '\\\\')
      ) THEN
      '-'
    WHEN rsp.potency IS NULL
      OR TRIM(rsp.potency) = ''
      OR TRIM(rsp.potency) IN ('0', 'NA', '/', '\\\\')
      OR rsp.potency = 0 THEN
      '-'
    ELSE
      CONCAT(rsp.potency, rsp.potencyunit)
  END AS `浓度`,
  /* —— 贮存条件：保存条件 +（可选）避光/干燥，用逗号拼；全无效时为 '-' —— */
  CASE
    WHEN (
        rsp.storagecondition IS NULL
        OR TRIM(rsp.storagecondition) = ''
        OR TRIM(rsp.storagecondition) IN ('0', 'NA', '/', '\\\\')
      )
      AND (
        rsp.issunblock IS NULL
        OR TRIM(rsp.issunblock) = ''
        OR TRIM(rsp.issunblock) IN ('0', '无要求', 'NA', '/', '\\\\')
      )
      AND (
        rsp.isdrystorage IS NULL
        OR TRIM(rsp.isdrystorage) = ''
        OR TRIM(rsp.isdrystorage) IN ('0', '无要求', 'NA', '/', '\\\\')
      ) THEN
      '-'
    ELSE
      CONCAT_WS(
        ', ',
        CASE
          WHEN rsp.storagecondition IS NULL
            OR TRIM(rsp.storagecondition) = ''
            OR TRIM(rsp.storagecondition) IN ('0', 'NA', '/', '\\\\') THEN
            NULL
          ELSE
            rsp.storagecondition
        END,
        CASE
          WHEN rsp.issunblock IS NULL
            OR TRIM(rsp.issunblock) = ''
            OR TRIM(rsp.issunblock) IN ('0', '无要求', 'NA', '/', '\\\\') THEN
            NULL
          ELSE
            '避光'
        END,
        CASE
          WHEN rsp.isdrystorage IS NULL
            OR TRIM(rsp.isdrystorage) = ''
            OR TRIM(rsp.isdrystorage) IN ('0', '无要求', 'NA', '/', '\\\\') THEN
            NULL
          ELSE
            '干燥'
        END
      )
  END AS `贮存条件`,
  CASE
    WHEN rsp.manufacturer IS NULL
      OR TRIM(rsp.manufacturer) = ''
      OR TRIM(rsp.manufacturer) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.manufacturer
  END AS `生产单位`,
  CASE
    WHEN rsp.customername IS NULL
      OR TRIM(rsp.customername) = ''
      OR TRIM(rsp.customername) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.customername
  END AS `供货单位`,
  CASE
    WHEN rsp.mfg IS NULL
      OR TRIM(rsp.mfg) = ''
      OR TRIM(rsp.mfg) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.mfg
  END AS `生产日期`,
  CASE
    WHEN rsp.validity IS NULL
      OR TRIM(rsp.validity) = ''
      OR TRIM(rsp.validity) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.validity
  END AS `有效期`,
  CASE
    WHEN rsp.sdname IS NULL
      OR TRIM(rsp.sdname) = ''
      OR TRIM(rsp.sdname) IN ('0', 'NA', '/', '\\\\') THEN
      '-'
    ELSE
      rsp.sdname
  END AS `SD`
FROM m_reagent_supplies rsp
WHERE
  rsp.testno   LIKE :full_like
  OR (
      rsp.projectno LIKE :prefix_like
      AND NOT EXISTS (
          SELECT 1 FROM m_reagent_supplies r2
          WHERE r2.testno LIKE :full_like
      )
  );
"""
##################################################################################################SQL区分隔

# —— 受试品信息合并：同名聚合、每列去重并用逗号连接 —— #
_NULLS = {"", "-", "NA", "/", "\\"}
def _normalize_cell(x) -> str:
    """把单元格转成可比较的字符串；无效值返回空串，便于后续过滤。"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip()
    return "" if s in _NULLS else s

def _uniq_join(values, sep=", "):
    """稳定去重并连接；如果全是无效值，返回 '-'。"""
    seen = set()
    out = []
    for v in values:
        v = _normalize_cell(v)
        if not v or v in seen:
            continue
        seen.add(v)
        out.append(v)
    return sep.join(out) if out else "-"

def _aggregate_supplies_by_name(df: pd.DataFrame) -> pd.DataFrame:
    """ 按“名称”聚合：同名的记录各列做去重合并。列保序输出；如果找不到“名称”则原样返回。 """
    if df is None or df.empty or "名称" not in df.columns:
        return df
    cols = list(df.columns)
    rows = []
    for name, g in df.groupby("名称", sort=False):
        row = {"名称": name}
        for col in cols:
            if col == "名称":
                continue
            row[col] = _uniq_join(g[col].tolist())
        rows.append(row)
    return pd.DataFrame(rows, columns=cols)

def _natural_sort_g(series: pd.Series) -> pd.Series:
    """G1,G2,...,G10 自然排序；无法提取数字的放最后。"""
    s = series.astype(str).str.extract(r'(\d+)')[0].astype(float)
    return s.fillna(1e9)

def _autosize(worksheet, dataframe: pd.DataFrame, max_rows: int = 20, min_width: int = 8, max_width: int = 20):
    """根据内容长度自动列宽（上下限可调）。"""
    if dataframe is None or dataframe.empty:
        return
    for i, col in enumerate(dataframe.columns):
        values = dataframe[col].astype(str).values[:max_rows]
        max_len = max([len(str(col))] + [len(x) for x in values])
        width = max(min_width, min(max_len + 2, max_width))
        worksheet.set_column(i, i, width)

###########################################################################################################工具函数分隔
def export_sql_to_excel(project_code: str, out_path: str):

    engine = create_engine(DB_URL, future=True)
    engine_sms = create_engine(DB_URL_SMS, future=True)  # ← 新增：受试品引擎

    with engine.connect() as conn:
        # 1) 项目信息
        df = pd.read_sql(text(SQL_PROJECT), conn, params={"project_code": project_code})

        # 2) 明细（纵表）
        note = ""
        if df.empty:
            note = "未查询到记录"
            df_vertical = pd.DataFrame(columns=["字段名", "字段值"])
            df_dose = pd.DataFrame(columns=["组别","受试品","动物只数","剂量","给药途径","给药频率","给药次数"])
            df_supplies = pd.DataFrame()  # 受试品信息（空）
        else:
            first_row = df.iloc[0].copy()
            # 把 1.0 这种整数小数转成 1
            for col in first_row.index:
                val = first_row[col]
                if isinstance(val, float) and val.is_integer():
                    first_row[col] = int(val)
            df_vertical = first_row.T.reset_index()
            df_vertical.columns = ["字段名", "字段值"]
            df_vertical["字段值"] = df_vertical["字段值"].fillna("").astype(str)

            # 3) 给药方案
            project_id = int(first_row["项目ID"])
            df_dose = pd.read_sql(text(SQL_DOSE), conn, params={"project_id": project_id})
            if not df_dose.empty and "组别" in df_dose.columns:
                df_dose = df_dose.sort_values(by="组别", key=_natural_sort_g)

            # 4) 受试品信息（DB2；优先用“上一个 SQL 的实验编号”，兜底用用户输入）
            proj_code = str(first_row.get("实验编号", "")).strip() or str(project_code).strip()

            full_like = f"%{proj_code}%"  # 用实验编号匹配实验号
            prefix_code = proj_code[:-2] if len(proj_code) > 2 else proj_code
            prefix_like = f"%{prefix_code}%"  # 若无，再用去尾两位匹配项目编号

            with engine_sms.connect() as conn2:
                df_supplies = pd.read_sql(
                    text(SQL_SUPPLIES),
                    conn2,
                    params={"full_like": full_like, "prefix_like": prefix_like}
                )

            # 按“名称”聚合：同名受试品的各列去重合并（浓度/规格不同会用逗号并列，相同只保留一个）
            df_supplies_agg = _aggregate_supplies_by_name(df_supplies)

            # 如果想同时保留“原始明细”，就两张表都写；否则用聚合结果覆盖
            keep_raw_supplies = False  # =True 时会额外写一张“受试品明细（原始）”

    out_path = Path(out_path)
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        # 全部数据
        df.to_excel(writer, index=False, sheet_name="全部数据")
        ws_all = writer.sheets["全部数据"]
        _autosize(ws_all, df)
        ws_all.freeze_panes(1, 0)

        # 明细
        df_vertical.to_excel(writer, index=False, sheet_name="明细")
        ws = writer.sheets["明细"]
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "text_wrap": True})
        for i, col in enumerate(df_vertical.columns):
            ws.write(0, i, col, header_fmt)
        _autosize(ws, df_vertical)
        ws.freeze_panes(1, 0)

        # 导出信息
        info = pd.DataFrame({
            "导出信息": ["导出时间", "项目编号", "总行数", "备注"],
            "结果": [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), project_code, len(df), note],
        })
        info.to_excel(writer, index=False, sheet_name="导出信息")
        ws_info = writer.sheets["导出信息"]
        _autosize(ws_info, info)
        ws_info.freeze_panes(1, 0)

        # 给药方案
        df_dose.to_excel(writer, index=False, sheet_name="给药方案")
        ws2 = writer.sheets["给药方案"]
        _autosize(ws2, df_dose)
        ws2.freeze_panes(1, 0)

        # 受试品信息（新）
        if keep_raw_supplies: # （可选）原始明细
            df_supplies.to_excel(writer, index=False, sheet_name="受试品明细（原始）")
            _autosize(writer.sheets["受试品明细（原始）"], df_supplies)
            writer.sheets["受试品明细（原始）"].freeze_panes(1, 0)

        df_supplies_agg.to_excel(writer, index=False, sheet_name="受试品信息")# 聚合后的“受试品信息”（供 Word 使用）
        ws3 = writer.sheets["受试品信息"]
        _autosize(ws3, df_supplies_agg)
        ws3.freeze_panes(1, 0)


    print(f"✅ 已导出 Excel：{out_path}")
    return out_path
