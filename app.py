# -*- coding: utf-8 -*-
"""
智策经营——AI 驱动的多维经营分析与决策支持系统
高级视觉版

重点：
1. 不再使用旧版 plotly express 的高风险函数，避免 pandas/plotly 兼容问题；
2. 数值列采用稳健解析，保留 Excel 中真实数值；
3. 图表绘制前先检查有效数据，不再出现空白图却没有说明的情况；
4. 有日期看趋势，无日期看结构；
5. 异常诊断解释“为什么异常”和“为什么高风险”；
6. 问数示例根据当前上传数据动态生成。
"""

import os
import re
import time
import json
import sqlite3
from io import BytesIO

import numpy as np
import pandas as pd
import requests
import streamlit as st
import plotly.graph_objects as go
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False


# ============================================================
# 0. 基础设置
# ============================================================

st.set_page_config(
    page_title="智策经营",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
:root {
    --zc-primary: #123A63;
    --zc-primary-2: #1E77D3;
    --zc-bg: #F3F6FB;
    --zc-card: #FFFFFF;
    --zc-text: #10213F;
    --zc-muted: #66758C;
    --zc-line: #E5EBF3;
    --zc-shadow: 0 14px 34px rgba(24, 48, 88, .08);
}

.stApp { background: radial-gradient(circle at top left, #F8FBFF 0, #F3F6FB 42%, #EEF3FA 100%); }
.block-container { padding-top: 1.1rem; padding-bottom: 2.6rem; max-width: 1680px; }
html, body, [class*="css"] { font-size: 17px; }
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #EEF4FB 0%, #E8EEF6 100%);
    border-right: 1px solid #DDE6F1;
    min-width: 350px !important;
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: var(--zc-text);
}
section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p,
section[data-testid="stSidebar"] li {
    font-size: 16px;
    line-height: 1.7;
}
label[data-testid="stWidgetLabel"] {
    font-size: 16px !important;
    font-weight: 800 !important;
    color: #243654 !important;
}
.stSelectbox div[data-baseweb="select"],
.stMultiSelect div[data-baseweb="select"],
.stTextInput input {
    min-height: 48px;
    border-radius: 14px !important;
    font-size: 16px !important;
}
.stButton button {
    border-radius: 14px !important;
    min-height: 46px;
    font-size: 17px !important;
    font-weight: 800 !important;
    padding: 0 24px !important;
}
.stDownloadButton button {
    border-radius: 14px !important;
    min-height: 46px;
    font-size: 17px !important;
    font-weight: 800 !important;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
    border-bottom: 1px solid #DDE6F1;
    padding-bottom: 6px;
}
.stTabs [data-baseweb="tab"] {
    font-size: 18px !important;
    font-weight: 900 !important;
    padding: 14px 20px !important;
    border-radius: 14px 14px 0 0;
}
.stTabs [aria-selected="true"] {
    background: #FFFFFF !important;
    box-shadow: 0 -2px 14px rgba(24, 48, 88, .05);
}
h1 { font-size: 40px !important; font-weight: 950 !important; color: var(--zc-text); }
h2 { font-size: 34px !important; font-weight: 950 !important; color: var(--zc-text); margin-top: 18px !important; }
h3 { font-size: 25px !important; font-weight: 900 !important; color: var(--zc-text); margin-top: 16px !important; }
div[data-testid="stMarkdownContainer"] p,
div[data-testid="stMarkdownContainer"] li {
    font-size: 17px;
    line-height: 1.85;
}
.hero {
    background:
      linear-gradient(135deg, rgba(16, 52, 92, .98), rgba(30, 119, 211, .94)),
      radial-gradient(circle at 92% 8%, rgba(255,255,255,.22), transparent 26%);
    border-radius: 28px;
    padding: 34px 38px;
    color: white;
    box-shadow: 0 22px 42px rgba(15, 61, 99, .20);
    margin-bottom: 24px;
    border: 1px solid rgba(255,255,255,.22);
}
.hero h1 { font-size: 42px !important; line-height: 1.22; margin: 0; font-weight: 950; color: white !important; }
.hero p { font-size: 18px; opacity: .96; margin-top: 12px; max-width: 980px; }
.metric-card {
    background: linear-gradient(180deg, #FFFFFF 0%, #FBFDFF 100%);
    border: 1px solid var(--zc-line);
    border-radius: 24px;
    box-shadow: var(--zc-shadow);
    padding: 24px 26px;
    min-height: 158px;
    position: relative;
    overflow: hidden;
}
.metric-card:before {
    content: "";
    position: absolute;
    top: 0;
    left: 0;
    height: 5px;
    width: 100%;
    background: linear-gradient(90deg, #1E77D3, #78B7FF);
}
.metric-title { color: var(--zc-muted); font-size: 16px; font-weight: 800; margin-bottom: 12px; }
.metric-value { color: #071F43; font-size: 38px; font-weight: 950; letter-spacing: .2px; line-height: 1.1; }
.metric-note { color: #8290A5; font-size: 14px; margin-top: 12px; line-height: 1.55; }
.section-card {
    background: var(--zc-card);
    border: 1px solid var(--zc-line);
    border-radius: 24px;
    box-shadow: var(--zc-shadow);
    padding: 26px 28px;
    margin-bottom: 20px;
}
.tip-card {
    background: linear-gradient(90deg, #EAF4FF 0%, #F6FBFF 100%);
    border-left: 7px solid #2F80ED;
    padding: 18px 22px;
    border-radius: 18px;
    color: #0B3A6A;
    margin-bottom: 18px;
    line-height: 1.85;
    font-size: 17px;
    box-shadow: 0 8px 18px rgba(47, 128, 237, .08);
}
.warn-card {
    background: linear-gradient(90deg, #FFF7E1 0%, #FFFDF5 100%);
    border-left: 7px solid #F2A900;
    padding: 18px 22px;
    border-radius: 18px;
    color: #704E00;
    margin-bottom: 18px;
    line-height: 1.85;
    font-size: 17px;
    box-shadow: 0 8px 18px rgba(242, 169, 0, .10);
}
.risk-card {
    background: #FFFFFF;
    border: 1px solid var(--zc-line);
    border-radius: 22px;
    padding: 20px 22px;
    min-height: 190px;
    box-shadow: var(--zc-shadow);
}
.example-card {
    background: #FFFFFF;
    border: 1px solid var(--zc-line);
    border-radius: 18px;
    padding: 17px 18px;
    min-height: 118px;
    box-shadow: 0 8px 20px rgba(30,55,90,.06);
}
.example-title { font-size: 15px; color: #60718A; font-weight: 900; margin-bottom: 8px; }
.example-q { font-size: 17px; color: #0B2142; line-height: 1.65; font-weight: 650; }
.small-text { color: #5E6F86; font-size: 16px; line-height: 1.85; }
[data-testid="stDataFrame"] {
    border-radius: 18px;
    overflow: hidden;
    border: 1px solid #E3EAF4;
    box-shadow: 0 8px 20px rgba(30,55,90,.05);
}
.js-plotly-plot {
    border-radius: 22px;
    overflow: hidden;
    border: 1px solid #E5EBF3;
    box-shadow: 0 10px 24px rgba(30,55,90,.06);
    background: #FFFFFF;
}
hr { border-color: #DDE6F1; }

.field-card {
    background: linear-gradient(180deg, #FFFFFF 0%, #FBFDFF 100%);
    border: 1px solid #E5EBF3;
    border-radius: 22px;
    padding: 20px 22px;
    min-height: 178px;
    box-shadow: 0 10px 24px rgba(30,55,90,.06);
}
.field-card-title {
    font-size: 19px;
    font-weight: 950;
    color: #10213F;
    margin-bottom: 10px;
}
.field-card-desc {
    font-size: 15px;
    color: #66758C;
    line-height: 1.65;
    margin-bottom: 12px;
}
.pill-wrap { display: flex; flex-wrap: wrap; gap: 8px; }
.field-pill {
    display: inline-block;
    background: #EEF5FF;
    color: #164B80;
    border: 1px solid #D6E8FF;
    border-radius: 999px;
    padding: 7px 12px;
    font-size: 15px;
    font-weight: 800;
}
.field-pill.warn {
    background: #FFF7E6;
    color: #8A5A00;
    border-color: #FFE2A8;
}
.field-pill.muted {
    background: #F2F5F9;
    color: #66758C;
    border-color: #E1E7F0;
}
.health-summary {
    background: linear-gradient(135deg, #FFFFFF 0%, #F7FBFF 100%);
    border: 1px solid #E5EBF3;
    border-radius: 24px;
    padding: 24px 28px;
    box-shadow: 0 12px 28px rgba(30,55,90,.07);
    min-height: 300px;
}
.health-summary h3 {
    margin-top: 0 !important;
    margin-bottom: 12px !important;
}
.health-line {
    font-size: 17px;
    line-height: 1.85;
    color: #40516A;
}
.progress-row {
    margin: 12px 0 16px 0;
}
.progress-label {
    display: flex;
    justify-content: space-between;
    font-size: 15px;
    font-weight: 850;
    color: #243654;
    margin-bottom: 6px;
}
.progress-track {
    width: 100%;
    height: 14px;
    background: #E9EEF6;
    border-radius: 999px;
    overflow: hidden;
}
.progress-fill {
    height: 14px;
    border-radius: 999px;
    background: linear-gradient(90deg, #1E77D3, #71B7FF);
}
.detail-expander-note {
    color: #66758C;
    font-size: 15px;
    line-height: 1.7;
    margin-bottom: 10px;
}


/* 汇报增强补充 */
div[data-testid="stMarkdownContainer"] p,
div[data-testid="stMarkdownContainer"] li {
    font-size: 18px !important;
    line-height: 1.95 !important;
}
.small-text { font-size: 18px !important; line-height: 1.95 !important; }
.detail-expander-note { font-size: 17px !important; line-height: 1.85 !important; }
.highlight-note {
    background: linear-gradient(90deg, #F4F8FF 0%, #FFFFFF 100%);
    border-left: 6px solid #1E77D3;
    border-radius: 16px;
    padding: 15px 18px;
    font-size: 18px;
    color: #334766;
    line-height: 1.9;
    margin: 12px 0 16px 0;
}
.ai-table-card {
    background: #FFFFFF;
    border: 1px solid #E5EBF3;
    border-radius: 18px;
    padding: 18px 20px;
    box-shadow: 0 8px 22px rgba(30,55,90,.06);
    margin: 12px 0 18px 0;
}
.ai-section-title {
    font-size: 21px;
    font-weight: 950;
    color: #10213F;
    margin: 10px 0 8px 0;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# 1. 数据读取与字段识别
# ============================================================

def normalize_column_name(col):
    col = str(col).strip()
    col = re.sub(r"\s+", "_", col)
    col = re.sub(r"[^\w\u4e00-\u9fff]", "_", col)
    col = re.sub(r"_+", "_", col)
    return col.strip("_")


def read_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file, encoding="utf-8")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="gbk")
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    raise ValueError("暂不支持该文件类型，请上传 xlsx、xls 或 csv。")


def is_id_like(col):
    c = str(col).lower()
    patterns = [
        r"(^|_)id($|_)", r"编号", r"序号", r"代码", r"code", r"key",
        r"index", r"row", r"流水", r"唯一", r"主键", r"no$"
    ]
    return any(re.search(p, c) for p in patterns)


def is_rate_like(col):
    c = str(col).lower()
    return any(k in c for k in ["rate", "ratio", "percent", "percentage", "pct", "margin", "discount", "折扣", "率", "占比", "比例", "完成率"])


def is_amount_like(col):
    c = str(col).lower()
    return any(k in c for k in [
        "amount", "cost", "expense", "spend", "revenue", "sales", "profit",
        "pay", "price", "fee", "total", "asset", "liab", "equity", "cash",
        "收入", "销售", "成本", "费用", "支出", "利润", "薪酬", "工资", "金额", "单价", "总额", "预算", "资产", "负债"
    ])


def is_quantity_like(col):
    c = str(col).lower()
    return any(k in c for k in ["qty", "quantity", "count", "number", "hours", "usage", "数量", "工时", "用量", "次数", "人数"])


def is_date_name_like(col):
    c = str(col).lower()
    return any(k in c for k in ["date", "time", "month", "year", "period", "日期", "时间", "月份", "年月", "期间"])


def is_progress_like(col):
    c = str(col).lower()
    return any(k in c for k in ["day_of", "days", "duration", "进度", "天数", "周期", "耗时", "时长"])


def parse_numeric_series(s, aggressive=False):
    """稳健数值解析，支持 Excel 数值、逗号、货币符号、万/亿、百分号、Day 12 等。"""
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    raw = s.astype(str).str.strip()
    raw = raw.replace({
        "": np.nan, "nan": np.nan, "NaN": np.nan, "None": np.nan,
        "null": np.nan, "NULL": np.nan, "-": np.nan, "—": np.nan
    })

    neg_mask = raw.str.match(r"^\(.*\)$", na=False)
    cleaned = raw.str.replace(r"^\((.*)\)$", r"\1", regex=True)
    cleaned = cleaned.str.replace(r"[,，\s]", "", regex=True)
    cleaned = cleaned.str.replace(r"(人民币|RMB|CNY|USD|HKD|￥|¥|\$|元)", "", regex=True)

    percent_mask = cleaned.str.contains(r"%|％", na=False)

    multiplier = pd.Series(1.0, index=s.index)
    multiplier = multiplier.mask(cleaned.str.contains("亿", na=False), 100000000.0)
    multiplier = multiplier.mask(cleaned.str.contains("万", na=False), 10000.0)
    multiplier = multiplier.mask(cleaned.str.contains(r"[kK]$", na=False), 1000.0)
    multiplier = multiplier.mask(cleaned.str.contains(r"[mM]$", na=False), 1000000.0)

    cleaned = cleaned.str.replace(r"(亿|万|%|％|[kKmM])", "", regex=True)
    parsed = pd.to_numeric(cleaned, errors="coerce")

    if aggressive:
        extracted = cleaned.str.extract(r"([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)")[0]
        extracted = pd.to_numeric(extracted, errors="coerce")
        parsed = parsed.fillna(extracted)

    parsed = parsed * multiplier
    parsed = parsed.mask(percent_mask, parsed / 100.0)
    parsed = parsed.mask(neg_mask, -parsed)
    return parsed


def try_parse_date(series):
    parsed = pd.to_datetime(series, errors="coerce")
    valid_ratio = parsed.notna().mean()
    if valid_ratio == 0:
        return parsed, 0, False, "无法解析为日期"

    years = parsed.dropna().dt.year
    unique_dates = parsed.dropna().nunique()
    suspicious_epoch = (years.between(1969, 1971).mean() > 0.8) if len(years) else False
    too_few_unique = unique_dates < 2
    valid = valid_ratio >= 0.55 and not suspicious_epoch and not too_few_unique

    if valid_ratio < 0.55:
        reason = "日期解析比例较低"
    elif suspicious_epoch:
        reason = "疑似数值被误解析为1970附近日期"
    elif too_few_unique:
        reason = "有效日期取值过少"
    else:
        reason = "有效日期字段"
    return parsed, valid_ratio, valid, reason


def infer_field_metadata(raw_df):
    rows = []
    for col in raw_df.columns:
        s = raw_df[col]
        sample_values = s.dropna().astype(str).head(3).tolist()
        numeric = parse_numeric_series(s, aggressive=False)
        nr = numeric.notna().mean()
        uniq = int(s.nunique(dropna=True))
        ur = uniq / max(len(s), 1)

        if is_date_name_like(col):
            _, dr, valid_date, date_reason = try_parse_date(s)
        else:
            dr, valid_date, date_reason = 0, False, ""

        if is_id_like(col):
            role, ftype, agg, ok = "标识字段", "ID/编号", "去重计数 / 不聚合", "否"
            meaning = "用于唯一标识记录、对象或业务单元，主要用于追踪和去重，不应作为可加总的经营指标。"
            note = "不建议用于相关性分析或趋势分析。"
        elif valid_date:
            role, ftype, agg, ok = "时间字段", "日期", "按年/月/日分组", "否"
            meaning = "表示业务发生或统计所属的时间，可用于趋势、环比、同比和周期变化分析。"
            note = date_reason
        elif nr > 0.75 and is_progress_like(col):
            role, ftype, agg, ok = "进度/时长字段", "数值", "均值 / 分段分析", "可作为辅助指标"
            meaning = "表示项目进行天数、周期、进度或耗时，适合做进度型分析，不应直接当作自然日期。"
            note = "可作为数值横轴或辅助指标。"
        elif nr > 0.75 and is_rate_like(col):
            role, ftype, agg, ok = "度量指标", "比例/比率", "平均 / 加权平均", "谨慎"
            meaning = "表示比例、比率、折扣或百分比类指标，适合比较水平，不适合直接求和。"
            note = "默认建议使用平均或加权平均；折扣字段不能求和。"
        elif nr > 0.75 and is_amount_like(col):
            role, ftype, agg, ok = "度量指标", "金额/成本/收入类", "求和 / 均值", "是"
            meaning = "表示金额、成本、费用、收入、利润或预算等经营度量，可作为主分析指标。"
            note = "适合做总量、均值、排名、结构和异常分析。"
        elif nr > 0.75 and is_quantity_like(col):
            role, ftype, agg, ok = "度量指标", "数量/工时/用量类", "求和 / 均值", "是"
            meaning = "表示数量、工时、用量或次数等投入/产出规模指标，可用于规模、效率或驱动关系分析。"
            note = "可与金额类指标做关系分析。"
        elif nr > 0.75:
            role, ftype, agg, ok = "数值字段", "数值", "求和 / 均值 / 中位数", "可选"
            meaning = "数值型字段，可根据业务含义作为主指标或辅助指标分析。"
            note = "建议结合字段含义确认是否可加总。"
        elif ur < 0.65:
            role, ftype, agg, ok = "维度字段", "类别/分组", "分组统计", "否"
            meaning = "表示类别、部门、地区、状态、类型或对象分组，可作为经营洞察口径。"
            note = "适合用于分组对比、贡献度分析和结构分析。"
        else:
            role, ftype, agg, ok = "文本字段", "文本", "不聚合 / 计数", "否"
            meaning = "文本描述类字段，通常用于展示或辅助检索，不适合作为数值分析指标。"
            note = "如需分析文本内容，可后续接入文本挖掘。"

        rows.append({
            "字段名": col,
            "字段含义解释": meaning,
            "字段类型": ftype,
            "推荐角色": role,
            "推荐聚合方式": agg,
            "是否适合作为主指标": ok,
            "唯一值数量": uniq,
            "数值解析比例": round(float(nr), 3),
            "日期解析比例": round(float(dr), 3),
            "注意事项": note,
            "示例值": "、".join(sample_values)
        })
    return pd.DataFrame(rows)


def get_metric_candidates(meta):
    temp = meta[meta["推荐角色"].isin(["度量指标", "数值字段", "进度/时长字段"])]
    return temp["字段名"].tolist()


def get_dimension_candidates(meta):
    return meta[meta["推荐角色"].isin(["维度字段"])]["字段名"].tolist()


def get_date_candidates(meta):
    return meta[meta["推荐角色"].isin(["时间字段"])]["字段名"].tolist()


def convert_selected_numeric(df, cols):
    out = df.copy()
    report = []
    for col in cols:
        if col in out.columns:
            before_valid = out[col].notna().sum()
            out[col] = parse_numeric_series(out[col], aggressive=True)
            after_valid = out[col].notna().sum()
            report.append({
                "字段": col,
                "转换前非空数": int(before_valid),
                "转换后有效数值数": int(after_valid),
                "缺失/无法解析数": int(out[col].isna().sum())
            })
    return out, pd.DataFrame(report)


def clean_data(raw_df, selected_numeric_cols, date_col, missing_strategy):
    clean = raw_df.copy()
    raw_rows = len(clean)
    dup = int(clean.duplicated().sum())
    clean = clean.drop_duplicates()

    reports = [{"处理环节": "删除完全重复行", "处理数量": dup, "说明": "整行完全一致时视为重复记录。"}]

    if date_col and date_col in clean.columns:
        parsed, ratio, valid, reason = try_parse_date(clean[date_col])
        clean[date_col] = parsed
        reports.append({"处理环节": "日期字段解析", "处理数量": int(parsed.isna().sum()), "说明": f"{date_col}：{reason}，解析成功比例 {ratio:.1%}。"})

    clean, numeric_report = convert_selected_numeric(clean, selected_numeric_cols)
    for _, r in numeric_report.iterrows():
        reports.append({"处理环节": f"数值转换：{r['字段']}", "处理数量": int(r["缺失/无法解析数"]), "说明": f"转换后有效数值 {int(r['转换后有效数值数'])} 条。"})

    if missing_strategy == "删除主指标缺失行" and selected_numeric_cols:
        main = selected_numeric_cols[0]
        before = len(clean)
        clean = clean.dropna(subset=[main])
        reports.append({"处理环节": "缺失值处理", "处理数量": before - len(clean), "说明": f"删除主指标 {main} 缺失的记录。"})
    elif missing_strategy == "数值中位数填充，类别填未知":
        fill_count = 0
        for col in clean.columns:
            na = int(clean[col].isna().sum())
            if na == 0:
                continue
            if pd.api.types.is_numeric_dtype(clean[col]):
                clean[col] = clean[col].fillna(clean[col].median())
            else:
                clean[col] = clean[col].fillna("未知")
            fill_count += na
        reports.append({"处理环节": "缺失值处理", "处理数量": fill_count, "说明": "数值列用中位数填充，类别/文本列填未知。"})
    else:
        reports.append({"处理环节": "缺失值处理", "处理数量": 0, "说明": "保留缺失值，在具体分析时自动忽略缺失。"})

    summary = {
        "原始行数": raw_rows,
        "清洗后行数": len(clean),
        "原始字段数": len(raw_df.columns),
        "清洗后字段数": len(clean.columns),
        "重复行数": dup,
        "缺失单元格数": int(clean.isna().sum().sum())
    }
    return clean, pd.DataFrame(reports), summary, numeric_report


# ============================================================
# 2. 图表与分析函数
# ============================================================

def money_fmt(x):
    if pd.isna(x):
        return "-"
    try:
        x = float(x)
    except Exception:
        return str(x)
    if abs(x) >= 100000000:
        return f"{x / 100000000:.2f}亿"
    if abs(x) >= 10000:
        return f"{x / 10000:.2f}万"
    return f"{x:,.2f}"


def pct_fmt(x):
    if pd.isna(x):
        return "-"
    return f"{float(x):.2%}"


def kpi_card(title, value, note=""):
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-note">{note}</div>
    </div>
    """, unsafe_allow_html=True)


def chart_layout(fig, height=420):
    height = max(height, 500)
    fig.update_layout(
        height=height,
        template="plotly_white",
        margin=dict(l=34, r=34, t=76, b=56),
        font=dict(size=15, family="Microsoft YaHei, PingFang SC, Arial"),
        title=dict(font=dict(size=23, family="Microsoft YaHei", color="#10213F"), x=0.02, xanchor="left"),
        legend=dict(font=dict(size=14), bgcolor="rgba(255,255,255,0.72)", bordercolor="#E5EBF3", borderwidth=1),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        hoverlabel=dict(font_size=15, bgcolor="#FFFFFF", bordercolor="#DDE6F1")
    )
    fig.update_xaxes(
        title_font=dict(size=17, color="#53657E"),
        tickfont=dict(size=14, color="#66758C"),
        gridcolor="#E8EEF6",
        zerolinecolor="#DDE6F1",
        showline=True,
        linecolor="#DDE6F1"
    )
    fig.update_yaxes(
        title_font=dict(size=17, color="#53657E"),
        tickfont=dict(size=14, color="#66758C"),
        gridcolor="#E8EEF6",
        zerolinecolor="#DDE6F1",
        showline=True,
        linecolor="#DDE6F1"
    )
    return fig


def to_finite_numeric_series(df, col, aggressive=True):
    """把任意字段稳健转换为可绘图的有限数值序列，避免 Plotly 拿到 object/Inf/空序列后画空图。"""
    if col not in df.columns:
        return pd.Series(dtype="float64")
    s = df[col]
    if pd.api.types.is_numeric_dtype(s):
        out = pd.to_numeric(s, errors="coerce")
    else:
        out = parse_numeric_series(s, aggressive=aggressive)
    out = pd.Series(out, index=df.index).replace([np.inf, -np.inf], np.nan)
    return out.dropna().astype(float)


def to_finite_numeric_frame(df, cols, aggressive=True):
    """批量把字段转换为有限数值，用于散点、相关矩阵、异常检测等。"""
    out = pd.DataFrame(index=df.index)
    for c in cols:
        if c not in df.columns:
            continue
        if pd.api.types.is_numeric_dtype(df[c]):
            s = pd.to_numeric(df[c], errors="coerce")
        else:
            s = parse_numeric_series(df[c], aggressive=aggressive)
        s = pd.Series(s, index=df.index).replace([np.inf, -np.inf], np.nan)
        out[c] = s.astype(float)
    return out


def has_enough_numeric(df, col, min_count=1, require_variation=False):
    s = to_finite_numeric_series(df, col)
    if len(s) < min_count:
        return False
    if require_variation and s.nunique(dropna=True) <= 1:
        return False
    return True


def show_no_chart_reason(title, reason):
    """统一的非空白提示：没有可画数据时只提示原因，不再展示空坐标轴。"""
    st.warning(f"“{title}”暂无法绘制：{reason}")


def figure_has_data(fig):
    """判断 figure 是否真的有可展示的数据点，避免出现只有坐标轴的空图。"""
    try:
        for tr in fig.data:
            for attr in ["x", "y", "z", "values", "r"]:
                vals = getattr(tr, attr, None)
                if vals is None:
                    continue
                arr = np.asarray(vals, dtype=object)
                if arr.size == 0:
                    continue
                # z 可能是二维
                flat = arr.ravel()
                non_empty = [v for v in flat if v is not None and str(v) != "nan"]
                if len(non_empty) > 0:
                    return True
    except Exception:
        return True
    return False


def safe_plotly_chart(fig, title="", height=500):
    fig = chart_layout(fig, height)
    if not figure_has_data(fig):
        show_no_chart_reason(title or "图表", "当前字段没有形成有效图形数据。请更换字段或检查数值转换结果。")
        return False
    st.plotly_chart(fig, use_container_width=True)
    return True



def safe_numeric_cols(df, cols):
    """不只看 dtype，而是重新尝试数值解析，防止 Excel/CSV 数字被当成文本导致图表空白。"""
    valid = []
    for c in cols:
        if c not in df.columns or is_id_like(c):
            continue
        s = to_finite_numeric_series(df, c)
        if len(s) > 0:
            # 同步回 df 里的数值列，后续所有分析都使用干净数值
            try:
                df[c] = to_finite_numeric_frame(df, [c])[c]
            except Exception:
                pass
            valid.append(c)
    return valid


def bar_chart(df, x, y, title):
    if x not in df.columns or y not in df.columns:
        show_no_chart_reason(title, "字段不存在。")
        return
    plot = df[[x, y]].copy()
    plot[y] = pd.to_numeric(plot[y], errors="coerce").replace([np.inf, -np.inf], np.nan)
    plot = plot.dropna(subset=[x, y])
    if plot.empty:
        show_no_chart_reason(title, f"{y} 没有有效数值，或 {x} 没有有效分组。")
        return

    # 类别过多时保留前 30，避免图表挤压；如果传入前面已经 topN，这里不会改变结果
    if len(plot) > 30:
        plot = plot.sort_values(y, ascending=False).head(30)

    horizontal = len(plot) >= 8
    fig = go.Figure()
    if horizontal:
        plot = plot.sort_values(y, ascending=True)
        fig.add_trace(go.Bar(
            x=plot[y].astype(float).tolist(),
            y=plot[x].astype(str).tolist(),
            orientation="h",
            text=[money_fmt(v) for v in plot[y]],
            textposition="outside",
            marker=dict(line=dict(width=0))
        ))
        fig.update_layout(title=title, xaxis_title=y, yaxis_title=x)
    else:
        fig.add_trace(go.Bar(
            x=plot[x].astype(str).tolist(),
            y=plot[y].astype(float).tolist(),
            text=[money_fmt(v) for v in plot[y]],
            textposition="outside",
            marker=dict(line=dict(width=0))
        ))
        fig.update_layout(title=title, xaxis_title=x, yaxis_title=y)
    safe_plotly_chart(fig, title, 520)


def line_chart(df, x, y, title):
    if x not in df.columns or y not in df.columns:
        show_no_chart_reason(title, "字段不存在。")
        return
    plot = df[[x, y]].copy()
    plot[y] = pd.to_numeric(plot[y], errors="coerce").replace([np.inf, -np.inf], np.nan)
    plot = plot.dropna(subset=[x, y])
    if plot.empty:
        show_no_chart_reason(title, f"{y} 没有有效数值，或 {x} 没有有效取值。")
        return

    x_values = plot[x].astype(str).tolist()
    y_values = plot[y].astype(float).tolist()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x_values,
        y=y_values,
        mode="lines+markers",
        line=dict(width=4),
        marker=dict(size=8, line=dict(width=1, color="#FFFFFF")),
        hovertemplate=f"{x}=%{{x}}<br>{y}=%{{y:,.2f}}<extra></extra>"
    ))
    fig.update_layout(title=title, xaxis_title=x, yaxis_title=y, hovermode="x unified")

    # 时间点较多时只显示少量横轴刻度，避免密密麻麻挤在一起
    if len(x_values) > 12:
        step = max(1, len(x_values) // 8)
        tickvals = [x_values[i] for i in range(0, len(x_values), step)]
        if x_values[-1] not in tickvals:
            tickvals.append(x_values[-1])
        fig.update_xaxes(type="category", tickmode="array", tickvals=tickvals, tickangle=0)

    safe_plotly_chart(fig, title, 540)

def histogram_chart(df, col, title):
    s = to_finite_numeric_series(df, col)
    if s.empty:
        show_no_chart_reason(title, f"{col} 没有可解析为数值的有效数据。")
        return

    values = s.astype(float).tolist()
    bins = min(45, max(8, int(np.sqrt(len(values)) * 2)))
    fig = go.Figure()
    fig.add_trace(go.Histogram(
        x=values,
        nbinsx=bins,
        opacity=0.88,
        marker=dict(line=dict(width=0))
    ))
    fig.update_layout(title=title, xaxis_title=col, yaxis_title="记录数", bargap=0.05)
    safe_plotly_chart(fig, title, 520)


def box_chart(df, col, title):
    s = to_finite_numeric_series(df, col)
    if s.empty:
        show_no_chart_reason(title, f"{col} 没有可解析为数值的有效数据。")
        return
    fig = go.Figure()
    fig.add_trace(go.Box(
        y=s.astype(float).tolist(),
        boxpoints="outliers",
        name=col,
        marker=dict(size=5, opacity=0.55)
    ))
    fig.update_layout(title=title, yaxis_title=col)
    safe_plotly_chart(fig, title, 520)


def recommend_fit_method(df, x, y, group_col=None):
    use_cols = [x, y] + ([group_col] if group_col else [])
    plot = df[use_cols].copy()
    num = to_finite_numeric_frame(plot, [x, y])
    plot[x] = num[x] if x in num.columns else np.nan
    plot[y] = num[y] if y in num.columns else np.nan
    plot = plot.dropna(subset=[x, y])
    if len(plot) < 30:
        return "全局OLS（直线）", "样本量较少，建议先使用全局OLS观察整体方向。"
    group_n = plot[group_col].nunique() if group_col and group_col in plot.columns else 0
    corr = abs(plot[[x, y]].corr().iloc[0, 1]) if plot[x].nunique() > 1 and plot[y].nunique() > 1 else 0
    if group_n >= 2 and group_n <= 12:
        return "按维度分组OLS（分段直线）", f"当前存在 {group_n} 个有效分组，不同群体可能有不同斜率，建议使用分组OLS比较差异。"
    if corr < 0.25 and len(plot) >= 80:
        return "LOWESS（局部加权）", "整体线性相关较弱，可能存在非线性或局部拐点，建议使用LOWESS观察局部变化。"
    if len(plot) >= 80:
        return "多项式回归（二阶）", "样本量较充足，可尝试二阶多项式观察是否存在先升后降或先降后升关系。"
    return "全局OLS（直线）", "当前数据适合先用全局OLS作为整体趋势参考。"


def _add_ols_line(fig, x_values, y_values, name, dash=None):
    temp = pd.DataFrame({"x": pd.to_numeric(pd.Series(x_values), errors="coerce"),
                         "y": pd.to_numeric(pd.Series(y_values), errors="coerce")}).replace([np.inf, -np.inf], np.nan).dropna()
    if len(temp) < 3 or temp["x"].nunique() <= 1:
        return
    coef = np.polyfit(temp["x"], temp["y"], 1)
    xs = np.linspace(temp["x"].min(), temp["x"].max(), 100)
    ys = coef[0] * xs + coef[1]
    fig.add_trace(go.Scatter(x=xs.tolist(), y=ys.tolist(), mode="lines", name=name, line=dict(width=3, dash=dash or "solid")))


def _add_poly2_line(fig, x_values, y_values, name):
    temp = pd.DataFrame({"x": pd.to_numeric(pd.Series(x_values), errors="coerce"),
                         "y": pd.to_numeric(pd.Series(y_values), errors="coerce")}).replace([np.inf, -np.inf], np.nan).dropna()
    if len(temp) < 5 or temp["x"].nunique() <= 2:
        return
    coef = np.polyfit(temp["x"], temp["y"], 2)
    xs = np.linspace(temp["x"].min(), temp["x"].max(), 160)
    ys = coef[0] * xs ** 2 + coef[1] * xs + coef[2]
    fig.add_trace(go.Scatter(x=xs.tolist(), y=ys.tolist(), mode="lines", name=name, line=dict(width=3, dash="dot")))


def _add_lowess_line(fig, x_values, y_values, name):
    temp = pd.DataFrame({"x": pd.to_numeric(pd.Series(x_values), errors="coerce"),
                         "y": pd.to_numeric(pd.Series(y_values), errors="coerce")}).replace([np.inf, -np.inf], np.nan).dropna().sort_values("x")
    if len(temp) < 10 or temp["x"].nunique() <= 5:
        return
    try:
        from statsmodels.nonparametric.smoothers_lowess import lowess
        smoothed = lowess(temp["y"], temp["x"], frac=0.28, return_sorted=True)
        fig.add_trace(go.Scatter(x=smoothed[:, 0].tolist(), y=smoothed[:, 1].tolist(), mode="lines", name=name, line=dict(width=4)))
    except Exception:
        try:
            temp["_bin"] = pd.qcut(temp["x"], q=min(20, max(5, len(temp) // 20)), duplicates="drop")
            binned = temp.groupby("_bin", observed=True).agg(x=("x", "mean"), y=("y", "mean")).dropna()
            fig.add_trace(go.Scatter(x=binned["x"].tolist(), y=binned["y"].tolist(), mode="lines+markers", name=name, line=dict(width=4)))
        except Exception:
            return


def scatter_chart(df, x, y, color_col=None, fit_method="全局OLS（直线）"):
    if x not in df.columns or y not in df.columns:
        show_no_chart_reason(f"{x} 与 {y} 的关系", "字段不存在。")
        return np.nan

    cols = [x, y] + ([color_col] if color_col and color_col in df.columns else [])
    plot = df[cols].copy()
    num = to_finite_numeric_frame(plot, [x, y])
    if x not in num.columns or y not in num.columns:
        show_no_chart_reason(f"{x} 与 {y} 的关系", "横轴或纵轴无法解析为有效数值。")
        return np.nan
    plot[x] = num[x]
    plot[y] = num[y]
    plot = plot.dropna(subset=[x, y])
    if plot.empty:
        show_no_chart_reason(f"{x} 与 {y} 的关系", "两个字段没有足够的共同有效数值。")
        return np.nan

    if len(plot) > 2500:
        plot = plot.sample(2500, random_state=42)

    fig = go.Figure()
    if color_col and color_col in plot.columns:
        groups = list(plot.groupby(color_col, dropna=False))
        for idx, (name, g) in enumerate(groups):
            fig.add_trace(go.Scatter(
                x=g[x].astype(float).tolist(),
                y=g[y].astype(float).tolist(),
                mode="markers",
                name=str(name),
                marker=dict(size=6, opacity=0.48),
                showlegend=idx < 20
            ))
    else:
        fig.add_trace(go.Scatter(
            x=plot[x].astype(float).tolist(),
            y=plot[y].astype(float).tolist(),
            mode="markers",
            name="样本点",
            marker=dict(size=6, opacity=0.50)
        ))

    corr = plot[[x, y]].corr().iloc[0, 1] if len(plot) >= 3 and plot[x].nunique() > 1 and plot[y].nunique() > 1 else np.nan

    if fit_method == "全局OLS（直线）":
        _add_ols_line(fig, plot[x], plot[y], "全局OLS趋势线", dash="dash")
    elif fit_method == "按维度分组OLS（分段直线）":
        if color_col and color_col in plot.columns and plot[color_col].nunique(dropna=False) <= 15:
            for name, g in plot.groupby(color_col, dropna=False):
                _add_ols_line(fig, g[x], g[y], f"{name} OLS")
        else:
            st.info("当前未选择合适分组字段，或分组数量过多，已改用全局OLS。")
            _add_ols_line(fig, plot[x], plot[y], "全局OLS趋势线", dash="dash")
    elif fit_method == "LOWESS（局部加权）":
        _add_lowess_line(fig, plot[x], plot[y], "LOWESS局部趋势")
    elif fit_method == "多项式回归（二阶）":
        _add_poly2_line(fig, plot[x], plot[y], "二阶多项式趋势")

    fig.update_layout(title=f"{x} 与 {y} 的关系", xaxis_title=x, yaxis_title=y, hovermode="closest")
    safe_plotly_chart(fig, f"{x} 与 {y} 的关系", 540)
    return corr


def heatmap_chart(pivot, title, colorscale="Blues"):
    if pivot is None or pivot.empty:
        st.warning("当前交叉维度下没有可用于绘制热力图的数据。")
        return
    plot = pivot.copy()
    plot = plot.apply(pd.to_numeric, errors="coerce").replace([np.inf, -np.inf], np.nan)
    plot = plot.dropna(how="all", axis=0).dropna(how="all", axis=1)
    if plot.empty:
        show_no_chart_reason(title, "透视表没有有效数值。")
        return

    z = plot.fillna(0).astype(float).values.tolist()
    x_labels = [str(c) for c in plot.columns.tolist()]
    y_labels = [str(i) for i in plot.index.tolist()]

    fig = go.Figure(data=go.Heatmap(
        z=z,
        x=x_labels,
        y=y_labels,
        colorscale=colorscale,
        colorbar=dict(title="数值"),
        hovertemplate="%{y} × %{x}<br>数值=%{z:,.2f}<extra></extra>"
    ))

    # 单元格不太多时加文本，增强可读性；避免过多文字把图压乱
    if len(x_labels) * len(y_labels) <= 100:
        for i, yy in enumerate(y_labels):
            for j, xx in enumerate(x_labels):
                fig.add_annotation(
                    x=xx, y=yy, text=money_fmt(plot.iloc[i, j]),
                    showarrow=False, font=dict(size=12, color="#17324D")
                )

    fig.update_layout(
        title=title,
        xaxis_title=str(pivot.columns.name) if pivot.columns.name else "列维度",
        yaxis_title=str(pivot.index.name) if pivot.index.name else "行维度",
        xaxis=dict(type="category", tickangle=-30, automargin=True),
        yaxis=dict(type="category", automargin=True)
    )
    safe_plotly_chart(fig, title, 560)


def corr_heatmap_chart(df, numeric_cols):
    # 优先用用户选中的数值字段；如果不足，自动扫描全表中能解析为数值的字段
    scan_cols = list(dict.fromkeys(list(numeric_cols) + list(df.columns)))
    numeric_df = to_finite_numeric_frame(df, [c for c in scan_cols if c in df.columns and not is_id_like(c)])
    cols = []
    for c in numeric_df.columns:
        valid = numeric_df[c].dropna()
        if len(valid) >= 3 and valid.nunique(dropna=True) > 1:
            cols.append(c)
    if len(cols) < 2:
        st.warning("可用于相关性矩阵的数值字段不足。请至少选择两个有效数值字段，且字段不能全为同一个值。")
        return

    corr = numeric_df[cols].corr(method="pearson").replace([np.inf, -np.inf], np.nan)
    corr = corr.dropna(how="all", axis=0).dropna(how="all", axis=1)
    if corr.empty or corr.shape[0] < 2 or corr.shape[1] < 2:
        show_no_chart_reason("数值指标相关性矩阵", "相关系数无法计算，可能是字段有效值过少或没有波动。")
        return

    for c in corr.columns:
        if c in corr.index:
            corr.loc[c, c] = 1.0

    z = corr.fillna(0).astype(float).values.tolist()
    x_labels = corr.columns.astype(str).tolist()
    y_labels = corr.index.astype(str).tolist()

    fig = go.Figure(data=go.Heatmap(
        z=z,
        x=x_labels,
        y=y_labels,
        zmin=-1,
        zmax=1,
        colorscale="RdBu",
        reversescale=True,
        colorbar=dict(title="相关系数"),
        hovertemplate="%{y} 与 %{x}<br>相关系数=%{z:.3f}<extra></extra>"
    ))

    if len(x_labels) * len(y_labels) <= 144:
        for i, yy in enumerate(y_labels):
            for j, xx in enumerate(x_labels):
                fig.add_annotation(
                    x=xx, y=yy, text=f"{corr.iloc[i, j]:.2f}",
                    showarrow=False, font=dict(size=12, color="#17324D")
                )

    fig.update_layout(
        title="数值指标相关性矩阵",
        xaxis_title="数值字段",
        yaxis_title="数值字段",
        xaxis=dict(type="category", tickangle=-30, automargin=True),
        yaxis=dict(type="category", automargin=True, autorange="reversed")
    )
    safe_plotly_chart(fig, "数值指标相关性矩阵", 600)

    pairs = []
    for i, a in enumerate(corr.columns):
        for j, b in enumerate(corr.index):
            if j <= i:
                continue
            val = corr.loc[b, a]
            if pd.notna(val) and abs(val) >= 0.6:
                pairs.append((a, b, val))
    if pairs:
        pairs = sorted(pairs, key=lambda t: abs(t[2]), reverse=True)[:5]
        msg = "；".join([f"{a} 与 {b} 的相关系数为 {v:.2f}" for a, b, v in pairs])
        st.info(f"系统识别到较强相关关系：{msg}。这些字段可能是影响主指标变化的重要线索。")
    else:
        st.info("当前数值字段之间未发现特别强的线性相关关系。接近 1 表示强正相关，接近 -1 表示强负相关，接近 0 表示线性关系较弱。")

def periodize(df, date_col):
    temp = df.copy()
    temp["_period"] = pd.to_datetime(temp[date_col], errors="coerce").dt.to_period("M").astype(str)
    temp = temp[temp["_period"].notna()]
    return temp


def build_trend(df, date_col, metric):
    temp = periodize(df, date_col)
    if temp.empty:
        return pd.DataFrame(columns=["期间", metric])
    trend = temp.groupby("_period", dropna=True)[metric].sum().reset_index()
    trend = trend.rename(columns={"_period": "期间"}).sort_values("期间")
    return trend



def trend_interpretation(trend, metric):
    if trend is None or len(trend) < 2:
        return "当前数据不足以判断趋势。建议至少提供两个以上时间周期的数据。"
    trend = trend.dropna(subset=[metric]).copy()
    if len(trend) < 2:
        return "当前主指标在时间维度上的有效数值不足，暂无法形成稳定趋势判断。"
    first, last = trend[metric].iloc[0], trend[metric].iloc[-1]
    change = (last-first)/abs(first) if first != 0 else np.nan
    diffs = trend[metric].diff().dropna()
    up_count, down_count, flat_count = int((diffs>0).sum()), int((diffs<0).sum()), int((diffs==0).sum())
    recent = trend.tail(min(3, len(trend)))
    recent_change = np.nan
    if len(recent) >= 2 and recent[metric].iloc[0] != 0:
        recent_change = (recent[metric].iloc[-1] - recent[metric].iloc[0]) / abs(recent[metric].iloc[0])
    max_row = trend.loc[trend[metric].idxmax()]
    min_row = trend.loc[trend[metric].idxmin()]
    mean_val = trend[metric].mean()
    cv = trend[metric].std()/abs(mean_val) if mean_val != 0 else np.nan
    if pd.isna(change): main_desc = "整体变化方向暂无法计算"
    elif change > 0.1: main_desc = f"整体呈上升趋势，末期较初期增长约 {change:.2%}"
    elif change < -0.1: main_desc = f"整体呈下降趋势，末期较初期下降约 {abs(change):.2%}"
    else: main_desc = "整体较为平稳，期初与期末差异不大"
    if not pd.isna(recent_change):
        if recent_change > 0.08: recent_desc = f"最近 {len(recent)} 期继续上行，累计增长约 {recent_change:.2%}"
        elif recent_change < -0.08: recent_desc = f"最近 {len(recent)} 期出现回落，累计下降约 {abs(recent_change):.2%}"
        else: recent_desc = f"最近 {len(recent)} 期变化幅度较小，短期表现相对稳定"
    else: recent_desc = "短期变化暂无法计算"
    if pd.isna(cv): vol_desc = "波动性暂无法计算"
    elif cv < 0.15: vol_desc = "整体波动较低，指标运行较稳定"
    elif cv < 0.45: vol_desc = "存在一定波动，需要结合高低点进一步观察"
    else: vol_desc = "波动较强，建议重点关注异常月份或阶段性变化"
    detail = f"在全部相邻周期中，上升 {up_count} 次、下降 {down_count} 次、持平 {flat_count} 次。"
    high_low = f"最高值出现在 {max_row['期间']}，为 {money_fmt(max_row[metric])}；最低值出现在 {min_row['期间']}，为 {money_fmt(min_row[metric])}。"
    return f"{main_desc}。{recent_desc}。{detail}{high_low}{vol_desc}。"

def dimension_summary(df, dim, metric):
    g = df.groupby(dim, dropna=False)[metric].agg(["sum", "mean", "count"]).reset_index()
    g.columns = [dim, f"{metric}合计", f"{metric}均值", "记录数"]
    return g.sort_values(f"{metric}合计", ascending=False)


def dimension_importance(df, dimensions, metric):
    rows = []
    total_var = df[metric].var()
    for dim in dimensions:
        if dim not in df.columns:
            continue
        g = df.groupby(dim, dropna=False)[metric].mean()
        score = 0 if pd.isna(total_var) or total_var == 0 else min(float(g.var() / total_var), 1)
        rows.append({"维度": dim, "差异解释度": score, "维度取值数": int(df[dim].nunique(dropna=True))})
    return pd.DataFrame(rows).sort_values("差异解释度", ascending=False) if rows else pd.DataFrame()



def contribution_analysis(df, date_col, dim, metric, start_period=None, end_period=None):
    if not date_col or not dim: return pd.DataFrame()
    temp = periodize(df, date_col)
    periods = sorted(temp['_period'].dropna().unique().tolist())
    if len(periods) < 2: return pd.DataFrame()
    if start_period is None or end_period is None:
        p0, p1 = periods[-2], periods[-1]
    else:
        p0, p1 = start_period, end_period
    if p0 == p1: return pd.DataFrame()
    a = temp[temp['_period']==p0].groupby(dim)[metric].sum()
    b = temp[temp['_period']==p1].groupby(dim)[metric].sum()
    comp = pd.concat([a,b], axis=1).fillna(0)
    comp.columns = [p0,p1]
    comp['变化额'] = comp[p1]-comp[p0]
    total_change = comp['变化额'].sum()
    comp['贡献占比'] = comp['变化额']/total_change if total_change != 0 else np.nan
    return comp.reset_index().sort_values('变化额', ascending=False)

def relationship_explanation(x, y, corr):
    if pd.isna(corr):
        return "当前两个指标缺失较多，暂无法计算稳定的相关性。"
    strength = "较强" if abs(corr) >= 0.6 else ("中等" if abs(corr) >= 0.3 else "较弱")
    direction = "正向" if corr > 0 else ("反向" if corr < 0 else "无明显")
    return f"相关系数为 {corr:.3f}，说明 {x} 与 {y} 的线性相关性{strength}，方向为{direction}关系。散点图主要用于观察两个数值指标之间是否存在同向、反向或异常偏离关系。"




def correlation_summary_tables(df, numeric_cols):
    """汇总所有经营数值指标之间的相关系数，便于用户直接查看整体关系。"""
    scan_cols = [c for c in numeric_cols if c in df.columns and not is_id_like(c)]
    numeric_df = to_finite_numeric_frame(df, scan_cols)
    valid_cols = []
    for c in numeric_df.columns:
        s = numeric_df[c].dropna()
        if len(s) >= 3 and s.nunique(dropna=True) > 1:
            valid_cols.append(c)
    if len(valid_cols) < 2:
        return pd.DataFrame(), pd.DataFrame()

    corr = numeric_df[valid_cols].corr(method="pearson").replace([np.inf, -np.inf], np.nan)
    corr = corr.dropna(how="all", axis=0).dropna(how="all", axis=1)

    rows = []
    cols = corr.columns.tolist()
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            a, b = cols[i], cols[j]
            val = corr.loc[a, b]
            if pd.isna(val):
                continue
            abs_val = abs(float(val))
            if abs_val >= 0.7:
                strength = "强相关"
            elif abs_val >= 0.4:
                strength = "中等相关"
            elif abs_val >= 0.2:
                strength = "弱相关"
            else:
                strength = "相关较弱"
            direction = "正相关" if val > 0 else ("负相关" if val < 0 else "无明显方向")
            rows.append({
                "指标A": a,
                "指标B": b,
                "相关系数": round(float(val), 3),
                "绝对值": round(abs_val, 3),
                "方向": direction,
                "强度判断": strength,
                "经营含义提示": f"{a} 与 {b} 呈{strength}、{direction}，可作为后续驱动分析或交叉核查线索。"
            })
    pair_df = pd.DataFrame(rows).sort_values("绝对值", ascending=False) if rows else pd.DataFrame()
    return corr.round(3), pair_df


def parse_markdown_tables(text):
    """将大模型输出中的 Markdown 表格解析成 DataFrame，避免页面显示 |---|---|。"""
    if not text:
        return [], ""
    lines = str(text).splitlines()
    blocks, normal_lines = [], []
    i = 0
    sep_pat = r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$"
    while i < len(lines):
        line = lines[i].strip()
        if "|" in line and i + 1 < len(lines) and re.match(sep_pat, lines[i + 1].strip()):
            table_lines = [lines[i].strip(), lines[i + 1].strip()]
            i += 2
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1
            header = [x.strip() for x in table_lines[0].strip("|").split("|")]
            data = []
            for row_line in table_lines[2:]:
                cells = [x.strip() for x in row_line.strip("|").split("|")]
                if len(cells) < len(header):
                    cells += [""] * (len(header) - len(cells))
                data.append(cells[:len(header)])
            df_table = pd.DataFrame(data, columns=header)
            before = "\n".join(normal_lines).strip()
            blocks.append((before, df_table))
            normal_lines = []
        else:
            normal_lines.append(lines[i])
            i += 1
    return blocks, "\n".join(normal_lines).strip()


def render_ai_answer_pretty(ans):
    """美化大模型解释结果：正文正常显示，Markdown表格转成真正表格。"""
    blocks, tail = parse_markdown_tables(ans)
    if not blocks:
        st.markdown(ans)
        return
    for before, table_df in blocks:
        if before:
            st.markdown(before)
        st.markdown('<div class="ai-table-card"><div class="ai-section-title">结构化核查清单</div></div>', unsafe_allow_html=True)
        st.dataframe(table_df, use_container_width=True, hide_index=True)
    if tail:
        st.markdown(tail)


def _sparse_period_ticks(periods, max_ticks=8):
    periods = [str(x) for x in periods]
    if len(periods) <= max_ticks:
        return periods
    step = max(1, int(np.ceil(len(periods) / max_ticks)))
    vals = [periods[i] for i in range(0, len(periods), step)]
    if periods[-1] not in vals:
        vals.append(periods[-1])
    return vals


def animated_trend_chart(trend, metric, title=None):
    """
    更清晰的动态趋势演示：
    1. 每一帧展示从起点到当前周期的累计轨迹；
    2. 横轴只显示少量区间刻度，避免时间标签密集重叠；
    3. 单独高亮当前点，增强趋势演变可读性。
    """
    if trend is None or len(trend) < 3 or "期间" not in trend.columns or metric not in trend.columns:
        st.info("当前未生成动态图：需要有效日期字段，并且时间周期不少于 3 个。")
        return

    plot = trend[["期间", metric]].dropna().copy()
    plot[metric] = pd.to_numeric(plot[metric], errors="coerce")
    plot = plot.dropna()
    if len(plot) < 3:
        st.info("当前未生成动态图：当前主指标在时间维度上的有效点不足。")
        return

    periods = plot["期间"].astype(str).tolist()
    values = plot[metric].astype(float).tolist()
    ymin, ymax = min(values), max(values)
    pad = (ymax - ymin) * 0.15 if ymax != ymin else max(abs(ymax) * 0.1, 1)
    tickvals = _sparse_period_ticks(periods, max_ticks=8)

    frames = []
    for i in range(1, len(plot) + 1):
        frames.append(go.Frame(
            data=[
                go.Scatter(
                    x=periods[:i],
                    y=values[:i],
                    mode="lines",
                    fill="tozeroy",
                    line=dict(width=5),
                    name=f"{metric}累计轨迹",
                    hovertemplate="期间=%{x}<br>数值=%{y:,.2f}<extra></extra>"
                ),
                go.Scatter(
                    x=[periods[i - 1]],
                    y=[values[i - 1]],
                    mode="markers+text",
                    marker=dict(size=16, symbol="circle"),
                    text=[money_fmt(values[i - 1])],
                    textposition="top center",
                    name="当前周期",
                    hovertemplate="当前周期=%{x}<br>数值=%{y:,.2f}<extra></extra>"
                )
            ],
            name=periods[i - 1]
        ))

    fig = go.Figure(
        data=[
            go.Scatter(
                x=[periods[0]],
                y=[values[0]],
                mode="lines",
                fill="tozeroy",
                line=dict(width=5),
                name=f"{metric}累计轨迹"
            ),
            go.Scatter(
                x=[periods[0]],
                y=[values[0]],
                mode="markers+text",
                marker=dict(size=16),
                text=[money_fmt(values[0])],
                textposition="top center",
                name="当前周期"
            )
        ],
        frames=frames
    )

    slider_steps = []
    label_step = max(1, int(np.ceil(len(periods) / 8)))
    for i, p in enumerate(periods):
        slider_steps.append({
            "args": [[p], {"frame": {"duration": 0, "redraw": True}, "mode": "immediate"}],
            "label": p if (i % label_step == 0 or i == len(periods) - 1) else "",
            "method": "animate"
        })

    fig.update_layout(
        title=title or f"{metric}动态趋势演示",
        xaxis=dict(
            title="期间",
            type="category",
            categoryorder="array",
            categoryarray=periods,
            tickmode="array",
            tickvals=tickvals,
            tickangle=0
        ),
        yaxis=dict(title=metric, range=[ymin - pad, ymax + pad]),
        hovermode="x unified",
        updatemenus=[{
            "type": "buttons",
            "showactive": False,
            "x": 0.02,
            "y": 1.15,
            "buttons": [
                {"label": "播放", "method": "animate", "args": [None, {"frame": {"duration": 520, "redraw": True}, "fromcurrent": True}]},
                {"label": "暂停", "method": "animate", "args": [[None], {"frame": {"duration": 0, "redraw": False}, "mode": "immediate"}]}
            ]
        }],
        sliders=[{
            "steps": slider_steps,
            "currentvalue": {"prefix": "当前周期：", "font": {"size": 15}},
            "x": 0.08,
            "len": 0.86,
            "pad": {"t": 45, "b": 10}
        }]
    )
    safe_plotly_chart(fig, title or f"{metric}动态趋势演示", 590)


def clean_md_inline(text):
    text = str(text)
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"`([^`]*)`", r"\1", text)
    text = text.replace("---", "").strip()
    return text


def add_dataframe_to_docx(doc, df_table, max_rows=12):
    if df_table is None or len(df_table) == 0:
        return
    show = df_table.head(max_rows).copy()
    table = doc.add_table(rows=1, cols=len(show.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for j, col in enumerate(show.columns):
        run = table.rows[0].cells[j].paragraphs[0].add_run(str(col))
        set_run_font(run, "黑体", 10, True)
    for _, row in show.iterrows():
        cells = table.add_row().cells
        for j, col in enumerate(show.columns):
            run = cells[j].paragraphs[0].add_run(clean_md_inline(row[col]))
            set_run_font(run, "宋体", 9, False)


def add_native_bar_visual(doc, title, rows, label_col="项目", value_col="数值", max_rows=10):
    """用 Word 原生表格做条形图，避免云端导出图片时中文乱码。"""
    if rows is None or len(rows) == 0:
        return
    add_heading(doc, title, 2)
    data = pd.DataFrame(rows).copy().head(max_rows)
    data[value_col] = pd.to_numeric(data[value_col], errors="coerce").fillna(0)
    max_val = data[value_col].abs().max()
    max_val = max(max_val, 1)

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    headers = [label_col, value_col, "可视化条"]
    for j, h in enumerate(headers):
        run = table.rows[0].cells[j].paragraphs[0].add_run(h)
        set_run_font(run, "黑体", 10, True)

    for _, r in data.iterrows():
        cells = table.add_row().cells
        label = str(r[label_col])
        val = float(r[value_col])
        bar_len = int(abs(val) / max_val * 24)
        bar = "█" * max(1, bar_len)
        vals = [label, money_fmt(val), bar]
        for j, v in enumerate(vals):
            run = cells[j].paragraphs[0].add_run(str(v))
            set_run_font(run, "宋体", 9, False)


def add_trend_visual_table(doc, trend, metric):
    if trend is None or len(trend) == 0:
        return
    show = trend.tail(10).copy()
    show[metric] = pd.to_numeric(show[metric], errors="coerce").fillna(0)
    rows = [{"期间": r["期间"], "数值": r[metric]} for _, r in show.iterrows()]
    add_native_bar_visual(doc, f"{metric}近10期趋势可视化", rows, "期间", "数值", 10)


def add_dimension_visual_table(doc, df, dim, metric):
    if not dim or dim not in df.columns:
        return
    g = dimension_summary(df, dim, metric).head(10)
    rows = [{dim: r[dim], f"{metric}合计": r[f"{metric}合计"]} for _, r in g.iterrows()]
    add_native_bar_visual(doc, f"按{dim}的{metric}Top10可视化", rows, dim, f"{metric}合计", 10)


def add_risk_visual_table(doc, anomaly_df):
    if anomaly_df is None or len(anomaly_df) == 0 or "风险等级" not in anomaly_df.columns:
        return
    rc = anomaly_df["风险等级"].value_counts().reset_index()
    rc.columns = ["风险等级", "数量"]
    rows = [{"风险等级": r["风险等级"], "数量": r["数量"]} for _, r in rc.iterrows()]
    add_native_bar_visual(doc, "风险等级分布可视化", rows, "风险等级", "数量", 10)


def add_markdown_body_to_docx(doc, text, df, main_metric, dimensions, date_col, anomaly_df):
    """
    将 AI Markdown 正文整理成正式 Word 结构：
    - 去掉 ###、**、--- 等 Markdown 符号；
    - 一级/二级标题转为 Word 标题；
    - Markdown 表格转为 Word 表格；
    - 在对应章节中穿插原生可视化表，不再把图表集中堆到最后。
    """
    lines = str(text).splitlines()
    current_section = ""
    i = 0
    sep_pat = r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$"
    inserted_trend = False
    inserted_dim = False
    inserted_risk = False

    while i < len(lines):
        raw = lines[i].strip()
        if not raw or raw == "---":
            i += 1
            continue

        # Markdown table
        if "|" in raw and i + 1 < len(lines) and re.match(sep_pat, lines[i + 1].strip()):
            table_lines = [raw, lines[i + 1].strip()]
            i += 2
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1
            header = [clean_md_inline(x.strip()) for x in table_lines[0].strip("|").split("|")]
            data = []
            for row_line in table_lines[2:]:
                cells = [clean_md_inline(x.strip()) for x in row_line.strip("|").split("|")]
                if len(cells) < len(header):
                    cells += [""] * (len(header) - len(cells))
                data.append(cells[:len(header)])
            add_dataframe_to_docx(doc, pd.DataFrame(data, columns=header), max_rows=15)
            continue

        # Headings
        heading_match = re.match(r"^#{1,6}\s*(.+)$", raw)
        cn_heading_match = re.match(r"^([一二三四五六七八九十]+、.+)$", raw)
        num_heading_match = re.match(r"^(\d+[\.、].+)$", raw)

        if heading_match or cn_heading_match or num_heading_match:
            title = heading_match.group(1) if heading_match else (cn_heading_match.group(1) if cn_heading_match else num_heading_match.group(1))
            title = clean_md_inline(title)
            if "企业经营分析决策简报" in title or "基于" in title and "主指标" in title:
                p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(title); set_run_font(run, "黑体", 14, True)
            else:
                current_section = title
                add_heading(doc, title, 1 if re.match(r"^[一二三四五六七八九十]+、", title) else 2)

                if ("核心经营表现" in title or "时间趋势" in title) and not inserted_trend:
                    if date_col:
                        trend = build_trend(df, date_col, main_metric)
                        if len(trend) >= 2:
                            add_trend_visual_table(doc, trend, main_metric)
                            inserted_trend = True
                    if dimensions and not inserted_dim:
                        add_dimension_visual_table(doc, df, dimensions[0], main_metric)
                        inserted_dim = True

                if ("风险识别" in title or "异常" in title) and not inserted_risk:
                    add_risk_visual_table(doc, anomaly_df)
                    inserted_risk = True
            i += 1
            continue

        # Bullets / normal paragraphs
        clean = clean_md_inline(raw)
        clean = re.sub(r"^[\-\*]\s+", "", clean)
        clean = clean.replace("✅", "【行动】").strip()
        if clean:
            add_para(doc, clean, indent=not clean.startswith(("【行动】", "注：", ">")))
        i += 1

    # 如果 AI 文本没有触发对应标题，也补充可视化，保证报告有图表类内容
    if date_col and not inserted_trend:
        trend = build_trend(df, date_col, main_metric)
        if len(trend) >= 2:
            add_trend_visual_table(doc, trend, main_metric)
    if dimensions and not inserted_dim:
        add_dimension_visual_table(doc, df, dimensions[0], main_metric)
    if anomaly_df is not None and len(anomaly_df) and not inserted_risk:
        add_risk_visual_table(doc, anomaly_df)




def extract_clean_ai_blocks(ai_text):
    """提取 AI 正文中的有效段落和表格，去掉报告大标题、一级编号标题和 Markdown 标记，避免 Word 中出现重复“一、二、三”。"""
    text = str(ai_text or "")
    table_blocks, tail = parse_markdown_tables(text)
    lines = text.splitlines()
    paras = []
    for line in lines:
        raw = line.strip()
        if not raw or raw == "---":
            continue
        if "|" in raw:
            continue
        # 删除 Markdown 标题、报告题名和一级编号标题
        if re.match(r"^#{1,6}\s*", raw):
            raw = re.sub(r"^#{1,6}\s*", "", raw).strip()
            if re.match(r"^[一二三四五六七八九十]+[、.．]", raw):
                continue
        if re.match(r"^(企业经营分析决策简报|经营分析决策简报|智策经营|基于\s*\d)", raw):
            continue
        if re.match(r"^[一二三四五六七八九十]+[、.．]\s*", raw):
            # 不保留 AI 自带一级标题，避免与系统报告大纲冲突
            continue
        raw = re.sub(r"^[\-\*•]\s*", "", raw)
        raw = clean_md_inline(raw)
        raw = re.sub(r"^\d+[\.、]\s*", "", raw)
        raw = raw.strip()
        if raw and len(raw) >= 8:
            paras.append(raw)
    # 去重并限制长度
    dedup = []
    for p in paras:
        if p not in dedup:
            dedup.append(p)
    return dedup[:12], table_blocks


def add_default_action_table(doc, main_metric, dimensions):
    rows = [
        {"核查方向": "主指标异常单元", "具体内容": f"复核{main_metric}处于高位或低位的经营单元", "输出要求": "标注异常来源、责任维度和复核结论"},
        {"核查方向": "原始明细数据", "具体内容": "核查订单、项目、客户、产品等原始明细记录", "输出要求": "确认是否存在重复、缺失、口径不一致或录入错误"},
        {"核查方向": "关键维度下钻", "具体内容": f"围绕{dimensions[0] if dimensions else '主要维度'}进行分组对比", "输出要求": "识别贡献最高、波动最大或风险最集中的维度项"},
        {"核查方向": "管理动作", "具体内容": "结合风险等级设置复核优先级和跟踪周期", "输出要求": "形成可执行的整改措施和后续监控指标"},
    ]
    add_dataframe_to_docx(doc, pd.DataFrame(rows), max_rows=8)

def generate_ai_report_docx(ai_text, df, main_metric, dimensions, date_col, anomaly_df, focus_list):
    if not DOCX_AVAILABLE:
        raise RuntimeError("未安装 python-docx，请先安装：python -m pip install python-docx")

    doc = Document()
    title = doc.add_paragraph(); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("智策经营——AI增强版决策简报"); set_run_font(r, "黑体", 20, True)
    subtitle = doc.add_paragraph(); subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = subtitle.add_run("AI 驱动的多维经营分析与决策支持系统自动生成"); set_run_font(r, "宋体", 11, False); r.font.color.rgb = RGBColor(90,90,90)

    ai_paras, ai_tables = extract_clean_ai_blocks(ai_text)

    add_heading(doc, "一、分析概况", 1)
    add_heading(doc, "1.1 数据范围与分析口径", 2)
    add_para(doc, f"本报告基于当前上传数据自动生成，清洗后数据共 {len(df):,} 条记录，主分析指标为“{main_metric}”。系统根据上传数据中实际存在的字段开展分析，不引用未上传或不存在的指标。", indent=True)
    if focus_list:
        add_para(doc, f"本次简报关注重点包括：{'、'.join(focus_list)}。", indent=True)

    add_heading(doc, "1.2 核心指标摘要", 2)
    table = doc.add_table(rows=1, cols=2); table.alignment = WD_TABLE_ALIGNMENT.CENTER; table.style = "Table Grid"
    table.rows[0].cells[0].text = "指标"; table.rows[0].cells[1].text = "数值"
    abnormal_n = int((anomaly_df['是否经营异常']).sum()) if anomaly_df is not None and len(anomaly_df) and '是否经营异常' in anomaly_df.columns else 0
    for k, v in [(f"{main_metric}合计", money_fmt(df[main_metric].sum())), (f"{main_metric}均值", money_fmt(df[main_metric].mean())), (f"{main_metric}最大值", money_fmt(df[main_metric].max())), (f"{main_metric}最小值", money_fmt(df[main_metric].min())), ("数据记录数", f"{len(df):,}"), ("异常经营单元数", str(abnormal_n))]:
        cells = table.add_row().cells; cells[0].text = k; cells[1].text = v

    add_heading(doc, "二、核心经营表现", 1)
    add_heading(doc, "2.1 总体表现", 2)
    add_para(doc, f"当前样本中，“{main_metric}”合计为 {money_fmt(df[main_metric].sum())}，均值为 {money_fmt(df[main_metric].mean())}，最大值为 {money_fmt(df[main_metric].max())}，最小值为 {money_fmt(df[main_metric].min())}。", indent=True)
    if date_col:
        trend = build_trend(df, date_col, main_metric)
        add_heading(doc, "2.2 时间趋势分析", 2)
        add_para(doc, trend_interpretation(trend, main_metric), indent=True)
        if len(trend) >= 2:
            add_mpl_figure_to_docx(doc, make_mpl_trend_figure(trend, main_metric), f"图1 {main_metric}时间趋势")
    else:
        add_heading(doc, "2.2 结构分布分析", 2)
        add_para(doc, f"当前未选择有效日期字段，系统采用结构分布方式观察“{main_metric}”的集中程度和离散情况。", indent=True)
        add_mpl_figure_to_docx(doc, make_mpl_distribution_figure(df, main_metric), f"图1 {main_metric}分布情况")

    add_heading(doc, "三、多维结构分析", 1)
    if dimensions:
        add_heading(doc, "3.1 主要维度表现", 2)
        add_mpl_figure_to_docx(doc, make_mpl_dimension_bar_figure(df, dimensions[0], main_metric), f"图2 按{dimensions[0]}汇总{main_metric}")
        g = dimension_summary(df, dimensions[0], main_metric)
        if len(g):
            top = g.iloc[0]
            add_para(doc, f"从“{dimensions[0]}”维度看，{top[dimensions[0]]} 的 {main_metric} 合计最高，为 {money_fmt(top[f'{main_metric}合计'])}。该维度项可作为后续重点下钻对象。", indent=True)
    else:
        add_para(doc, "当前数据中可用于分组下钻的维度字段较少，建议补充地区、部门、客户类型、产品类别等维度字段，以提升结构分析效果。", indent=True)

    add_heading(doc, "四、风险识别与异常诊断", 1)
    if anomaly_df is not None and len(anomaly_df):
        abnormal = anomaly_df[anomaly_df["是否经营异常"]].copy() if "是否经营异常" in anomaly_df.columns else pd.DataFrame()
        high_n = int((abnormal.get("风险等级", pd.Series(dtype=str)) == "高风险").sum()) if len(abnormal) else 0
        add_para(doc, f"系统当前识别出 {len(abnormal)} 个异常经营单元，其中高风险单元 {high_n} 个。风险识别结果用于提示优先核查对象。", indent=True)
        add_mpl_figure_to_docx(doc, make_mpl_risk_figure(anomaly_df), "图3 风险等级分布")
        if len(abnormal):
            add_heading(doc, "4.1 重点异常单元", 2)
            top_anom = abnormal.sort_values("风险得分", ascending=False).head(5)
            cols = [c for c in ["风险等级", "风险得分", "异常依据"] if c in top_anom.columns]
            add_dataframe_to_docx(doc, top_anom[cols], max_rows=5)
    else:
        add_para(doc, "当前数据暂未形成稳定的异常诊断结果。", indent=True)

    add_heading(doc, "五、AI增强解读与管理建议", 1)
    add_heading(doc, "5.1 AI综合判断", 2)
    if ai_paras:
        for p in ai_paras[:4]:
            add_para(doc, p, indent=True)
    else:
        add_para(doc, "系统已完成经营指标、维度结构和风险结果的综合分析。后续应结合业务背景，对高贡献维度和高风险经营单元进行重点复核。", indent=True)

    add_heading(doc, "5.2 重点问题归纳", 2)
    if anomaly_df is not None and len(anomaly_df) and '是否经营异常' in anomaly_df.columns:
        abnormal = anomaly_df[anomaly_df['是否经营异常']]
        add_para(doc, f"重点问题主要集中在异常经营单元识别和高风险对象复核方面。当前异常单元数量为 {len(abnormal)}，建议优先核查风险得分较高、异常依据较多的对象。", indent=True)
    else:
        add_para(doc, "当前未识别出明显高风险对象，但仍建议定期跟踪主指标趋势和维度结构变化。", indent=True)

    add_heading(doc, "5.3 后续核查与管理建议", 2)
    if ai_tables:
        for _, tb in ai_tables[:2]:
            add_dataframe_to_docx(doc, tb, max_rows=12)
    else:
        add_default_action_table(doc, main_metric, dimensions)
    add_para(doc, "以上建议用于辅助管理层确定核查优先级。实际决策仍需结合企业业务背景、政策制度、预算目标和原始业务单据进行综合判断。", indent=True)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ============================================================
# 3. 异常诊断
# ============================================================

def anomaly_detection(df, main_metric, numeric_cols, dimensions, date_col=None):
    agg_cols = [c for c in numeric_cols if c in df.columns and c != "" and not is_id_like(c)]
    if len(agg_cols) == 0 or main_metric not in df.columns:
        return pd.DataFrame()

    temp = df.copy()
    group_cols = []
    if date_col:
        temp = periodize(temp, date_col)
        if not temp.empty:
            group_cols.append("_period")
    for d in dimensions[:2]:
        if d in temp.columns:
            group_cols.append(d)

    if group_cols:
        unit = temp.groupby(group_cols, dropna=False).agg({c: "sum" for c in agg_cols}).reset_index()
        unit["记录数"] = temp.groupby(group_cols, dropna=False).size().values
    else:
        unit = temp[agg_cols].copy()
        unit["记录数"] = 1

    model_cols = [c for c in agg_cols if c in unit.columns and unit[c].notna().sum() > 0]
    if len(model_cols) == 0:
        return pd.DataFrame()

    X_df = unit[model_cols].replace([np.inf, -np.inf], np.nan)
    X_df = X_df.fillna(X_df.median(numeric_only=True))

    if len(unit) >= 8 and len(model_cols) >= 1:
        scaler = StandardScaler()
        X = scaler.fit_transform(X_df)
        contamination = min(0.12, max(0.04, 8 / max(len(unit), 1)))
        model = IsolationForest(n_estimators=160, contamination=contamination, random_state=42)
        labels = model.fit_predict(X)
        raw = -model.decision_function(X)
        unit["是否AI异常"] = labels == -1
        unit["异常得分"] = 0 if raw.max() == raw.min() else (raw - raw.min()) / (raw.max() - raw.min()) * 100
    else:
        unit["是否AI异常"] = False
        unit["异常得分"] = 0.0

    q95 = X_df.quantile(0.95)
    q05 = X_df.quantile(0.05)

    evidence_list = []
    main_component = []
    combo_component = []
    rule_component = []

    for _, row in unit.iterrows():
        evidences = []
        mc = cc = rc = 0

        if main_metric in unit.columns:
            val = row[main_metric]
            pct_rank = (unit[main_metric] <= val).mean()
            if pct_rank >= 0.95:
                evidences.append(f"{main_metric}处于高位区间，高于约{pct_rank:.0%}的经营单元")
                mc = 30
            elif pct_rank <= 0.05:
                evidences.append(f"{main_metric}处于低位区间，低于约{(1-pct_rank):.0%}的经营单元")
                mc = 20

        extreme_cols = []
        for c in model_cols:
            val = row[c]
            if val >= q95[c] or val <= q05[c]:
                extreme_cols.append(c)
        if len(extreme_cols) >= 2:
            evidences.append(f"多项指标同时处于极端区间：{', '.join(extreme_cols[:4])}")
            cc = min(10 * len(extreme_cols), 35)
        elif len(extreme_cols) == 1:
            evidences.append(f"{extreme_cols[0]}处于极端区间")
            cc = 12

        for c in model_cols:
            cname = c.lower()
            val = row[c]
            if ("profit" in cname or "利润" in cname) and val < 0:
                evidences.append(f"{c}为负，存在收益或盈利风险")
                rc += 20
            if is_rate_like(c) and (val > 1.2 or val < -0.2):
                evidences.append(f"{c}超出常见比例范围，需要复核口径")
                rc += 10
            if ("cost" in cname or "费用" in cname or "成本" in cname or "expense" in cname) and val >= q95[c]:
                evidences.append(f"{c}处于高位区间，提示资源投入或支出压力较高")
                rc += 12

        if not evidences and row.get("是否AI异常", False):
            evidences.append("多指标组合与整体样本差异较大，被模型识别为综合异常")
        if not evidences:
            evidences.append("未识别出显著异常依据")

        evidence_list.append("；".join(evidences))
        main_component.append(mc)
        combo_component.append(cc)
        rule_component.append(min(rc, 25))

    unit["异常依据"] = evidence_list
    unit["主指标偏离"] = main_component
    unit["多指标组合偏离"] = combo_component
    unit["业务规则风险"] = rule_component
    unit["模型异常贡献"] = np.where(unit["是否AI异常"], np.minimum(unit["异常得分"] * 0.35, 35), np.minimum(unit["异常得分"] * 0.15, 15))
    unit["风险得分"] = (unit["主指标偏离"] + unit["多指标组合偏离"] + unit["业务规则风险"] + unit["模型异常贡献"]).clip(0, 100).round(1)

    def level(score):
        if score >= 70:
            return "高风险"
        if score >= 45:
            return "中风险"
        if score > 0:
            return "低风险"
        return "正常"

    unit["风险等级"] = unit["风险得分"].apply(level)
    unit["是否经营异常"] = unit["风险得分"] >= 45
    return unit.sort_values("风险得分", ascending=False)


def radar_chart(comp):
    theta = comp["风险来源"].tolist()
    r = comp["得分"].tolist()
    if theta:
        theta = theta + [theta[0]]
        r = r + [r[0]]
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(r=r, theta=theta, fill="toself", name="风险组成"))
    fig.update_layout(
        title="风险雷达图",
        polar=dict(radialaxis=dict(visible=True, range=[0, max(40, max(r) if r else 40)])),
        showlegend=False
    )
    st.plotly_chart(chart_layout(fig, 380), use_container_width=True)


# ============================================================
# 4. 问数助手
# ============================================================

def get_secret_value(name):
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.getenv(name, "")


def get_llm_configs():
    return {
        "通义千问": {
            "api_key": get_secret_value("QWEN_API_KEY"),
            "base_url": get_secret_value("QWEN_BASE_URL") or "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
            "model": get_secret_value("QWEN_MODEL") or "qwen-plus"
        },
        "DeepSeek": {
            "api_key": get_secret_value("DEEPSEEK_API_KEY"),
            "base_url": get_secret_value("DEEPSEEK_BASE_URL") or "https://api.deepseek.com/chat/completions",
            "model": get_secret_value("DEEPSEEK_MODEL") or "deepseek-chat"
        },
        "智谱GLM": {
            "api_key": get_secret_value("ZHIPU_API_KEY"),
            "base_url": get_secret_value("ZHIPU_BASE_URL") or "https://open.bigmodel.cn/api/paas/v4/chat/completions",
            "model": get_secret_value("ZHIPU_MODEL") or "glm-4-flash"
        }
    }


def call_llm(provider_name, system_prompt, user_prompt, temperature=0.1):
    cfg = get_llm_configs()[provider_name]
    if not cfg["api_key"]:
        raise ValueError(f"{provider_name} 未配置 API Key。请在 .streamlit/secrets.toml 中配置。")
    headers = {"Authorization": f"Bearer {cfg['api_key']}", "Content-Type": "application/json"}
    payload = {
        "model": cfg["model"],
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": temperature
    }
    start = time.time()
    resp = requests.post(cfg["base_url"], headers=headers, json=payload, timeout=60)
    elapsed = time.time() - start
    if resp.status_code != 200:
        raise RuntimeError(f"API 请求失败：{resp.status_code}｜{resp.text[:500]}")
    data = resp.json()
    return data["choices"][0]["message"]["content"], elapsed


def to_sqlite_df(df):
    sql_df = df.copy()
    sql_df.columns = [normalize_column_name(c) for c in sql_df.columns]
    for col in sql_df.columns:
        if pd.api.types.is_datetime64_any_dtype(sql_df[col]):
            sql_df[col] = sql_df[col].dt.strftime("%Y-%m-%d")
    return sql_df


def build_schema_text(df, meta=None):
    lines = []
    for col in df.columns:
        dtype = str(df[col].dtype)
        samples = df[col].dropna().astype(str).head(3).tolist()
        meaning = ""
        if meta is not None and col in meta["字段名"].values:
            meaning = meta.loc[meta["字段名"] == col, "字段含义解释"].iloc[0]
        lines.append(f"- {col}｜类型：{dtype}｜含义：{meaning}｜示例：{samples}")
    return "\n".join(lines)


def extract_sql(text):
    if not text:
        return ""
    text = text.strip()
    m = re.search(r"```sql(.*?)```", text, flags=re.S | re.I)
    if m:
        return m.group(1).strip()
    m = re.search(r"```(.*?)```", text, flags=re.S)
    if m:
        return m.group(1).strip()
    m = re.search(r"(SELECT\s+.*)", text, flags=re.S | re.I)
    if m:
        sql = m.group(1).strip()
        if ";" in sql:
            sql = sql.split(";")[0] + ";"
        return sql
    return text


def is_safe_select_sql(sql):
    if not sql:
        return False, "SQL 为空"
    low = sql.strip().strip(";").lower()
    if not low.startswith("select"):
        return False, "只允许 SELECT 查询"
    banned = ["insert", "update", "delete", "drop", "alter", "create", "truncate", "replace", "attach", "detach", "pragma"]
    if any(re.search(rf"\b{b}\b", low) for b in banned):
        return False, "检测到非查询类 SQL，已拦截"
    return True, ""


def score_llm_query(ok, result_df, sql, explanation, elapsed, columns):
    score = 0
    detail = {}

    detail["SQL可执行性"] = 40 if ok else 0
    score += detail["SQL可执行性"]

    if ok and isinstance(result_df, pd.DataFrame) and len(result_df) > 0:
        detail["查询结果有效性"] = 20
    elif ok:
        detail["查询结果有效性"] = 8
    else:
        detail["查询结果有效性"] = 0
    score += detail["查询结果有效性"]

    used_cols = [c for c in columns if str(c).lower() in str(sql).lower()]
    detail["字段匹配度"] = 15 if used_cols else (5 if ok else 0)
    score += detail["字段匹配度"]

    detail["结果解释质量"] = 15 if explanation and len(str(explanation)) > 40 else (6 if explanation else 0)
    score += detail["结果解释质量"]

    if pd.isna(elapsed):
        speed = 0
    elif elapsed <= 3:
        speed = 10
    elif elapsed <= 8:
        speed = 8
    elif elapsed <= 15:
        speed = 5
    else:
        speed = 2
    detail["响应速度"] = speed
    score += speed

    return min(int(score), 100), detail


def run_llm_sql_question(provider, question, df, meta):
    sql_df = to_sqlite_df(df)
    con = sqlite3.connect(":memory:")
    sql_df.to_sql("data", con, index=False, if_exists="replace")
    schema = build_schema_text(sql_df, meta=None)

    system_prompt = """
你是一个严谨的数据分析 SQL 助手。你的任务是把用户的自然语言问题转换为 SQLite SELECT 查询语句。
要求：
1. 只输出 SQL，不要解释；
2. 表名固定为 data；
3. 只能使用 SELECT；
4. 字段名必须严格来自字段信息，中文或特殊字段名请使用双引号；
5. 尽量使用 LIMIT 20 控制输出；
6. 若问题涉及排名、最高、最低，请使用 ORDER BY；
7. 若问题涉及分组统计，请使用 GROUP BY。
"""
    user_prompt = f"字段信息：\n{schema}\n\n用户问题：{question}\n\n请生成 SQLite SELECT 查询语句。"
    content, elapsed = call_llm(provider, system_prompt, user_prompt, temperature=0)
    sql = extract_sql(content)

    safe, reason = is_safe_select_sql(sql)
    if not safe:
        score, detail = score_llm_query(False, pd.DataFrame(), sql, "", elapsed, sql_df.columns)
        return {"模型": provider, "SQL": sql, "是否成功": False, "错误信息": reason, "结果": pd.DataFrame(), "解释": "", "响应时间": elapsed, "综合评分": score, "评分明细": detail}

    try:
        result = pd.read_sql_query(sql, con)
        ok, error = True, ""
    except Exception as e:
        result = pd.DataFrame()
        ok, error = False, str(e)

    explanation = ""
    if ok:
        try:
            exp_system = "你是一名经营数据分析顾问。请根据查询结果，用中文简洁解释结论、经营含义和建议。"
            exp_user = f"用户问题：{question}\nSQL：{sql}\n查询结果：\n{result.head(20).to_markdown(index=False)}\n\n请输出：1. 查询结论；2. 经营含义；3. 后续建议。"
            explanation, _ = call_llm(provider, exp_system, exp_user, temperature=0.2)
        except Exception as e:
            explanation = f"SQL 已执行成功，但结果解释生成失败：{e}"

    score, detail = score_llm_query(ok, result, sql, explanation, elapsed, sql_df.columns)
    return {"模型": provider, "SQL": sql, "是否成功": ok, "错误信息": error, "结果": result, "解释": explanation, "响应时间": elapsed, "综合评分": score, "评分明细": detail}


def generate_dynamic_questions(main_metric, dimensions, numeric_cols, date_col):
    qs = []
    d1 = dimensions[0] if dimensions else None
    d2 = dimensions[1] if len(dimensions) > 1 else None
    other = next((c for c in numeric_cols if c != main_metric and not is_id_like(c)), None)

    if d1:
        qs.append(("排名类", f"按 {d1} 统计 {main_metric} 最高的前10项"))
        qs.append(("对比类", f"比较不同 {d1} 的 {main_metric} 均值差异"))
        qs.append(("结构类", f"各 {d1} 的 {main_metric} 占比分别是多少"))
    if d2:
        qs.append(("交叉分析类", f"按 {d1} 和 {d2} 交叉统计 {main_metric}"))
    if date_col:
        qs.append(("趋势类", f"按月份统计 {main_metric} 的变化趋势"))
        qs.append(("波动类", f"找出 {main_metric} 变化最大的月份"))
    if other:
        qs.append(("关系类", f"分析 {other} 与 {main_metric} 是否存在关系"))
    qs.append(("异常类", f"查询 {main_metric} 异常偏高或偏低的记录"))
    qs.append(("风险类", "当前风险最高的经营单元有哪些"))
    return qs


# ============================================================
# 5. Word 简报
# ============================================================

def set_run_font(run, font="宋体", size=11, bold=False):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    run.font.bold = bold


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_run_font(run, "黑体", 15 if level == 1 else 13, True)
    return p


def add_para(doc, text, indent=True):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    if indent:
        p.paragraph_format.first_line_indent = Pt(24)
    run = p.add_run(str(text))
    set_run_font(run, "宋体", 11, False)
    return p





# ---------- Word 图表导出工具：用于管理层简报中的真实可视化 ----------
def _pick_mpl_chinese_font():
    """尽量选择可用中文字体，避免 Word 图表中文乱码；找不到时仍可生成图表，并在正文中解释。"""
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf", r"C:\Windows\Fonts\simsun.ttc",
        "/System/Library/Fonts/PingFang.ttc", "/System/Library/Fonts/STHeiti Light.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/arphic/ukai.ttc",
    ]
    for fp in candidates:
        if os.path.exists(fp):
            try:
                from matplotlib import font_manager
                font_manager.fontManager.addfont(fp)
                return font_manager.FontProperties(fname=fp).get_name()
            except Exception:
                continue
    return None


def _setup_mpl_style():
    font_name = _pick_mpl_chinese_font()
    if font_name:
        plt.rcParams["font.sans-serif"] = [font_name, "Microsoft YaHei", "SimHei", "Arial Unicode MS", "DejaVu Sans"]
    else:
        plt.rcParams["font.sans-serif"] = ["DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.facecolor"] = "white"
    plt.rcParams["axes.facecolor"] = "white"


def _short_label(x, max_len=14):
    s = str(x)
    return s if len(s) <= max_len else s[:max_len] + "…"


def _mpl_money_axis(x, pos=None):
    """Matplotlib 图内尽量使用 ASCII 数值后缀，避免云端环境缺少中文字体时出现方框。"""
    try:
        x = float(x)
    except Exception:
        return str(x)
    sign = "-" if x < 0 else ""
    ax = abs(x)
    if ax >= 1_000_000_000:
        return f"{sign}{ax/1_000_000_000:.1f}B"
    if ax >= 1_000_000:
        return f"{sign}{ax/1_000_000:.1f}M"
    if ax >= 1_000:
        return f"{sign}{ax/1_000:.0f}K"
    return f"{x:.0f}"


def _mpl_value_label(x):
    """Matplotlib 图内数据标签，避免使用“万/亿”等中文单位。"""
    try:
        x = float(x)
    except Exception:
        return str(x)
    sign = "-" if x < 0 else ""
    ax = abs(x)
    if ax >= 1_000_000_000:
        return f"{sign}{ax/1_000_000_000:.2f}B"
    if ax >= 1_000_000:
        return f"{sign}{ax/1_000_000:.2f}M"
    if ax >= 1_000:
        return f"{sign}{ax/1_000:.1f}K"
    return f"{x:.2f}"

def add_mpl_figure_to_docx(doc, fig, caption):
    """把 Matplotlib 图表以 PNG 插入 Word；若图中使用 C1/R1 等编码，自动补充中文映射表。"""
    try:
        img = BytesIO()
        fig.savefig(img, format="png", dpi=180, bbox_inches="tight", facecolor="white")
        mapping = getattr(fig, "_zc_mapping", None)
        plt.close(fig)
        img.seek(0)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(caption)
        set_run_font(run, "黑体", 10, True)
        doc.add_picture(img, width=Inches(6.2))
        if mapping:
            map_df = pd.DataFrame(mapping)
            if len(map_df):
                p2 = doc.add_paragraph()
                run2 = p2.add_run("图中编码说明：")
                set_run_font(run2, "黑体", 10, True)
                add_dataframe_to_docx(doc, map_df, max_rows=20)
        return True
    except Exception as e:
        try:
            plt.close(fig)
        except Exception:
            pass
        add_para(doc, f"图表“{caption}”生成失败：{e}。", indent=False)
        return False

def make_mpl_trend_figure(trend, metric):
    _setup_mpl_style()
    fig, ax = plt.subplots(figsize=(8.8, 4.6))
    plot = trend.copy()
    plot[metric] = pd.to_numeric(plot[metric], errors="coerce")
    plot = plot.dropna(subset=["期间", metric])
    xs = list(range(len(plot)))
    ys = plot[metric].astype(float).tolist()
    ax.plot(xs, ys, linewidth=2.8, marker="o", markersize=4.8)
    ax.fill_between(xs, ys, alpha=0.10)
    if len(xs) > 0:
        max_i = int(np.nanargmax(ys)); min_i = int(np.nanargmin(ys))
        ax.scatter([max_i], [ys[max_i]], s=78, zorder=4)
        ax.scatter([min_i], [ys[min_i]], s=78, zorder=4)
        ax.annotate(f"Max {_mpl_value_label(ys[max_i])}", (max_i, ys[max_i]), xytext=(0, 12), textcoords="offset points", ha="center", fontsize=9)
        ax.annotate(f"Min {_mpl_value_label(ys[min_i])}", (min_i, ys[min_i]), xytext=(0, -18), textcoords="offset points", ha="center", fontsize=9)
    periods = plot["期间"].astype(str).tolist()
    if len(periods) > 8:
        step = max(1, int(np.ceil(len(periods) / 8)))
        tick_idx = list(range(0, len(periods), step))
        if len(periods)-1 not in tick_idx:
            tick_idx.append(len(periods)-1)
    else:
        tick_idx = list(range(len(periods)))
    ax.set_xticks(tick_idx)
    ax.set_xticklabels([periods[i] for i in tick_idx], rotation=0, fontsize=9)
    ax.yaxis.set_major_formatter(FuncFormatter(_mpl_money_axis))
    ax.set_title("Trend of Key Metric", fontsize=14, fontweight="bold", loc="left")
    ax.set_xlabel("Period")
    ax.set_ylabel("Value")
    ax.grid(axis="y", alpha=0.25)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return fig

def make_mpl_dimension_bar_figure(df, dim, metric, topn=10):
    _setup_mpl_style()
    g = dimension_summary(df, dim, metric).head(topn).copy()
    g[f"{metric}合计"] = pd.to_numeric(g[f"{metric}合计"], errors="coerce").fillna(0)
    g = g.sort_values(f"{metric}合计", ascending=True)
    fig, ax = plt.subplots(figsize=(8.8, max(4.2, 0.42 * len(g) + 1.2)))
    raw_labels = g[dim].astype(str).tolist()
    code_labels = [f"C{i+1}" for i in range(len(raw_labels))]
    vals = g[f"{metric}合计"].astype(float).tolist()
    bars = ax.barh(code_labels, vals, height=0.62)
    ax.xaxis.set_major_formatter(FuncFormatter(_mpl_money_axis))
    ax.set_title("Dimension Summary", fontsize=14, fontweight="bold", loc="left")
    ax.set_xlabel("Total Value")
    ax.set_ylabel("Category Code")
    ax.grid(axis="x", alpha=0.22)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    max_val = max([abs(v) for v in vals] + [1])
    for bar, val in zip(bars, vals):
        ax.text(val + max_val * 0.015, bar.get_y() + bar.get_height()/2, _mpl_value_label(val), va="center", fontsize=9)
    fig._zc_mapping = [{"图中编码": code, "对应维度项": raw} for code, raw in zip(code_labels, raw_labels)]
    fig.tight_layout()
    return fig

def make_mpl_distribution_figure(df, metric):
    _setup_mpl_style()
    s = to_finite_numeric_series(df, metric)
    fig, ax = plt.subplots(figsize=(8.8, 4.4))
    if len(s) > 0:
        ax.hist(s.astype(float).tolist(), bins=min(30, max(8, int(np.sqrt(len(s))))), alpha=0.86)
        ax.axvline(float(s.mean()), linestyle="--", linewidth=2, label="Mean")
        ax.legend(fontsize=9)
    ax.xaxis.set_major_formatter(FuncFormatter(_mpl_money_axis))
    ax.set_title("Metric Distribution", fontsize=14, fontweight="bold", loc="left")
    ax.set_xlabel("Value")
    ax.set_ylabel("Count")
    ax.grid(axis="y", alpha=0.22)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return fig

def make_mpl_risk_figure(anomaly_df):
    _setup_mpl_style()
    fig, ax = plt.subplots(figsize=(8.0, 4.2))
    mapping = []
    if anomaly_df is not None and len(anomaly_df) and "风险等级" in anomaly_df.columns:
        order = ["高风险", "中风险", "低风险", "正常"]
        counts = anomaly_df["风险等级"].value_counts()
        labels = [x for x in order if x in counts.index] + [x for x in counts.index if x not in order]
        vals = [int(counts.get(x, 0)) for x in labels]
        code_labels = [f"R{i+1}" for i in range(len(labels))]
        bars = ax.bar(code_labels, vals, width=0.55)
        mapping = [{"图中编码": code, "风险等级": label} for code, label in zip(code_labels, labels)]
        for bar, val in zip(bars, vals):
            ax.text(bar.get_x() + bar.get_width()/2, val + max(vals + [1])*0.02, str(val), ha="center", fontsize=10)
    ax.set_title("Risk Level Distribution", fontsize=14, fontweight="bold", loc="left")
    ax.set_xlabel("Risk Code")
    ax.set_ylabel("Count")
    ax.grid(axis="y", alpha=0.22)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    if mapping:
        fig._zc_mapping = mapping
    fig.tight_layout()
    return fig

def add_core_visuals_to_docx(doc, df, main_metric, dimensions, date_col, anomaly_df):
    """在 Word 简报中插入核心可视化：趋势/分布、维度结构、风险分布。"""
    inserted = 0
    if date_col:
        trend = build_trend(df, date_col, main_metric)
        if trend is not None and len(trend) >= 2:
            add_mpl_figure_to_docx(doc, make_mpl_trend_figure(trend, main_metric), f"图1 {main_metric}时间趋势")
            inserted += 1
    if inserted == 0:
        add_mpl_figure_to_docx(doc, make_mpl_distribution_figure(df, main_metric), f"图1 {main_metric}分布情况")
        inserted += 1
    if dimensions:
        add_mpl_figure_to_docx(doc, make_mpl_dimension_bar_figure(df, dimensions[0], main_metric), f"图2 按{dimensions[0]}汇总{main_metric}")
        inserted += 1
    if anomaly_df is not None and len(anomaly_df):
        add_mpl_figure_to_docx(doc, make_mpl_risk_figure(anomaly_df), "图3 风险等级分布")
        inserted += 1
    return inserted

def generate_report_docx(df, main_metric, dimensions, date_col, anomaly_df, focus_list):
    if not DOCX_AVAILABLE:
        raise RuntimeError("未安装 python-docx，请先安装：python -m pip install python-docx")
    doc = Document()
    title = doc.add_paragraph(); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("智策经营——管理层决策简报"); set_run_font(r, "黑体", 20, True)
    subtitle = doc.add_paragraph(); subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = subtitle.add_run("AI 驱动的多维经营分析与决策支持系统自动生成"); set_run_font(r, "宋体", 11, False); r.font.color.rgb = RGBColor(90,90,90)

    add_heading(doc,"一、分析概况",1)
    add_heading(doc,"1.1 数据范围与分析口径",2)
    add_para(doc, f"本报告基于用户上传的数据自动生成。当前清洗后数据共包含 {len(df):,} 条记录，主分析指标为“{main_metric}”。系统根据数据中实际存在的字段进行分析，不会引用未上传或不存在的指标。")
    if focus_list:
        add_para(doc, f"本次简报关注重点包括：{'、'.join(focus_list)}。")

    add_heading(doc,"1.2 核心指标摘要",2)
    table=doc.add_table(rows=1, cols=2); table.alignment=WD_TABLE_ALIGNMENT.CENTER; table.style="Table Grid"
    table.rows[0].cells[0].text='指标'; table.rows[0].cells[1].text='数值'
    abnormal_n = int((anomaly_df['是否经营异常']).sum()) if anomaly_df is not None and len(anomaly_df) else 0
    for k,v in [(f'{main_metric}合计',money_fmt(df[main_metric].sum())),(f'{main_metric}均值',money_fmt(df[main_metric].mean())),(f'{main_metric}最大值',money_fmt(df[main_metric].max())),(f'{main_metric}最小值',money_fmt(df[main_metric].min())),('数据记录数',f'{len(df):,}'),('异常经营单元数',str(abnormal_n))]:
        cells=table.add_row().cells; cells[0].text=k; cells[1].text=v

    add_heading(doc,"二、核心经营表现",1)
    add_heading(doc,"2.1 总体表现",2)
    add_para(doc, f"当前样本中，“{main_metric}”合计为 {money_fmt(df[main_metric].sum())}，均值为 {money_fmt(df[main_metric].mean())}，最大值为 {money_fmt(df[main_metric].max())}，最小值为 {money_fmt(df[main_metric].min())}。均值反映一般经营水平，最大值和最小值可帮助定位高值或低值经营单元。")

    if date_col:
        trend=build_trend(df,date_col,main_metric)
        add_heading(doc,"2.2 时间趋势分析",2)
        add_para(doc,trend_interpretation(trend,main_metric))
        if len(trend)>=2:
            add_mpl_figure_to_docx(doc, make_mpl_trend_figure(trend, main_metric), f"图1 {main_metric}时间趋势")
        if len(trend)>=3:
            recent=trend.tail(3)
            add_para(doc, f"最近三个周期分别为 {', '.join(recent['期间'].astype(str).tolist())}，对应数值为 {', '.join([money_fmt(v) for v in recent[main_metric].tolist()])}。建议重点关注最近一期相较前期的变化方向，以及是否与长期趋势一致。")
    else:
        add_heading(doc,"2.2 结构分布分析",2)
        add_para(doc, f"当前未选择有效日期字段，系统采用结构分布方式观察“{main_metric}”的集中程度和离散情况。")
        add_mpl_figure_to_docx(doc, make_mpl_distribution_figure(df, main_metric), f"图1 {main_metric}分布情况")

    if dimensions:
        add_heading(doc,"2.3 多维结构分析",2)
        add_mpl_figure_to_docx(doc, make_mpl_dimension_bar_figure(df, dimensions[0], main_metric), f"图2 按{dimensions[0]}汇总{main_metric}")
        for dim in dimensions[:3]:
            g=dimension_summary(df,dim,main_metric); top=g.iloc[0]
            add_para(doc, f"在“{dim}”维度下，{top[dim]} 的 {main_metric} 表现最高，合计为 {money_fmt(top[f'{main_metric}合计'])}，记录数为 {int(top['记录数'])}。如该维度项长期占比过高，应判断是正常规模优势，还是存在资源集中、费用集中或结构失衡。")

    add_heading(doc,"三、AI 异常诊断",1)
    if anomaly_df is not None and len(anomaly_df):
        abnormal=anomaly_df[anomaly_df['是否经营异常']].copy(); high=abnormal[abnormal['风险等级']=='高风险']
        add_para(doc, f"系统基于主指标偏离、多指标组合偏离、动态业务规则和模型异常贡献识别异常经营单元。当前共识别出 {len(abnormal)} 个异常经营单元，其中高风险单元 {len(high)} 个。风险得分越高，说明该经营单元越需要优先核查。")
        add_mpl_figure_to_docx(doc, make_mpl_risk_figure(anomaly_df), "图3 风险等级分布")
        top_anom=abnormal.sort_values('风险得分',ascending=False).head(5)
        if len(top_anom):
            add_heading(doc,"3.1 重点异常单元",2)
            # 用表格展示 Top 风险对象，提升可读性
            risk_table = top_anom[[c for c in ["风险等级", "风险得分", "异常依据"] if c in top_anom.columns]].copy()
            add_dataframe_to_docx(doc, risk_table, max_rows=5)
            add_heading(doc,"3.2 异常原因归纳",2)
            add_para(doc,"高风险通常不是由单一数值造成，而是由主指标处于高位或低位、多项指标同时偏离、以及数据中实际存在的业务规则风险叠加形成。建议结合维度、相关指标和业务背景进行交叉核查。")
    else:
        add_para(doc,"当前数据不足以进行稳定的异常诊断。")

    add_heading(doc,"四、管理建议",1)
    if dimensions:
        add_para(doc, f"第一，建议优先围绕“{dimensions[0]}”等关键维度开展下钻分析，重点关注主指标贡献较高或波动较大的维度项，判断其变化是否来自业务规模扩大、资源投入增加，还是管理效率下降。")
    add_para(doc, f"第二，针对“{main_metric}”处于高位的经营单元，建议复核对应业务单据、审批流程、资源投入和预算执行情况，识别是否存在集中支出、异常消耗、重复记录或口径不一致。")
    add_para(doc,"第三，针对多指标同时偏离的经营单元，建议开展交叉检查：一方面核查原始数据填报是否准确，另一方面结合相关数值指标判断是否存在投入产出不匹配、成本结构异常或业务流程异常。")
    add_para(doc,"第四，若系统识别出某些维度项长期占比过高，建议设置后续监控阈值，并定期输出同口径报表，以便判断该问题是短期波动还是持续性结构问题。")

    add_heading(doc,"五、后续关注重点",1)
    add_para(doc,"后续建议持续关注主指标变化、维度结构变化、异常经营单元数量及高风险单元变化。如果接入预算、目标、行业基准或更完整的业务字段，系统可进一步提升异常解释和决策建议的精度。")
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio

# ============================================================
# 6. 页面主体
# ============================================================

st.markdown("""
<div class="hero">
    <h1>智策经营——AI 驱动的多维经营分析与决策支持系统</h1>
    <p>先理解字段，再自适应分析：系统基于上传数据自动完成字段解释、经营态势、经营洞察、异常诊断、问数助手与管理层简报生成。</p>
</div>
""", unsafe_allow_html=True)

st.sidebar.title("📁 数据与参数")
uploaded_file = st.sidebar.file_uploader("上传经营数据文件", type=["xlsx", "xls", "csv"])

st.sidebar.markdown("---")
st.sidebar.markdown("### 🧭 功能导航")
st.sidebar.markdown("""
- 🧩 数据治理  
- 🏠 经营态势  
- 📊 经营洞察  
- ⚠️ 风险识别  
- 💬 问数助手  
- 📝 决策简报  
""")

if uploaded_file is None:
    st.info("请先在左侧上传 Excel 或 CSV 文件。")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="section-card"><h3>🧩 先理解字段</h3><p class="small-text">自动识别字段含义、角色、聚合方式和注意事项。</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="section-card"><h3>🤖 再智能诊断</h3><p class="small-text">基于实际存在字段判断异常，并解释异常原因与风险来源。</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="section-card"><h3>📝 最后生成简报</h3><p class="small-text">输出结构完整、格式规范的管理层决策简报。</p></div>', unsafe_allow_html=True)
    st.stop()

try:
    raw_df = read_uploaded_file(uploaded_file)
except Exception as e:
    st.error(f"文件读取失败：{e}")
    st.stop()

raw_df.columns = [normalize_column_name(c) for c in raw_df.columns]
meta = infer_field_metadata(raw_df)

metric_candidates = get_metric_candidates(meta)
date_candidates = get_date_candidates(meta)
dimension_candidates = get_dimension_candidates(meta)

if not metric_candidates:
    st.error("系统未识别到适合作为主指标的数值字段，请检查数据格式。")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ 字段配置")

main_metric = st.sidebar.selectbox("主分析指标", options=metric_candidates, index=0)
date_choice = st.sidebar.selectbox("日期字段", options=["无"] + date_candidates, index=0)
date_col = None if date_choice == "无" else date_choice

default_nums = list(dict.fromkeys([main_metric] + [c for c in metric_candidates if c != main_metric][:4]))
selected_numeric_cols = st.sidebar.multiselect("参与分析的数值字段", options=metric_candidates, default=default_nums)
selected_dimensions = st.sidebar.multiselect("分析维度字段", options=dimension_candidates, default=dimension_candidates[:3])
missing_strategy = st.sidebar.selectbox("缺失值处理策略", options=["保留并在分析时忽略", "删除主指标缺失行", "数值中位数填充，类别填未知"], index=0)

# 确保主指标一定被转换
if main_metric not in selected_numeric_cols:
    selected_numeric_cols = [main_metric] + selected_numeric_cols

df, quality_report, clean_summary, numeric_report = clean_data(raw_df, selected_numeric_cols, date_col, missing_strategy)

valid_date_col = None
if date_col:
    parsed, ratio, valid, reason = try_parse_date(df[date_col])
    if valid:
        valid_date_col = date_col
        df[date_col] = parsed
    else:
        st.sidebar.warning(f"日期字段“{date_col}”暂不作为有效日期使用：{reason}")

numeric_cols = safe_numeric_cols(df, selected_numeric_cols)

if main_metric not in numeric_cols:
    st.error(f"主指标“{main_metric}”没有可用数值。请在左侧重新选择主指标，或在“数据治理”中查看数值转换报告。")
    st.stop()

anomaly_df = anomaly_detection(df, main_metric, numeric_cols, selected_dimensions, valid_date_col)



# ============================================================
# 数据治理可视化辅助函数
# ============================================================

def compute_data_health_score(clean_summary, meta, numeric_report, metric_candidates, dimension_candidates, date_candidates):
    rows = max(clean_summary.get("清洗后行数", 0), 1)
    cols = max(clean_summary.get("清洗后字段数", 0), 1)
    missing = clean_summary.get("缺失单元格数", 0)
    dup = clean_summary.get("重复行数", 0)

    missing_rate = min(missing / max(rows * cols, 1), 1)
    dup_rate = min(dup / max(clean_summary.get("原始行数", rows), 1), 1)

    missing_score = max(0, 100 - missing_rate * 220)
    dup_score = max(0, 100 - dup_rate * 260)

    if numeric_report is not None and len(numeric_report) > 0 and "转换后有效数值数" in numeric_report.columns and "转换前非空数" in numeric_report.columns:
        valid_rates = []
        for _, r in numeric_report.iterrows():
            before = float(r.get("转换前非空数", 0))
            after = float(r.get("转换后有效数值数", 0))
            valid_rates.append(after / before if before > 0 else 1)
        numeric_score = np.mean(valid_rates) * 100 if valid_rates else 80
    else:
        numeric_score = 80

    structure_score = 100 if metric_candidates and dimension_candidates else (72 if metric_candidates else 45)
    date_score = 100 if date_candidates else 82

    score = 0.28 * missing_score + 0.18 * dup_score + 0.24 * numeric_score + 0.20 * structure_score + 0.10 * date_score
    return int(round(max(0, min(score, 100))))


def health_level(score):
    if score >= 88:
        return "优秀", "数据结构较完整，适合直接开展经营分析。"
    if score >= 75:
        return "良好", "数据整体可用，建议关注少量质量问题。"
    if score >= 60:
        return "可用", "数据可以分析，但建议先处理缺失、重复或字段口径问题。"
    return "待治理", "数据质量或字段结构存在明显不足，建议先完成数据治理。"


def gauge_chart(score):
    level, desc = health_level(score)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=score,
        number={"suffix": "分", "font": {"size": 42}},
        title={"text": f"数据健康评分｜{level}", "font": {"size": 22}},
        gauge={
            "axis": {"range": [0, 100], "tickwidth": 1},
            "bar": {"thickness": 0.28},
            "steps": [
                {"range": [0, 60], "color": "#FCE8E6"},
                {"range": [60, 75], "color": "#FFF4D6"},
                {"range": [75, 88], "color": "#E8F3FF"},
                {"range": [88, 100], "color": "#E7F7EF"},
            ],
            "threshold": {"line": {"width": 4}, "thickness": 0.75, "value": score}
        }
    ))
    fig.update_layout(height=300, margin=dict(l=20, r=20, t=55, b=20), paper_bgcolor="#FFFFFF", font=dict(family="Microsoft YaHei"))
    st.plotly_chart(fig, use_container_width=True)
    return level, desc


def role_distribution_chart(meta):
    role_count = meta["推荐角色"].value_counts().reset_index()
    role_count.columns = ["字段角色", "字段数量"]
    order = ["度量指标", "数值字段", "维度字段", "时间字段", "进度/时长字段", "标识字段", "文本字段"]
    role_count["排序"] = role_count["字段角色"].apply(lambda x: order.index(x) if x in order else 99)
    role_count = role_count.sort_values(["排序", "字段数量"], ascending=[True, False])
    fig = go.Figure(go.Bar(
        x=role_count["字段数量"],
        y=role_count["字段角色"],
        orientation="h",
        text=role_count["字段数量"],
        textposition="outside",
        hovertemplate="%{y}<br>字段数量=%{x}<extra></extra>"
    ))
    fig.update_layout(title="字段角色分布", xaxis_title="字段数量", yaxis_title="字段角色")
    st.plotly_chart(chart_layout(fig, 500), use_container_width=True)


def data_quality_issue_chart(clean_summary, quality_report, numeric_report):
    duplicate = int(clean_summary.get("重复行数", 0))
    missing = int(clean_summary.get("缺失单元格数", 0))

    date_fail = 0
    if quality_report is not None and len(quality_report):
        q = quality_report[quality_report["处理环节"].astype(str).str.contains("日期字段解析", na=False)]
        if len(q):
            date_fail = int(q["处理数量"].sum())

    numeric_fail = 0
    if numeric_report is not None and len(numeric_report) and "缺失/无法解析数" in numeric_report.columns:
        numeric_fail = int(numeric_report["缺失/无法解析数"].sum())

    issues = pd.DataFrame({
        "问题类型": ["重复行", "缺失单元格", "日期解析失败", "数值缺失/无法解析"],
        "数量": [duplicate, missing, date_fail, numeric_fail]
    })
    fig = go.Figure(go.Bar(
        x=issues["数量"],
        y=issues["问题类型"],
        orientation="h",
        text=issues["数量"],
        textposition="outside",
        hovertemplate="%{y}<br>数量=%{x}<extra></extra>"
    ))
    fig.update_layout(title="数据质量问题概览", xaxis_title="数量", yaxis_title="问题类型")
    st.plotly_chart(chart_layout(fig, 500), use_container_width=True)


def make_pills(items, css_class=""):
    if not items:
        return '<span class="field-pill muted">暂无</span>'
    pills = []
    for item in items[:12]:
        pills.append(f'<span class="field-pill {css_class}">{item}</span>')
    if len(items) > 12:
        pills.append(f'<span class="field-pill muted">+{len(items)-12}</span>')
    return '<div class="pill-wrap">' + "".join(pills) + "</div>"


def field_recommendation_cards(meta, metric_candidates, dimension_candidates, date_candidates):
    id_fields = meta[meta["推荐角色"].isin(["标识字段"])]["字段名"].tolist()
    cautious_fields = meta[(meta["字段类型"].astype(str).str.contains("比例|比率", na=False)) | (meta["注意事项"].astype(str).str.contains("谨慎|不能求和|加权", na=False))]["字段名"].tolist()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="field-card">
            <div class="field-card-title">🎯 推荐主指标</div>
            <div class="field-card-desc">适合作为经营表现、规模、成本、收入或利润的核心分析对象。</div>
            {make_pills(metric_candidates)}
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="field-card">
            <div class="field-card-title">🧭 推荐分析维度</div>
            <div class="field-card-desc">适合用于分组对比、结构分析、贡献度分析和下钻定位。</div>
            {make_pills(dimension_candidates)}
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="field-card">
            <div class="field-card-title">⏱️ 时间字段</div>
            <div class="field-card-desc">存在时间字段时，可开展趋势、环比、阶段变化和贡献度分析。</div>
            {make_pills(date_candidates)}
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="field-card">
            <div class="field-card-title">⚠️ 谨慎使用</div>
            <div class="field-card-desc">ID类字段不应直接分析；折扣、比例类字段通常不适合求和。</div>
            {make_pills(list(dict.fromkeys(id_fields + cautious_fields)), "warn")}
        </div>
        """, unsafe_allow_html=True)


def numeric_validity_progress(numeric_report):
    if numeric_report is None or len(numeric_report) == 0:
        st.info("暂无数值转换检查结果。")
        return

    st.markdown("#### 数值字段有效率")
    st.markdown(
        "用于判断金额、数量、利润等字段是否能被系统正确识别为数值。"
        "有效率越高，后续图表、问数和风险识别结果越稳定。"
    )

    rows = []
    for _, r in numeric_report.iterrows():
        field = r.get("字段", "")
        before = float(r.get("转换前非空数", 0))
        after = float(r.get("转换后有效数值数", 0))
        rate = after / before if before > 0 else 1
        rows.append((field, max(0, min(rate, 1))))
    rows = sorted(rows, key=lambda x: x[1], reverse=True)

    for field, rate in rows[:12]:
        left, right = st.columns([5, 1])
        with left:
            st.markdown(f"**{field}**")
        with right:
            st.markdown(f"**{rate:.1%}**")
        st.progress(float(rate))

def chart_data_diagnostic(df, main_metric, numeric_cols, dimensions):
    rows = []
    check_cols = list(dict.fromkeys([main_metric] + list(numeric_cols)))
    for c in check_cols:
        if c not in df.columns:
            continue
        s = to_finite_numeric_series(df, c)
        rows.append({
            "字段": c,
            "有效数值数": int(len(s)),
            "唯一数值数": int(s.nunique(dropna=True)) if len(s) else 0,
            "最小值": float(s.min()) if len(s) else np.nan,
            "最大值": float(s.max()) if len(s) else np.nan,
            "是否可绘图": "是" if len(s) > 0 else "否"
        })
    return pd.DataFrame(rows)


def data_understanding_summary(score, main_metric, metric_candidates, dimension_candidates, date_candidates, clean_summary):
    level, desc = health_level(score)
    metric_text = "、".join(metric_candidates[:4]) if metric_candidates else "暂无明确主指标"
    dim_text = "、".join(dimension_candidates[:4]) if dimension_candidates else "暂无明确维度字段"
    date_text = "、".join(date_candidates[:2]) if date_candidates else "未识别到有效日期字段"
    rows = clean_summary.get("清洗后行数", 0)
    cols = clean_summary.get("清洗后字段数", 0)
    st.markdown(f"""
    <div class="health-summary">
        <h3>🤖 AI数据理解摘要</h3>
        <div class="health-line">系统识别到当前数据共有 <b>{rows:,}</b> 条记录、<b>{cols}</b> 个字段，数据健康状态为 <b>{level}</b>：{desc}</div>
        <div class="health-line">推荐优先以 <b>{main_metric}</b> 作为当前主分析指标；可选主指标包括：{metric_text}。</div>
        <div class="health-line">推荐用于分组分析的维度包括：{dim_text}。</div>
        <div class="health-line">时间字段识别结果：{date_text}。如果没有有效日期字段，系统将自动转为结构分析模式。</div>
    </div>
    """, unsafe_allow_html=True)


tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🧩 数据治理",
    "🏠 经营态势",
    "📊 经营洞察",
    "⚠️ 风险识别",
    "💬 问数助手",
    "📝 决策简报"
])


# ============================================================
# Tab 1
# ============================================================


with tab1:
    st.subheader("🧩 数据治理")
    st.markdown('<div class="tip-card">系统自动评估数据质量、识别字段角色并生成字段资产地图，帮助用户快速确认哪些字段适合作为主指标、分析维度和时间字段；详细字段表与处理日志默认折叠，便于按需追溯分析依据。</div>', unsafe_allow_html=True)

    health_score = compute_data_health_score(clean_summary, meta, numeric_report, metric_candidates, dimension_candidates, date_candidates)

    top_left, top_right = st.columns([1.05, 1.55])
    with top_left:
        level, desc = gauge_chart(health_score)
    with top_right:
        data_understanding_summary(health_score, main_metric, metric_candidates, dimension_candidates, date_candidates, clean_summary)

    st.markdown("### 数据资产概览")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("数据记录数", f"{clean_summary['清洗后行数']:,}", "清洗后的可分析记录")
    with c2:
        kpi_card("字段数量", f"{clean_summary['清洗后字段数']}", "当前数据字段数")
    with c3:
        kpi_card("可选主指标", f"{len(metric_candidates)}", "系统推荐的数值指标")
    with c4:
        kpi_card("可选维度", f"{len(dimension_candidates)}", "可用于分组分析的字段")

    st.markdown("### 字段资产地图")
    c1, c2 = st.columns([1.05, 1.45])
    with c1:
        role_distribution_chart(meta)
    with c2:
        field_recommendation_cards(meta, metric_candidates, dimension_candidates, date_candidates)

    st.markdown("### 数据质量体检")
    c1, c2 = st.columns([1.05, 1.45])
    with c1:
        data_quality_issue_chart(clean_summary, quality_report, numeric_report)
    with c2:
        numeric_validity_progress(numeric_report)

    st.markdown("### 当前分析口径")
    config_df = pd.DataFrame({
        "配置项": ["主分析指标", "有效日期字段", "分析维度字段", "数值字段", "缺失值策略"],
        "当前选择": [
            main_metric,
            valid_date_col if valid_date_col else "无",
            "、".join(selected_dimensions) if selected_dimensions else "无",
            "、".join(numeric_cols) if numeric_cols else "无",
            missing_strategy
        ]
    })
    st.dataframe(config_df, use_container_width=True)

    st.markdown("### 明细查看")
    with st.expander("查看字段语义识别表", expanded=False):
        st.markdown('<div class="detail-expander-note">字段解释、推荐角色、聚合方式和注意事项都会在这里展示，适合需要核对字段口径时查看。</div>', unsafe_allow_html=True)
        st.dataframe(meta, use_container_width=True, height=420)

    with st.expander("查看数据质量处理日志", expanded=False):
        st.dataframe(quality_report, use_container_width=True)

    with st.expander("查看数值转换检查", expanded=False):
        st.dataframe(numeric_report, use_container_width=True)

    with st.expander("查看图表可绘制性诊断", expanded=False):
        st.markdown('<div class="detail-expander-note">如果某张图画不出来，优先看这里：有效数值数为 0 或唯一值过少时，系统会提示无法形成有效图形，而不是显示空白图。</div>', unsafe_allow_html=True)
        st.dataframe(chart_data_diagnostic(df, main_metric, numeric_cols, selected_dimensions), use_container_width=True)

    with st.expander("查看数据预览", expanded=False):
        st.dataframe(df.head(30), use_container_width=True)


# ============================================================
# Tab 2
# ============================================================

with tab2:
    st.subheader("🏠 经营态势")

    abnormal_count = int(anomaly_df["是否经营异常"].sum()) if len(anomaly_df) else 0
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card(f"{main_metric}合计", money_fmt(df[main_metric].sum()), "基于当前上传数据")
    with c2:
        kpi_card(f"{main_metric}均值", money_fmt(df[main_metric].mean()), "反映平均水平")
    with c3:
        kpi_card(f"{main_metric}最大值", money_fmt(df[main_metric].max()), "用于识别高值记录")
    with c4:
        kpi_card("异常经营单元", f"{abnormal_count}", "中高风险单元数量")

    if valid_date_col:
        st.markdown("### 核心趋势")
        trend_all = build_trend(df, valid_date_col, main_metric)
        periods_all = trend_all["期间"].astype(str).tolist()
        if len(periods_all) >= 2:
            fc1, fc2, fc3 = st.columns([1.2, 1.2, 2])
            with fc1:
                start_p = st.selectbox("趋势起始时间", periods_all, index=0, key="dash_start_period")
            with fc2:
                end_p = st.selectbox("趋势结束时间", periods_all, index=len(periods_all)-1, key="dash_end_period")
            sidx, eidx = periods_all.index(start_p), periods_all.index(end_p)
            if sidx > eidx:
                sidx, eidx = eidx, sidx
            selected_periods = periods_all[sidx:eidx+1]
            trend = trend_all[trend_all["期间"].isin(selected_periods)].copy()
        else:
            selected_periods = periods_all
            trend = trend_all.copy()
        c1, c2 = st.columns([2, 1])
        with c1:
            line_chart(trend, "期间", main_metric, f"{main_metric}时间趋势")
        with c2:
            st.markdown(f'<div class="section-card"><h4>趋势解读</h4><p class="small-text">{trend_interpretation(trend, main_metric)}</p></div>', unsafe_allow_html=True)

        st.markdown("### 动态趋势演示")
        st.markdown("""
        <div class="highlight-note">
        <b>显示条件：</b>只有当上传数据中存在有效日期字段、主指标可按时间汇总，且时间周期不少于 3 个时，系统才会生成动态图。
        动态图按时间顺序逐步展开主指标轨迹，用于观察经营指标从起点到当前周期的演变过程；横轴只保留关键区间刻度，避免时间标签过密。
        </div>
        """, unsafe_allow_html=True)
        animated_trend_chart(trend, main_metric, f"{main_metric}动态趋势演示")

        if selected_dimensions:
            st.markdown("### 分维度趋势对比")
            dc1, dc2, dc3 = st.columns([1.4, 1, 1.2])
            with dc1:
                dim = st.selectbox("选择趋势拆分维度", selected_dimensions, index=0, key="dash_dim_trend")
            with dc2:
                topn_line = st.slider("展示TopN维度项", 3, 12, 6, key="dash_dim_topn")
            temp = periodize(df, valid_date_col)
            if selected_periods:
                temp = temp[temp["_period"].isin(selected_periods)]
            top_dims = df.groupby(dim)[main_metric].sum().sort_values(ascending=False).head(topn_line).index.tolist()
            with dc3:
                highlight = st.selectbox("高亮维度项", ["无"] + [str(x) for x in top_dims], key="dash_highlight_dim")
            dim_trend = temp[temp[dim].isin(top_dims)].groupby(["_period", dim])[main_metric].sum().reset_index()
            fig = go.Figure()
            for name, g in dim_trend.groupby(dim):
                is_high = (highlight != "无" and str(name) == str(highlight))
                fig.add_trace(go.Scatter(x=g["_period"].astype(str), y=g[main_metric], mode="lines+markers", name=str(name), line=dict(width=4 if is_high else 1.6), opacity=1.0 if is_high or highlight == "无" else 0.22, marker=dict(size=7 if is_high else 4)))
            fig.update_layout(title=f"按{dim}对比{main_metric}趋势", xaxis_title="期间", yaxis_title=main_metric, hovermode="closest")
            st.plotly_chart(chart_layout(fig, 460), use_container_width=True)
            st.caption("维度项较多时，系统默认展示主指标贡献较高的TopN。可通过“高亮维度项”突出某一条线，其余线降低透明度，避免画面过乱。")
    else:
        st.markdown('<div class="warn-card">当前未选择有效日期字段，系统已切换为“结构分析模式”：重点展示主指标在不同维度下的分布、排名和异常情况。</div>', unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            histogram_chart(df, main_metric, f"{main_metric}分布")
        with c2:
            box_chart(df, main_metric, f"{main_metric}箱线图")

        if selected_dimensions:
            dim = selected_dimensions[0]
            g = dimension_summary(df, dim, main_metric).head(15)
            c1, c2 = st.columns(2)
            with c1:
                bar_chart(g, dim, f"{main_metric}合计", f"按{dim}对比{main_metric}合计")
            with c2:
                g2 = g.sort_values(f"{main_metric}均值", ascending=False)
                bar_chart(g2, dim, f"{main_metric}均值", f"按{dim}对比{main_metric}均值")

    st.markdown("### 重点风险提示")
    st.markdown("""
    <div class="highlight-note">
    <b>说明：</b>经营态势页只承担“总览预警”作用，帮助管理者快速看到是否存在需要优先关注的高风险对象；
    详细异常原因、风险得分构成和AI解释请进入“风险识别”页面下钻查看。
    </div>
    """, unsafe_allow_html=True)
    if len(anomaly_df):
        top = anomaly_df[anomaly_df["是否经营异常"]].sort_values("风险得分", ascending=False).head(3)
        if len(top):
            ignore_cols = set(numeric_cols + ["记录数", "是否AI异常", "异常得分", "异常依据", "主指标偏离", "多指标组合偏离", "业务规则风险", "模型异常贡献", "风险得分", "风险等级", "是否经营异常"])
            title_cols = [c for c in anomaly_df.columns if c not in ignore_cols]
            rows = []
            for _, row in top.iterrows():
                obj = "｜".join([f"{c}={row[c]}" for c in title_cols[:2]]) if title_cols else f"记录 {row.name}"
                rows.append({
                    "风险对象": obj,
                    "风险等级": row["风险等级"],
                    "风险得分": row["风险得分"],
                    "简要原因": str(row["异常依据"])[:95] + ("..." if len(str(row["异常依据"])) > 95 else "")
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.success("当前未识别出明显中高风险经营单元。")


# ============================================================
# Tab 3
# ============================================================

with tab3:
    st.subheader("📊 经营洞察")
    sub1, sub2, sub3, sub4 = st.tabs(["维度对比", "贡献度分析", "指标关系", "热力图"])

    with sub1:
        if not selected_dimensions:
            st.warning("当前未选择维度字段，无法进行维度对比。")
        else:
            dim = st.selectbox("选择分析维度", selected_dimensions, key="dim_compare")
            g = dimension_summary(df, dim, main_metric)
            c1, c2 = st.columns(2)
            with c1:
                bar_chart(g.head(20), dim, f"{main_metric}合计", f"{dim}维度下{main_metric}合计排名")
            with c2:
                pie_df = g.head(10).copy()
                pie_val_col = f"{main_metric}合计"
                pie_df[pie_val_col] = pd.to_numeric(pie_df[pie_val_col], errors="coerce").replace([np.inf, -np.inf], np.nan)
                pie_df = pie_df.dropna(subset=[pie_val_col])
                if len(pie_df) and pie_df[pie_val_col].abs().sum() > 0:
                    fig = go.Figure(go.Pie(
                        labels=pie_df[dim].astype(str).tolist(),
                        values=pie_df[pie_val_col].astype(float).tolist(),
                        hole=0.35
                    ))
                    fig.update_layout(title=f"{dim}维度贡献占比 Top10")
                    safe_plotly_chart(fig, f"{dim}维度贡献占比 Top10", 460)
                else:
                    show_no_chart_reason(f"{dim}维度贡献占比 Top10", "该维度下没有可用于饼图的正负有效数值。")

            st.dataframe(g, use_container_width=True)

            importance = dimension_importance(df, selected_dimensions, main_metric)
            if len(importance):
                st.markdown("#### 智能维度推荐")
                st.dataframe(importance, use_container_width=True)
                best_dim = importance.iloc[0]["维度"]
                st.info(f"系统建议优先关注“{best_dim}”维度，因为该维度下主指标的组间差异较明显。")

    with sub2:
        if valid_date_col and selected_dimensions:
            dim = st.selectbox("选择贡献度维度", selected_dimensions, key="contrib_dim")
            periods = sorted(periodize(df, valid_date_col)["_period"].dropna().unique().tolist())
            if len(periods) >= 2:
                pc1, pc2 = st.columns(2)
                with pc1:
                    p0 = st.selectbox("对比起始时间", periods, index=max(0, len(periods)-2), key="contrib_p0")
                with pc2:
                    p1 = st.selectbox("对比结束时间", periods, index=len(periods)-1, key="contrib_p1")
                comp = contribution_analysis(df, valid_date_col, dim, main_metric, p0, p1)
                if len(comp):
                    bar_chart(comp.head(15), dim, "变化额", f"{p0} 至 {p1} 的{main_metric}变化贡献")
                    st.dataframe(comp, use_container_width=True)
                    st.info("贡献度分析用于识别两个指定时间段之间，主指标变化主要由哪些维度项推动。正值代表拉动增长，负值代表拖累下降。")
                else:
                    st.info("请选择两个不同时间段进行贡献度分析。")
            else:
                st.info("当前时间数据不足以计算贡献度。")
        else:
            st.info("贡献度分析需要有效日期字段和至少一个维度字段。若无日期字段，可使用维度对比和结构分析替代。")

    with sub3:
        relation_candidates = [c for c in numeric_cols if not is_id_like(c)]
        if len(relation_candidates) < 2:
            st.warning("可用于关系分析的数值指标不足。系统会自动排除 ID/编号类字段。")
        else:
            c1, c2, c3 = st.columns([1, 1, 1.2])
            with c1:
                x_col = st.selectbox("X轴指标", relation_candidates, index=relation_candidates.index(main_metric) if main_metric in relation_candidates else 0)
            with c2:
                default_y_idx = 1 if len(relation_candidates) > 1 and relation_candidates[0] == x_col else 0
                y_col = st.selectbox("Y轴指标", relation_candidates, index=default_y_idx)
            with c3:
                group_col = st.selectbox("分组字段（可选）", ["无"] + selected_dimensions, index=1 if selected_dimensions else 0, key="relation_group")
                group_col = None if group_col == "无" else group_col
            rec_method, rec_reason = recommend_fit_method(df, x_col, y_col, group_col)
            methods = ["全局OLS（直线）", "按维度分组OLS（分段直线）", "LOWESS（局部加权）", "多项式回归（二阶）"]
            st.markdown(f'<div class="tip-card"><b>智能推荐：</b>{rec_method}。{rec_reason}</div>', unsafe_allow_html=True)
            fit_method = st.radio("选择拟合方法", methods, index=methods.index(rec_method), horizontal=True, key="fit_method_radio")
            corr = scatter_chart(df, x_col, y_col, group_col, fit_method)
            st.markdown(f'<div class="tip-card">{relationship_explanation(x_col, y_col, corr)}</div>', unsafe_allow_html=True)
            st.caption("全局OLS用于观察整体线性方向；分组OLS用于比较不同群体斜率；LOWESS用于发现局部拐点和非线性；二阶多项式用于观察先升后降或先降后升关系。")

            st.markdown("### 全部经营指标相关系数汇总")
            corr_matrix, corr_pairs = correlation_summary_tables(df, relation_candidates)
            if len(corr_pairs):
                st.markdown('<div class="highlight-note">相关系数表用于汇总所有经营数值指标之间的线性关系。数值越接近 1 表示正向关系越强，越接近 -1 表示反向关系越强，越接近 0 表示线性关系较弱。该结果用于发现线索，不直接等同于因果关系。</div>', unsafe_allow_html=True)
                st.dataframe(corr_pairs, use_container_width=True, hide_index=True, height=360)
                with st.expander("查看相关系数矩阵表"):
                    st.dataframe(corr_matrix, use_container_width=True)
            else:
                st.info("当前可用于相关系数汇总的有效数值指标不足。")

    with sub4:
        heat_type = st.radio("选择热力图类型", ["数值指标相关性矩阵", "维度交叉分布热力图"], horizontal=True)
        if heat_type == "数值指标相关性矩阵":
            corr_heatmap_chart(df, numeric_cols)
        else:
            if len(selected_dimensions) >= 2:
                row_dim = st.selectbox("行维度", selected_dimensions, key="heat_row")
                col_dim = st.selectbox("列维度", [d for d in selected_dimensions if d != row_dim], key="heat_col")
                agg_name = st.selectbox("聚合方式", ["求和", "均值", "计数"], index=0)
                topn = st.slider("TopN", 5, 20, 10)
                if agg_name == "计数":
                    row_top = df.groupby(row_dim).size().sort_values(ascending=False).head(topn).index
                    col_top = df.groupby(col_dim).size().sort_values(ascending=False).head(topn).index
                else:
                    row_top = df.groupby(row_dim)[main_metric].sum().sort_values(ascending=False).head(topn).index
                    col_top = df.groupby(col_dim)[main_metric].sum().sort_values(ascending=False).head(topn).index
                temp = df[df[row_dim].isin(row_top) & df[col_dim].isin(col_top)]
                if agg_name == "求和":
                    pivot = temp.pivot_table(index=row_dim, columns=col_dim, values=main_metric, aggfunc="sum", fill_value=0)
                elif agg_name == "均值":
                    pivot = temp.pivot_table(index=row_dim, columns=col_dim, values=main_metric, aggfunc="mean", fill_value=0)
                else:
                    pivot = temp.pivot_table(index=row_dim, columns=col_dim, values=main_metric, aggfunc="count", fill_value=0)
                heatmap_chart(pivot, f"{row_dim} × {col_dim} 的 {main_metric} 交叉分布热力图（{agg_name}）")
                st.info("交叉分布热力图用于观察两个维度组合下主指标的分布差异；数值指标相关性矩阵用于观察数值字段之间的相关关系。")
            else:
                st.info("维度交叉分布热力图需要至少两个维度字段。")


# ============================================================
# Tab 4
# ============================================================

with tab4:
    st.subheader("⚠️ 风险识别")
    st.markdown("""
    <div class="tip-card">
    <b>异常诊断说明：</b><br>
    本模块中的“异常”并不等同于数据错误，也不预设某一种固定经营场景。系统会先基于当前上传数据中实际存在的字段构建经营单元，再从四个方面综合判断风险：<br>
    ① <b>主指标分布偏离</b>：主指标处于当前样本的高位或低位区间；<br>
    ② <b>多指标组合偏离</b>：多个数值字段同时处于极端区间，说明该单元与整体样本差异较大；<br>
    ③ <b>动态业务规则</b>：仅当数据中真实存在利润、成本、费用、比例、预算等字段时，才生成对应的业务异常标签；<br>
    ④ <b>模型异常贡献</b>：Isolation Forest 根据多指标组合识别出的离群程度。<br>
    风险得分越高，表示该经营单元越需要优先核查；系统不会根据不存在的字段生成异常结论。
    </div>
    """, unsafe_allow_html=True)

    if len(anomaly_df) == 0:
        st.warning("当前数据不足以进行异常诊断。")
    else:
        abnormal = anomaly_df[anomaly_df["是否经营异常"]].copy()
        high = abnormal[abnormal["风险等级"] == "高风险"]

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            kpi_card("异常经营单元", f"{len(abnormal)}", "中高风险单元数量")
        with c2:
            kpi_card("高风险单元", f"{len(high)}", "风险得分≥70")
        with c3:
            kpi_card("最高风险得分", f"{anomaly_df['风险得分'].max():.1f}", "越高表示偏离越明显")
        with c4:
            kpi_card("异常占比", pct_fmt(len(abnormal) / max(len(anomaly_df), 1)), "异常单元 / 全部单元")

        c1, c2 = st.columns(2)
        with c1:
            fig = go.Figure(go.Histogram(x=anomaly_df["风险得分"], nbinsx=20))
            fig.update_layout(title="风险得分分布", xaxis_title="风险得分", yaxis_title="数量")
            st.plotly_chart(chart_layout(fig, 360), use_container_width=True)
        with c2:
            rc = anomaly_df["风险等级"].value_counts().reset_index()
            rc.columns = ["风险等级", "数量"]
            bar_chart(rc, "风险等级", "数量", "风险等级分布")

        st.markdown("### Top 异常经营单元")
        top_anom = anomaly_df.sort_values("风险得分", ascending=False).head(10)
        st.dataframe(top_anom, use_container_width=True, height=320)

        st.markdown("### 异常解释与风险组成")
        if len(top_anom):
            chosen_idx = st.selectbox("选择一个异常单元查看解释", top_anom.index.tolist())
            row = anomaly_df.loc[chosen_idx]

            ignore_cols = set(numeric_cols + ["记录数", "是否AI异常", "异常得分", "异常依据", "主指标偏离", "多指标组合偏离", "业务规则风险", "模型异常贡献", "风险得分", "风险等级", "是否经营异常"])
            title_cols = [c for c in anomaly_df.columns if c not in ignore_cols]
            obj = "｜".join([f"{c}={row[c]}" for c in title_cols[:3]]) if title_cols else f"记录 {chosen_idx}"

            c1, c2 = st.columns([1, 1])
            with c1:
                st.markdown(f"""
                <div class="section-card">
                <h4>为什么是异常？</h4>
                <p class="small-text"><b>对象：</b>{obj}</p>
                <p class="small-text"><b>风险等级：</b>{row['风险等级']}｜<b>风险得分：</b>{row['风险得分']}</p>
                <p class="small-text"><b>异常依据：</b>{row['异常依据']}</p>
                </div>
                """, unsafe_allow_html=True)
            with c2:
                comp = pd.DataFrame({
                    "风险来源": ["主指标偏离", "多指标组合偏离", "业务规则风险", "模型异常贡献"],
                    "得分": [row["主指标偏离"], row["多指标组合偏离"], row["业务规则风险"], row["模型异常贡献"]]
                })
                radar_chart(comp)

            st.markdown("#### 风险得分构成")
            st.dataframe(comp, use_container_width=True)

            with st.expander("🤖 使用大模型生成个性化异常原因与行动建议"):
                provider = st.selectbox("选择大模型", list(get_llm_configs().keys()), key="anom_llm")
                if st.button("生成AI异常解释", key="anom_ai_btn"):
                    try:
                        sys_p = "你是经营风险诊断专家。请根据异常单元数据，解释为什么异常、为什么高风险，并给出可执行建议。"
                        user_p = f"异常对象：{obj}\n风险等级：{row['风险等级']}\n风险得分：{row['风险得分']}\n异常依据：{row['异常依据']}\n风险得分构成：{comp.to_dict(orient='records')}\n该单元完整数据：{row.to_dict()}\n\n请输出：1. 异常原因；2. 高风险判断依据；3. 可能影响；4. 管理建议；5. 下一步核查数据。"
                        ans, elapsed = call_llm(provider, sys_p, user_p, temperature=0.2)
                        st.success(f"生成完成，耗时 {elapsed:.2f} 秒")
                        render_ai_answer_pretty(ans)
                    except Exception as e:
                        st.error(f"AI异常解释生成失败：{e}")

        st.markdown("### 异常依据补充")
        st.info("上表中的“异常依据”会说明具体触发原因，例如主指标处于高位区间、多项指标同时极端、或实际存在字段触发业务规则。建议结合原始明细进一步核查。")


# ============================================================
# Tab 5
# ============================================================

with tab5:
    st.subheader("💬 问数助手")
    st.markdown("用户可输入任意与当前数据相关的问题，系统会调用所选大模型生成 SQL、执行查询，并返回结果解释。")

    with st.expander("查看问数任务综合评分规则"):
        st.markdown("""
        **问数任务综合评分**用于比较不同大模型在本轮数据问数任务中的表现，不等同于严格人工标注准确率。评分规则如下：

        1. **SQL可执行性（40分）**：生成的SQL必须是安全的 `SELECT` 查询，并且能在当前上传数据表上成功运行；若字段名错误、语法错误或生成非查询语句，则该项为0分。
        2. **查询结果有效性（20分）**：SQL成功执行且返回非空结果，通常给20分；如果成功执行但结果为空，说明查询条件或理解可能有偏差，通常只给8分。
        3. **字段匹配度（15分）**：系统会检查SQL是否使用了当前数据中的真实字段。使用了相关字段通常给15分；虽然可执行但字段使用较弱时会降分。
        4. **结果解释质量（15分）**：模型需要基于查询结果给出清晰的经营解释，包括结论、含义和建议；解释过短或缺失会扣分。
        5. **响应速度（10分）**：响应时间≤3秒为10分；3—8秒为8分；8—15秒为5分；超过15秒为2分；调用失败为0分。

        因此，分数越高表示本轮问数结果越“可用”，但它不是严格人工标注准确率。
        """)

    st.markdown("### 💡 你可以这样问")
    examples = generate_dynamic_questions(main_metric, selected_dimensions, numeric_cols, valid_date_col)
    rows = [examples[i:i+3] for i in range(0, len(examples), 3)]
    for row_qs in rows[:2]:
        cols = st.columns(3)
        for col_obj, (typ, q) in zip(cols, row_qs):
            with col_obj:
                st.markdown(f"""
                <div class="example-card">
                    <div class="example-title">{typ}</div>
                    <div class="example-q">{q}</div>
                </div>
                """, unsafe_allow_html=True)

    providers = list(get_llm_configs().keys())
    selected_providers = st.multiselect("选择参与问数/对比的大模型", providers, default=providers)
    question = st.text_input("请输入问题", value=examples[0][1] if examples else f"请分析{main_metric}的整体情况")

    if "llm_records" not in st.session_state:
        st.session_state["llm_records"] = []

    if st.button("开始问数", type="primary"):
        if not selected_providers:
            st.warning("请至少选择一个大模型。")
        else:
            for provider in selected_providers:
                st.markdown(f"## {provider}")
                with st.spinner(f"{provider} 正在生成 SQL 并查询..."):
                    try:
                        pack = run_llm_sql_question(provider, question, df, meta)
                    except Exception as e:
                        pack = {"模型": provider, "SQL": "", "是否成功": False, "错误信息": str(e), "结果": pd.DataFrame(), "解释": "", "响应时间": np.nan, "综合评分": 0, "评分明细": {}}

                if pack["SQL"]:
                    st.code(pack["SQL"], language="sql")
                if pack["是否成功"]:
                    st.success(f"执行成功｜响应时间：{pack['响应时间']:.2f}秒｜综合评分：{pack['综合评分']}")
                    st.dataframe(pack["结果"], use_container_width=True)
                    st.markdown("#### 模型解释")
                    st.info(pack["解释"])
                else:
                    st.error(f"执行失败：{pack['错误信息']}")

                with st.expander("查看评分明细"):
                    st.json(pack["评分明细"])

                st.session_state["llm_records"].append({
                    "模型": provider,
                    "问题": question,
                    "SQL可执行": "是" if pack["是否成功"] else "否",
                    "结果行数": len(pack["结果"]) if isinstance(pack["结果"], pd.DataFrame) else 0,
                    "响应时间秒": round(pack["响应时间"], 2) if not pd.isna(pack["响应时间"]) else None,
                    "问数任务综合评分": pack["综合评分"],
                    "错误信息": pack["错误信息"]
                })

    st.markdown("### 模型本轮/历史对比")
    records = st.session_state.get("llm_records", [])
    if records:
        compare_df = pd.DataFrame(records)
        st.dataframe(compare_df.tail(20), use_container_width=True)
        summary = compare_df.groupby("模型").agg(
            测试次数=("模型", "count"),
            成功次数=("SQL可执行", lambda x: (x == "是").sum()),
            平均响应时间=("响应时间秒", "mean"),
            平均评分=("问数任务综合评分", "mean")
        ).reset_index()
        bar_chart(summary, "模型", "平均评分", "平均问数任务综合评分")
    else:
        st.info("暂无问数记录。")


# ============================================================
# Tab 6
# ============================================================

with tab6:
    st.subheader("📝 决策简报生成")

    focus = st.multiselect(
        "选择本次简报关注重点",
        ["综合经营", "趋势变化", "维度结构", "异常风险", "成本/费用", "利润/收益", "预算执行", "人力/薪酬", "采购/供应商"],
        default=["综合经营", "维度结构", "异常风险"]
    )

    st.markdown("### 决策简报预览")
    st.info(f"系统将基于当前上传数据、主分析指标“{main_metric}”和用户选择的关注重点生成 Word 决策简报。下方展示的是即将写入报告的核心分析口径和代表性图表，完整报告将在点击生成后下载。当前数据共 {len(df):,} 条记录。")

    if selected_dimensions:
        dim = selected_dimensions[0]
        full_g = dimension_summary(df, dim, main_metric)
        g = full_g.head(10)
        chart_title = f"代表性图表预览：按{dim}汇总{main_metric}" if len(full_g) <= 10 else f"代表性图表预览：按{dim}汇总{main_metric} Top10"
        bar_chart(g, dim, f"{main_metric}合计", chart_title)

    if st.button("生成并下载 Word 简报", type="primary"):
        try:
            bio = generate_report_docx(df, main_metric, selected_dimensions, valid_date_col, anomaly_df, focus)
            st.download_button(
                "下载 Word 简报",
                data=bio,
                file_name="智策经营_管理层决策简报.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Word 简报生成失败：{e}")

    with st.expander("使用大模型生成更丰富的简报正文"):
        provider = st.selectbox("选择大模型", list(get_llm_configs().keys()), key="report_llm")
        if st.button("生成AI简报正文", key="report_ai_btn"):
            try:
                abnormal_count = int(anomaly_df["是否经营异常"].sum()) if len(anomaly_df) else 0
                top_anom = anomaly_df.sort_values("风险得分", ascending=False).head(5).to_dict(orient="records") if len(anomaly_df) else []
                sys_p = "你是企业经营分析顾问。请生成正式、结构清晰、管理层可读的决策简报正文。不得引用数据中不存在的字段。"
                user_p = f"主指标：{main_metric}\n关注重点：{'、'.join(focus)}\n数据记录数：{len(df)}\n主指标合计：{df[main_metric].sum()}\n主指标均值：{df[main_metric].mean()}\n维度字段：{selected_dimensions}\n异常经营单元数：{abnormal_count}\nTop异常：{top_anom}\n\n请按以下结构输出：一、分析概况；二、核心经营表现；三、风险识别；四、原因分析与管理建议；五、后续关注重点。"
                ans, elapsed = call_llm(provider, sys_p, user_p, temperature=0.25)
                st.success(f"生成完成，耗时 {elapsed:.2f} 秒")
                render_ai_answer_pretty(ans)
                try:
                    ai_bio = generate_ai_report_docx(ans, df, main_metric, selected_dimensions, valid_date_col, anomaly_df, focus)
                    st.download_button(
                        "下载 AI增强版 Word 简报",
                        data=ai_bio,
                        file_name="智策经营_AI增强版决策简报.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as doc_e:
                    st.warning(f"AI简报正文已生成，但Word下载文件生成失败：{doc_e}")
            except Exception as e:
                st.error(f"AI简报正文生成失败：{e}")
