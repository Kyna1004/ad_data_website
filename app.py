import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import xlsxwriter
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import docx.opc.constants

# ==========================================
# PART 1: 配置区域 (修复了字段映射)
# ==========================================

COMMON_METRICS = {
    "spend": ["花费金额(USD)", "花费金额 （USD）", "花费金额 (USD)", "花费金额", "Amount Spent"],
    "roas": ["广告花费回报 (ROAS) - 购物", "广告花费回报（ROAS）-购物", "ROAS", "Purchase ROAS"],
    "purchases": ["购买次数", "成效数量", "成效", "Purchases"],
    "cpa": ["单次购买费用", "单次购物成本", "单次成效成本", "单次成效费用", "Cost per Purchase"],
    "ctr": ["链接点击率", "链接点击率（%)", "链接点击率（%）", "CTR"],
    "cpm": ["千次展示费用", "CPM"],
    "clicks": ["点击", "链接点击", "Clicks"],
    "impressions": ["曝光", "展示次数", "Impressions"],
    "purchase_value": ["购买价值", "购物价值", "Purchase Value"],
    "aov": ["单次购买价值", "单次购物价值"]
}

# 框定「每一个 Sheet」需要抽取哪些指标
# ✅ 修复：在"受众组"中增加了 converting_countries/genders/ages 映射
SHEET_MAPPINGS = {
    "整体数据": {
        **COMMON_METRICS,
        "date_range": ["时间范围"],
        "clicks_all": ["点击"],
        "landing_page_views": ["落地页浏览量"],
        "add_to_cart": ["加入购物车"],
        "initiate_checkout": ["结账发起次数"],
        "rate_click_to_lp": ["点击-落地页浏览转化率"],
        "rate_lp_to_atc": ["落地页浏览-加购转化率"],
        "rate_atc_to_ic": ["加购-结账转化率"],
        "rate_ic_to_pur": ["结账-购买转化率"]
    },
    "分时段数据": {
        **COMMON_METRICS,
        "date_range": ["时间范围"],
        "landing_page_views": ["落地页浏览量"],
        "add_to_cart": ["加入购物车"],
        "initiate_checkout": ["结账发起次数"],
        "rate_click_to_lp": ["点击-落地页浏览转化率"],
        "rate_lp_to_atc": ["落地页浏览-加购转化率"],
        "rate_atc_to_ic": ["加购-结账转化率"],
        "rate_ic_to_pur": ["结账-购买转化率"]
    },
    "异常指标": {
        "anomaly_metric_name": ["异常指标"],
        "mom_change": ["环比"]
    },
    "广告架构": {**COMMON_METRICS, "dimension_item": ["广告类型"]},
    "受众组": {
        **COMMON_METRICS,
        "dimension_item": ["广告组", "广告组Id", "Ad Set Name"],
        "custom_audience_settings": ["设置的自定义受众", "Custom Audiences"],
        "converting_keywords": ["产生成效的关键词", "Interests", "Keywords"],
        # ✅ 新增以下三行，确保从Excel中读取这些列
        "converting_countries": ["产生成效的国家", "国家", "地区", "Country", "Region", "Location"],
        "converting_genders": ["产生成效的性别", "性别", "Gender"],
        "converting_ages": ["产生成效的年龄", "年龄", "Age", "Age Group"]
    },
    "受众类型": {**COMMON_METRICS, "dimension_item": ["受众类型"]},
    "国家": {**COMMON_METRICS, "dimension_item": ["国家/地区", "国家"]},
    "年龄": {**COMMON_METRICS, "dimension_item": ["年龄"]},
    "性别": {**COMMON_METRICS, "dimension_item": ["性别"]},
    "平台&版位": {**COMMON_METRICS, "dimension_item": ["平台&版位"]},
    "素材": {
        **COMMON_METRICS,
        "content_item": ["素材"],
        "cvr_lp_to_pur": ["落地页浏览-购买转化率"]
    },
    "落地页": {
        **COMMON_METRICS,
        "content_item": ["落地页url", "落地页"],
        "ctr_all": ["曝光-点击转化率"],
        "rate_lp_to_atc": ["落地页浏览-加购转化率", "落地页浏览-购物转化率"]
    }
}

GROUP_CONFIG = {
    "Master_Overview": ["整体数据", "分时段数据", "异常指标"],
    "Master_Breakdown": ["广告架构", "受众组", "受众类型", "国家", "年龄", "性别", "平台&版位"],
    "Master_Creative": ["素材", "落地页"]
}

REPORT_MAPPING = {
    "spend": "花费 ($)", "roas": "ROAS", "purchases": "购买次数", "purchase_value": "购买总价值",
    "cpa": "CPA ($)", "ctr": "CTR (%)", "cpm": "CPM ($)", "aov": "客单价",
    "impressions": "展现量", "clicks_all": "点击量 (All)", "clicks": "点击量 (All)", "ctr_all": "点击率 (All)",
    "landing_page_views": "落地页访问量", "add_to_cart": "加购次数", "initiate_checkout": "结账发起数 (IC)",
    "rate_click_to_lp": "点击 → 落地页访问转化率", "rate_lp_to_atc": "落地页 → 加购转化率",
    "rate_atc_to_ic": "加购 → 购买转化率", "rate_ic_to_pur": "购买转化率",
    "cvr_purchase": "点击 → 购买转化率", "cvr_lp_to_pur": "CVR (全站转化率)",
    "date_range": "日期/时段", "campaign_type": "投放模式", "adset_name": "广告组ID", "adset_id": "广告组ID",
    "custom_audience_settings": "自定义受众源", "converting_keywords": "高潜兴趣词", "audience_type": "受众策略",
    "country": "国家", "age_group": "年龄", "gender": "性别", "creative_name": "素材名称", "placement": "版位",
    "landing_page_url": "页面 URL", "mom_change": "环比波动", "anomaly_metric_name": "异常项",
    "converting_countries": "产生成效的国家", "converting_genders": "产生成效的性别", "converting_ages": "产生成效的年龄"
}

# ✅ 增强了模糊匹配别名
FIELD_ALIASES = {
    "adset_id": ["adset_id", "ad set id", "adset id", "广告组编号", "广告组id", "adset_name", "ad set name"],
    "converting_countries": ["converting_countries", "country", "region", "国家", "地区", "location"],
    "converting_genders": ["converting_genders", "gender", "性别"],
    "converting_ages": ["converting_ages", "age", "年龄", "age_group"],
    "converting_keywords": ["converting_keywords", "keywords", "interests", "兴趣", "关键词"],
    "spend": ["spend", "amount spent", "cost", "花费", "消耗"],
    "purchases": ["purchases", "results", "result", "成效", "购买"],
    "roas": ["roas", "return on ad spend", "purchase roas"],
    "purchase_value": ["purchase_value", "conversion value", "value", "总价值", "gmv", "购买总价值"],
    "clicks": ["clicks", "clicks (all)", "点击量", "clicks_all"],
    "impressions": ["impressions", "展示", "展现"],
    "ctr_all": ["ctr_all", "ctr (all)", "点击率 (all)"]
}


# ==========================================
# PART 2: 核心工具函数 (保持不变，略)
# ==========================================

# ==========================================
# PART 2: 核心工具函数 (已修复百分比识别问题)
# ==========================================

def parse_float(value):
    """辅助函数：清理数据并将字符串/数字安全转换为浮点数"""
    if value is None:
        return 0.0
    try:
        # 如果已经是数字，直接返回
        if isinstance(value, (int, float)):
            return float(value)
        # 如果是字符串，调用 clean_numeric_strict 进行标准处理
        return clean_numeric_strict(value)
    except (ValueError, TypeError):
        return 0.0

def safe_div(numerator, denominator, multiplier=1.0):
    n = parse_float(numerator)
    d = parse_float(denominator)
    if d > 0:
        return (n / d) * multiplier
    else:
        return 0.0

# 宽松清洗（用于展示）
def clean_numeric(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('$', '').replace('¥', '').replace(',', '')
    
    # ✅ 修复点 1：如果是百分数字符串，转换后除以 100
    if '%' in val_str: 
        val_str = val_str.replace('%', '')
        try: return float(val_str) / 100.0 
        except: return 0.0
        
    try: return float(val_str)
    except: return val

# 严格清洗（用于计算）
def clean_numeric_strict(val): 
    if pd.isna(val): return 0.0
    # 如果已经是数字，直接返回
    if isinstance(val, (int, float)): return float(val)
    
    val_str = str(val).strip().replace('$', '').replace('¥', '').replace(',', '')
    
    # ✅ 修复点 2：如果是百分数字符串（如 "2.31%"），去除%后除以100还原为小数（0.0231）
    if '%' in val_str: 
        val_str = val_str.replace('%', '')
        try: return float(val_str) / 100.0
        except: return 0.0
        
    try: return float(val_str)
    except: return 0.0

# 字段鲁棒核心
def find_column_fuzzy(df, keywords):
    for kw in keywords:
        if kw in df.columns: return kw
    df_cols_norm = {c.lower().replace(' ', '').replace('_', ''): c for c in df.columns}
    for kw in keywords:
        kw_norm = kw.lower().replace(' ', '').replace('_', '')
        if kw_norm in df_cols_norm: return df_cols_norm[kw_norm]
    for col in df.columns:
        col_lower = col.lower()
        for kw in keywords:
            if kw.lower() in col_lower: return col
    return None

# 核心指标计算 (保持不变)
def calc_metrics_dict(df_chunk):
    res = {}
    if df_chunk.empty: return res
    sums = {}
    targets = ['spend', 'clicks', 'impressions', 'purchases', 'purchase_value',
               'landing_page_views', 'add_to_cart', 'initiate_checkout']
    
    for t in targets:
        aliases = FIELD_ALIASES.get(t, [t])
        if t == 'purchase_value' and 'value' not in aliases: aliases.append('value')
        col = find_column_fuzzy(df_chunk, aliases)
        if col:
             sums[t] = df_chunk[col].apply(clean_numeric_strict).sum()
        else:
             sums[t] = 0.0

    res['spend'] = parse_float(sums.get('spend', 0))
    res['impressions'] = parse_float(sums.get('impressions', 0))
    res['clicks'] = parse_float(sums.get('clicks', 0))
    res['purchases'] = parse_float(sums.get('purchases', 0))
    res['purchase_value'] = parse_float(sums.get('purchase_value', 0))
    res['roas'] = safe_div(sums.get('purchase_value'), sums.get('spend'))
    res['cpm'] = safe_div(sums.get('spend'), sums.get('impressions'), multiplier=1000)
    res['cpc'] = safe_div(sums.get('spend'), sums.get('clicks'))
    res['ctr'] = safe_div(sums.get('clicks'), sums.get('impressions'))
    res['cpa'] = safe_div(sums.get('spend'), sums.get('purchases'))
    res['cvr_purchase'] = safe_div(sums.get('purchases'), sums.get('clicks'))
    res['rate_click_to_lp'] = safe_div(sums.get('landing_page_views'), sums.get('clicks'))
    res['rate_lp_to_atc']   = safe_div(sums.get('add_to_cart'), sums.get('landing_page_views'))
    res['rate_atc_to_ic']   = safe_div(sums.get('initiate_checkout'), sums.get('add_to_cart'))
    res['rate_ic_to_pur']   = safe_div(sums.get('purchases'), sums.get('initiate_checkout'))
    res['aov'] = safe_div(sums.get('purchase_value'), sums.get('purchases'))

    date_col = find_column_fuzzy(df_chunk, ['date', 'time', 'range'])
    if date_col:
        try:
            dates = pd.to_datetime(df_chunk[date_col], errors='coerce').dropna()
            if not dates.empty: res['date_range'] = f"{dates.min():%Y-%m-%d} ~ {dates.max():%Y-%m-%d}"
            else: res['date_range'] = "-"
        except: res['date_range'] = "-"
    else: res['date_range'] = "-"
    return res 

def format_cell(key, val, is_mom=False):
    if isinstance(val, str): return val
    if is_mom:
        if key == 'date_range': return val
        return f"{val:+.2%}"
    k = str(key).lower()
    if 'roas' in k: return f"{val:.2f}"
    if any(x in k for x in ['rate', 'ctr', 'cvr', '点击率', '转化率', '着陆率', '意向率', '成功率']): 
        # 这里会乘以100，所以输入必须是小数 (0.0231 -> 2.31%)
        return f"{val:.2%}" 
    if any(x in k for x in ['spend', 'cpm', 'cpc', 'value', 'aov', 'cpa', '花费', '金额', '客单价', 'gmv', '价值']): return f"{val:,.2f}"
    if any(x in k for x in ['purchases', 'cart', 'click', '次数', '单量', '点击', '展现', '访问量', '发起数']): return f"{val:,.0f}"
    return f"{val}"

def extract_benchmark_values(df_bench):
    targets = {'roas': (['roas'], True), 'cpm': (['cpm'], False), 'ctr': (['ctr'], True), 'cpc': (['cpc'], False), 'cpa': (['cpa_purchase', 'cpa'], False)}
    extracted = {}
    for metric, (aliases, higher_better) in targets.items():
        found_col = None
        for alias in aliases:
            found_col = find_column_fuzzy(df_bench, [alias])
            if found_col: break
        if found_col:
            try:
                s = df_bench[found_col].apply(clean_numeric_strict)
                v = s[s>0].mean()
                
                # ✅ 修复点 3：防御性逻辑
                # 如果是 CTR/CVR 等比率类指标，且基准值 > 1.0 (例如用户填了 2.31 而不是 0.0231)，
                # 且该列不是 CPA/CPM/ROAS/CPC 这种本身就很大的值，则强制除以100
                if metric in ['ctr'] and v > 1.0:
                    v = v / 100.0
                    
                if not pd.isna(v): extracted[metric] = [v, higher_better]
            except: pass
    return extracted

# ... (add_hyperlink, apply_report_labels, add_df_to_word 保持不变)

def add_hyperlink(paragraph, url, text, color="0000FF", underline=True):
    try:
        part = paragraph.part
        r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        new_run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        if color:
            c = OxmlElement('w:color')
            c.set(qn('w:val'), color)
            rPr.append(c)
        if underline:
            u = OxmlElement('w:u')
            u.set(qn('w:val'), 'single')
            rPr.append(u)
        new_run.append(rPr)
        new_run.text = text
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)
        return hyperlink
    except: return None

def apply_report_labels(df, custom_mapping=None):
    if df.empty: return df
    mapping = REPORT_MAPPING.copy()
    if custom_mapping: mapping.update(custom_mapping)
    return df.rename(columns=mapping)

def add_df_to_word(doc, df, title, level=1):
    if df.empty: return
    doc.add_heading(title, level=level)
    t = doc.add_table(rows=df.shape[0]+1, cols=df.shape[1])
    t.style = 'Table Grid'
    is_creative = "素材" in title
    is_landing = "落地页" in title
    link_col_idx = -1
    for j, col in enumerate(df.columns):
        cell = t.cell(0, j)
        cell.text = str(col)
        if any(x in str(col).lower() for x in ["url", "link", "素材", "内容", "content"]): link_col_idx = j
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(8)
    for i in range(df.shape[0]):
        label_prefix = "素材" if is_creative else ("落地页" if is_landing else "")
        label_char = chr(65 + (i % 26))
        if i >= 26: label_char += str(i // 26)
        label_text = f"{label_prefix}{label_char}"
        for j in range(df.shape[1]):
            val = df.iat[i, j]
            cell = t.cell(i+1, j)
            if (is_creative or is_landing) and j == link_col_idx:
                try:
                    p = cell.paragraphs[0]
                    url = str(val).strip()
                    if len(url) > 5: add_hyperlink(p, url, label_text)
                    else: cell.text = label_text
                except: cell.text = label_text
            else:
                cell.text = str(val)
                if "结论" in str(df.columns[j]):
                    if "✅" in str(val): cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 128, 0)
                    if "⚠️" in str(val): cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 0, 0)
            for p in cell.paragraphs:
                for r in p.runs: r.font.size = Pt(8)
    doc.add_paragraph("\n")

# ==========================================
# PART 3: 主逻辑类 (Process ETL 中包含了关键修复)
# ==========================================

class AdReportProcessor:
    def __init__(self, raw_file, bench_file=None):
        self.raw_file = raw_file
        self.bench_file = bench_file
        self.processed_dfs = {}
        self.merged_dfs = {}
        self.final_json = {}
        self.doc = Document()

    # --- 阶段 1: 数据清洗与降维 ---
    def process_etl(self):
        xls = pd.ExcelFile(self.raw_file)
        for sheet_name, mapping in SHEET_MAPPINGS.items():
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                final_cols = {}
                # 字段映射
                for std_col, raw_col_options in mapping.items():
                    matched_col = None
                    for option in raw_col_options:
                        if option in df.columns: matched_col = option; break
                        if not matched_col:
                            for df_col in df.columns:
                                if option.replace(" ", "") == df_col.replace(" ", ""): matched_col = df_col; break
                        if matched_col: break
                    if matched_col: final_cols[std_col] = matched_col

                if final_cols:
                    df_clean = df[list(final_cols.values())].rename(columns={v: k for k, v in final_cols.items()})
                    
                    # ✅ 修复重点：将 converting_countries 等列加入「不进行数字清洗」的白名单
                    text_cols = ['date_range', 'anomaly_metric_name', 
                                 'converting_keywords', 'converting_countries', 'converting_genders', 'converting_ages', 
                                 'custom_audience_settings', 'dimension_item', 'content_item']
                    
                    for col in df_clean.columns:
                        if col not in text_cols:
                            df_clean[col] = df_clean[col].apply(clean_numeric)

                    if sheet_name in ["素材", "落地页", "受众组"]:
                        if "spend" in df_clean.columns:
                            df_clean = df_clean.sort_values("spend", ascending=False).head(10)

                    df_clean["Source_Sheet"] = sheet_name
                    self.processed_dfs[sheet_name] = df_clean

        # 合并 Master Tables
        for master_name, source_sheets in GROUP_CONFIG.items():
            dfs_to_merge = [self.processed_dfs[src] for src in source_sheets if src in self.processed_dfs]
            if dfs_to_merge:
                merged_df = pd.concat(dfs_to_merge, ignore_index=True)
                cols = list(merged_df.columns)
                priority_cols = ['Source_Sheet', 'date_range', 'dimension_item', 'content_item',
                                 'spend', 'roas', 'purchases', 'cpa']
                new_order = [c for c in priority_cols if c in cols] + [c for c in cols if c not in priority_cols]
                self.merged_dfs[master_name] = merged_df[new_order]

    # --- 阶段 2: 报告生成与架构诊断 ---
    def generate_report(self):
        benchmark_targets = {'roas': [2.0, True], 'cpm': [20.0, False], 'ctr': [0.015, True], 'cpc': [1.5, False], 'cpa': [30.0, False]}
        if self.bench_file:
            try:
                df_b = pd.read_excel(self.bench_file)
                benchmark_targets = extract_benchmark_values(df_b)
            except: pass

        self.doc.add_heading('广告投放深度分析报告', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.final_json = {"report_title": "广告投放深度分析报告", "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d")}

        # 1. 大盘总览 (保持不变)
        df_ov = pd.DataFrame()
        if "Master_Overview" in self.merged_dfs:
            df_src = self.merged_dfs["Master_Overview"]
            mask = df_src['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in ["分时", "Time"]))
            df_ov = df_src[mask].copy() if not df_src[mask].empty else df_src.copy()

        if not df_ov.empty:
            date_col = find_column_fuzzy(df_ov, ['date', 'time', '时间'])
            if date_col:
                try:
                    df_ov['temp_date'] = pd.to_datetime(df_ov[date_col], errors='coerce')
                    df_clean = df_ov.dropna(subset=['temp_date']).sort_values('temp_date')
                    dates = df_clean['temp_date'].unique()
                    raw_overall = calc_metrics_dict(df_clean)
                    if len(dates) >= 2:
                        mid_date = dates[len(dates)//2]
                        raw_prev = calc_metrics_dict(df_clean[df_clean['temp_date'] < mid_date])
                        raw_curr = calc_metrics_dict(df_clean[df_clean['temp_date'] >= mid_date])
                        raw_mom = {}
                        for k, v_curr in raw_curr.items():
                            if k == 'date_range': raw_mom[k] = "-"
                            else:
                                v_prev = raw_prev.get(k, 0)
                                raw_mom[k] = (v_curr - v_prev) / v_prev if v_prev > 0 else 0.0
                    else:
                        raw_prev = {k: "-" for k in raw_overall}; raw_curr = raw_overall; raw_mom = {k: "-" for k in raw_overall}

                    col_order = ["date_range", "spend", "roas", "cpa", "cpm", "cpc", "ctr", "cvr_purchase",
                                 "rate_click_to_lp", "rate_lp_to_atc", "rate_ic_to_pur", "aov", "add_to_cart", "purchases", "purchase_value"]
                    final_data = []
                    for label, r in zip(["整体数据", "上周期值", "本周期", "环比"], [raw_overall, raw_prev, raw_curr, raw_mom]):
                        row = {"Label": label}
                        is_m = (label == "环比")
                        for c in col_order: row[c] = format_cell(c, r.get(c, 0), is_mom=is_m)
                        row['date_range'] = label
                        final_data.append(row)

                    df_f = pd.DataFrame(final_data, columns=col_order)
                    df_f_display = apply_report_labels(df_f)
                    add_df_to_word(self.doc, df_f_display, "1. 数据大盘总览", level=1)
                    self.final_json['1_data_overview'] = df_f.to_dict(orient='records')

                    # 2. Benchmark
                    raw_current = calc_metrics_dict(df_clean)
                    bench_data = []
                    for metric_key in ['roas', 'cpm', 'ctr', 'cpc', 'cpa']:
                        curr_val = raw_current.get(metric_key, 0)
                        bench_val, higher_is_better = benchmark_targets.get(metric_key, [0, True])
                        conclusion = "-"
                        if curr_val != 0:
                            diff = curr_val - bench_val
                            if higher_is_better: conclusion = "✅ 优于大盘" if diff > 0 else ("⚠️ 低于大盘" if diff < 0 else "持平")
                            else: conclusion = "✅ 优于大盘" if diff < 0 else ("⚠️ 高于大盘" if diff > 0 else "持平")
                        bench_data.append({
                            "指标": REPORT_MAPPING.get(metric_key, metric_key.upper()),
                            "当前账户": format_cell(metric_key, curr_val),
                            "行业基准": format_cell(metric_key, bench_val),
                            "对比结论": conclusion
                        })
                    df_b = pd.DataFrame(bench_data)
                    add_df_to_word(self.doc, df_b, "2. 行业 Benchmark 对比", level=1)
                    self.final_json['2_industry_benchmark'] = df_b.to_dict(orient='records')
                except Exception as e: st.warning(f"大盘计算警告: {e}")

        # 3. 受众组
        self.doc.add_heading("3. 受众组分析", level=1)
        self.final_json['3_audience_analysis'] = {}
        audience_configs = [
            ("3.1 国家分析", ["国家", "Country"], True, "国家"),
            ("3.2 性别分析", ["性别", "Gender"], False, "性别"),
            ("3.3 年龄分析", ["年龄", "Age"], False, "年龄段"),
            ("3.4 受众组分析表", ["受众", "Audience"], True, "受众组名称"),
        ]

        if "Master_Breakdown" in self.merged_dfs:
            df_bd = self.merged_dfs["Master_Breakdown"]
            for title, keywords, top10, dim_label in audience_configs:
                mask = df_bd['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in keywords))
                df_curr = df_bd[mask].copy()
                if not df_curr.empty:
                    if not find_column_fuzzy(df_curr, ['cpc']): df_curr['cpc'] = df_curr['spend'] / df_curr['clicks'].replace(0, np.nan) if 'clicks' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['cpm']): df_curr['cpm'] = (df_curr['spend'] / df_curr['impressions'].replace(0, np.nan)) * 1000 if 'impressions' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['ctr']): df_curr['ctr'] = df_curr['clicks'] / df_curr['impressions'].replace(0, np.nan) if 'impressions' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['cpa']): df_curr['cpa'] = df_curr['spend'] / df_curr['purchases'].replace(0, np.nan) if 'purchases' in df_curr else 0

                    req_cols = ["dimension_item", "spend", "ctr", "cpc", "cpm", "cpa", "roas"]
                    # ✅ 现在，如果是受众表，这些字段已经被保留下来了
                    if "受众" in title: req_cols += ["converting_countries", "converting_keywords", "converting_genders", "converting_ages"]

                    rename_map = {}; valid_cols = []
                    for req in req_cols:
                        aliases = FIELD_ALIASES.get(req, [req])
                        found = find_column_fuzzy(df_curr, aliases)
                        if found: valid_cols.append(found); rename_map[found] = req
                        else: 
                            # 对于文本字段，给 "-" 而不是 0.0
                            default_val = "-" if "converting" in req else 0.0
                            df_curr[req] = default_val; valid_cols.append(req)

                    df_final = df_curr[valid_cols].rename(columns=rename_map)
                    if "dimension_item" in df_final.columns:
                         df_final = df_final[~df_final['dimension_item'].astype(str).str.lower().str.contains('unknow', na=False)]

                    if top10 and 'spend' in df_final.columns: df_final = df_final.sort_values('spend', ascending=False).head(10)
                    df_clean = df_final.round(2)
                    df_display = apply_report_labels(df_clean, custom_mapping={'dimension_item': dim_label})
                    add_df_to_word(self.doc, df_display, title, level=2)
                    self.final_json['3_audience_analysis'][title] = df_clean.to_dict(orient='records')

        # 4. 素材与落地页 (保持不变)
        if "Master_Creative" in self.merged_dfs:
            df_cr = self.merged_dfs["Master_Creative"]
            for title, keywords, label, json_key in [("4. 素材分析", ["素材", "Creative"], "素材名称", "4_creative_analysis"), ("6. 落地页分析", ["落地页", "Landing"], "落地页 URL", "6_landing_page_analysis")]:
                mask = df_cr['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in keywords))
                df_curr = df_cr[mask].copy()
                if not df_curr.empty:
                    if not find_column_fuzzy(df_curr, ['cpc']): df_curr['cpc'] = df_curr['spend'] / df_curr['clicks'].replace(0, np.nan) if 'clicks' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['cpa']): df_curr['cpa'] = df_curr['spend'] / df_curr['purchases'].replace(0, np.nan) if 'purchases' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['ctr']):
                         if 'impressions' in df_curr and 'clicks' in df_curr: df_curr['ctr'] = df_curr['clicks'] / df_curr['impressions'].replace(0, np.nan)
                         else: df_curr['ctr'] = np.nan
                    if 'cpc' in df_curr.columns and 'cpm' in df_curr.columns:
                        mask_fix = (df_curr['ctr'].isna() | (df_curr['ctr'] == 0)) & (df_curr['cpc'] > 0)
                        if mask_fix.any(): df_curr.loc[mask_fix, 'ctr'] = df_curr.loc[mask_fix, 'cpm'] / (df_curr.loc[mask_fix, 'cpc'] * 1000)
                    df_curr['ctr'] = df_curr['ctr'].fillna(0) * 100 

                    req_cols = ["content_item", "spend", "ctr", "cpc", "cpm", "roas", "cpa"]
                    rename_map = {}; valid_cols = []
                    for req in req_cols:
                        aliases = FIELD_ALIASES.get(req, [req])
                        found = find_column_fuzzy(df_curr, aliases)
                        if found: valid_cols.append(found); rename_map[found] = req
                        else: df_curr[req] = 0.0; valid_cols.append(req)
                    df_final = df_curr[valid_cols].rename(columns=rename_map)
                    if 'spend' in df_final.columns: df_final = df_final.sort_values('spend', ascending=False).head(10)
                    df_clean = df_final.round(2) 
                    df_display = apply_report_labels(df_clean, custom_mapping={'content_item': label})
                    add_df_to_word(self.doc, df_display, title, level=1)
                    self.final_json[json_key] = df_clean.to_dict(orient='records')
                    
        # 5. 版位 (保持不变)
        if "Master_Breakdown" in self.merged_dfs:
             self.doc.add_heading("5. 版位分析", level=1)
             df_bd = self.merged_dfs["Master_Breakdown"]
             mask = df_bd['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in ["版位", "Placement"]))
             df_curr = df_bd[mask].copy()
             if not df_curr.empty:
                 if not find_column_fuzzy(df_curr, ['cpc']): df_curr['cpc'] = df_curr['spend'] / df_curr['clicks'].replace(0, np.nan) if 'clicks' in df_curr else 0
                 if not find_column_fuzzy(df_curr, ['cpa']): df_curr['cpa'] = df_curr['spend'] / df_curr['purchases'].replace(0, np.nan) if 'purchases' in df_curr else 0
                 if not find_column_fuzzy(df_curr, ['ctr']): df_curr['ctr'] = df_curr['clicks'] / df_curr['impressions'].replace(0, np.nan) if 'impressions' in df_curr else 0
                 if not find_column_fuzzy(df_curr, ['cpm']): df_curr['cpm'] = (df_curr['spend'] / df_curr['impressions'].replace(0, np.nan)) * 1000 if 'impressions' in df_curr else 0
                 req_cols = ['dimension_item', 'spend', 'ctr', 'cpc', 'cpm', 'roas', 'cpa']
                 rename_map = {}; valid_cols = []
                 for c in req_cols:
                     aliases = FIELD_ALIASES.get(c, [c])
                     f = find_column_fuzzy(df_curr, aliases)
                     if f: valid_cols.append(f); rename_map[f] = c
                     else: df_curr[c] = 0.0; valid_cols.append(c)
                 df_clean = df_curr[valid_cols].rename(columns=rename_map).round(2)
                 df_top5 = df_clean.sort_values('spend', ascending=False).head(5)
                 add_df_to_word(self.doc, apply_report_labels(df_top5, {'dimension_item': '版位'}), "5.1 版位花费 TOP 5", level=2)
                 mean_ctr = df_clean['ctr'].mean(); mean_cpm = df_clean['cpm'].mean()
                 mask_pot = (df_clean['ctr'] > mean_ctr) & (df_clean['cpm'] < mean_cpm)
                 df_pot = df_clean[mask_pot].sort_values('ctr', ascending=False).head(5)
                 if df_pot.empty: df_pot = df_clean.sort_values('ctr', ascending=False).head(5)
                 add_df_to_word(self.doc, apply_report_labels(df_pot, {'dimension_item': '版位'}), "5.2 版位高潜力", level=2)
                 self.final_json['5_placement_analysis'] = {"top_spend": df_top5.to_dict('records'), "high_potential": df_pot.to_dict('records')}

        # 7. 架构诊断 (保持不变)
        rows = []
        if "Master_Overview" in self.merged_dfs:
             metrics = calc_metrics_dict(self.merged_dfs["Master_Overview"])
             if not metrics: metrics = {} 
             rows.append({
                "模块": "预算结构", 
                "当前结构数据表现": (
                    f"总花费: ${float(str(metrics.get('spend', 0)).replace(',', '')):,.2f}\n"
                    f"CPA: ${float(str(metrics.get('cpa', 0)).replace(',', '')):.2f}\n"
                    f"ROAS: {float(str(metrics.get('roas', 0)).replace(',', '')):.2f}"
                ), 
                "存在的问题": ""
             })
        if "Master_Breakdown" in self.merged_dfs:
            df_bd = self.merged_dfs["Master_Breakdown"]
            mask = df_bd['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in ["受众", "Audience"]))
            df_aud = df_bd[mask]
            s_col = find_column_fuzzy(df_aud, ['spend']); active_count = len(df_aud[df_aud[s_col] > 0]) if s_col else 0
            top_share = "0%"
            if not df_aud.empty and s_col:
                total_s = df_aud[s_col].sum()
                if total_s > 0: top_share = f"{df_aud[s_col].max()/total_s:.1%}"
            rows.append({"模块": "受众结构", "当前结构数据表现": f"活跃受众组数: {active_count}\nTop1 花费占比: {top_share}", "存在的问题": ""})
        if "Master_Creative" in self.merged_dfs:
             df_cr = self.merged_dfs["Master_Creative"]
             mask = df_cr['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in ["素材", "Creative"]))
             df_mat = df_cr[mask]
             s_col = find_column_fuzzy(df_mat, ['spend']); active_count = len(df_mat[df_mat[s_col] > 0]) if s_col else 0
             rows.append({"模块": "素材结构", "当前结构数据表现": f"活跃素材数: {active_count}", "存在的问题": ""})

        df_struct = pd.DataFrame(rows)
        add_df_to_word(self.doc, df_struct, "7. 广告架构分析", level=1)
        if "Master_Overview" in self.merged_dfs:
             self.final_json['7_structure_analysis'] = df_struct.to_dict(orient='records')

# ==========================================
# PART 4: Streamlit UI (保持不变)
# ==========================================
def set_artistic_style():
    st.markdown("""
        <style>
        /* 引入现代无衬线字体 */
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
        
        /* 全局重置 */
        .stApp {
            background-color: #FAFBFF; /* 极淡的蓝紫色背景 */
            font-family: 'Plus Jakarta Sans', sans-serif;
            color: #2D3748;
        }
        
        /* 标题样式 - 柔和渐变 */
        h1 {
            font-size: 3rem !important;
            font-weight: 800;
            background: linear-gradient(120deg, #845EC2, #D65DB1, #FF6F91);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -1px;
            margin-bottom: 10px;
            padding-bottom: 10px;
        }

        /* 装饰性分割线 - 柔光 */
        .art-divider {
            height: 2px;
            background: linear-gradient(90deg, rgba(132, 94, 194, 0.1), rgba(255, 111, 145, 0.5), rgba(132, 94, 194, 0.1));
            margin: 20px 0 50px 0;
            border-radius: 1px;
        }

        /* 模块标题 (如 1 数据源输入) */
        .section-header {
            font-size: 1.1rem;
            font-weight: 600;
            color: #4A5568;
            background: rgba(255, 255, 255, 0.6);
            backdrop-filter: blur(10px);
            display: inline-block;
            padding: 8px 16px;
            border-radius: 20px; /* 大圆角 */
            margin-bottom: 15px;
            border: 1px solid rgba(255, 255, 255, 0.8);
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        }
        .section-header span {
            display: inline-block;
            background: linear-gradient(120deg, #845EC2, #FF9671);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-weight: 800;
            margin-right: 8px;
        }

        /* 卡片容器 (功能指南) - 悬浮玻璃态 */
        .guide-card {
            background: #FFFFFF;
            border: 1px solid #EDF2F7;
            padding: 30px;
            border-radius: 24px; /* 大圆角 */
            position: relative;
            box-shadow: 0 10px 30px -5px rgba(132, 94, 194, 0.08); /* 柔和紫色阴影 */
            transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
            margin-bottom: 40px;
            overflow: hidden;
        }
        .guide-card::before {
            content: "";
            position: absolute;
            top: 0; left: 0; width: 100%; height: 6px;
            background: linear-gradient(90deg, #845EC2, #D65DB1, #FF9671, #FFC75F);
        }
        .guide-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 20px 40px -5px rgba(132, 94, 194, 0.15);
        }
        .guide-title {
            font-weight: 700;
            font-size: 1.1rem;
            color: #2D3748;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
        }

        /* 文件上传组件重构 - 极简白 */
        [data-testid='stFileUploader'] {
            background-color: #FFFFFF;
            border: 1px dashed #CBD5E0;
            border-radius: 16px;
            padding: 25px;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        }
        [data-testid='stFileUploader']:hover {
            border-color: #D65DB1; /* 悬浮变为粉紫色 */
            background-color: #FDFAFC;
            transform: translateY(-2px);
            box-shadow: 0 12px 24px -10px rgba(214, 93, 177, 0.2);
        }
        [data-testid='stFileUploader'] section {
            padding: 0;
        }
        /* 上传按钮样式覆盖 */
        [data-testid='stFileUploader'] button {
            border: 1px solid #E2E8F0;
            color: #4A5568;
            font-weight: 500;
            border-radius: 8px;
            box-shadow: none;
            background: white;
        }

        /* 主按钮 (开始处理) - 强渐变 */
        div.stButton > button:first-child {
            background: linear-gradient(135deg, #845EC2 0%, #D65DB1 100%);
            color: white;
            border: none;
            padding: 16px 40px;
            font-size: 18px;
            border-radius: 50px; /* 胶囊形状 */
            width: 100%;
            font-weight: 700;
            letter-spacing: 0.5px;
            box-shadow: 0 10px 20px -5px rgba(132, 94, 194, 0.4);
            transition: all 0.3s ease;
        }
        div.stButton > button:first-child:hover {
            background: linear-gradient(135deg, #956FD3 0%, #E76EC2 100%);
            box-shadow: 0 15px 30px -5px rgba(132, 94, 194, 0.5);
            transform: translateY(-3px) scale(1.02);
        }
        div.stButton > button:first-child:active {
            transform: translateY(1px);
        }

        /* 下载按钮群 - 柔和卡片风格 */
        [data-testid="stDownloadButton"] button {
            background-color: #FFFFFF;
            color: #4A5568;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            font-weight: 600;
            padding: 10px 20px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            transition: all 0.3s;
        }
        [data-testid="stDownloadButton"] button:hover {
            background: linear-gradient(135deg, #FF9671 0%, #FFC75F 100%);
            color: white;
            border-color: transparent;
            box-shadow: 0 8px 15px -3px rgba(255, 150, 113, 0.4);
            transform: translateY(-2px);
        }

        /* 状态提示框美化 */
        .stAlert {
            background-color: #FFFFFF;
            border: none;
            border-radius: 12px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
            border-left: 4px solid #845EC2;
        }
        
        /* 进度条颜色 */
        .stProgress > div > div > div > div {
            background: linear-gradient(90deg, #845EC2, #D65DB1, #FF9671);
        }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# 🧩 Mock Logic (逻辑层)
# ==========================================
class MockProcessor:
    def __init__(self, raw, bench):
        self.raw = raw
        self.bench = bench
        self.merged_dfs = {}
        self.final_json = {}
        class MockDoc:
            def save(self, b): b.write(b"content")
        self.doc = MockDoc()

    def run_pipeline(self):
        time.sleep(1.5)
        self.merged_dfs = {"Overview": pd.DataFrame({'A': [1,2], 'B': [3,4]})}
        self.final_json = {"status": "success", "score": 98}

# ==========================================
# 🚀 主程序
# ==========================================
def main():
    st.set_page_config(page_title="Ad-Opt System", layout="wide", page_icon="📊")
    set_artistic_style()

    # --- Header ---
    c1, c2 = st.columns([0.1, 0.9])
    with c1:
        st.title("广告优化报告数据生产系统")
    
    # 艺术分割线
    st.markdown('<div class="art-divider"></div>', unsafe_allow_html=True)

    # --- Guide Section (自定义 HTML 结构) ---
    st.markdown("""
        <div class="guide-card">
            <div class="guide-title">⚡ 功能使用指南 / USER GUIDE</div>
            <ul style="list-style-type: square; padding-left: 20px; color: #444; font-family: monospace;">
                <li style="margin-bottom: 8px;">请上传 <strong>[周期性复盘报告]</strong> 与 <strong>[行业 Benchmark]</strong> 源文件。</li>
                <li style="margin-bottom: 8px;">系统将自动执行：数据清洗 -> ETL处理 -> 大模型架构生成。</li>
                <li>输出包含：JSON (AI分析用)、Excel (透视用)、Word (审查用)。</li>
            </ul>
        </div>
    """, unsafe_allow_html=True)

    # --- Input Section ---
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown('<div class="section-header"><span>1️⃣ 数据源输入</span></div>', unsafe_allow_html=True)
        f1 = st.file_uploader("上传 [数据报表] (Excel)", type=["xlsx", "xls"], key="f1")
        if f1:
            st.markdown(f"<div style='color:#00D1FF; font-weight:bold; margin-top:5px;'>✔ 已加载: {f1.name}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div style='color:#999; font-size:0.8rem; margin-top:5px;'>等待上传...</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="section-header"><span>2️⃣ 行业基准</span></div>', unsafe_allow_html=True)
        f2 = st.file_uploader("上传 [行业 Benchmark]", type=["xlsx", "xls"], key="f2")
        if f2:
            st.markdown(f"<div style='color:#00D1FF; font-weight:bold; margin-top:5px;'>✔ 已加载: {f2.name}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div style='color:#999; font-size:0.8rem; margin-top:5px;'>可选 (若无则使用默认基准)</div>", unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)

    # --- Action Area ---
    # 居中布局
    _, btn_col, _ = st.columns([1, 1.5, 1])
    with btn_col:
        start = st.button("🚀 开始处理数据 / START PROCESS")

    # --- Processing ---
    if start:
        if not f1:
            st.error("⚠️ 缺少核心数据源！MISSING DATA SOURCE")
        else:
            processor = MockProcessor(f1, f2)
            
            # 自定义进度条容器
            with st.status("🔮 正在进行数据炼金术...", expanded=True) as status:
                st.write("✨ 清洗脏数据...")
                time.sleep(0.8)
                st.write("🔥 融合基准指标...")
                time.sleep(0.8)
                processor.run_pipeline()
                status.update(label="✅ 处理完成 / COMPLETED", state="complete", expanded=False)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="section-header"><span>3️⃣ 结果下载 / DOWNLOAD</span></div>', unsafe_allow_html=True)

            d1, d2, d3 = st.columns(3, gap="medium")
            with d1:
                st.download_button("💾 下载 JSON", data="{}", file_name="data.json", use_container_width=True)
            with d2:
                st.download_button("📊 下载 Excel", data=b"xls", file_name="data.xlsx", use_container_width=True)
            with d3:
                st.download_button("📝 下载 Word", data=b"doc", file_name="report.docx", use_container_width=True)

if __name__ == "__main__":
    main()
