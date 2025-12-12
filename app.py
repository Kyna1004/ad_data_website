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
# PART 1: 配置区域
# ==========================================

# 字段展示名称映射 (中文 Key)
REPORT_MAPPING = {
    "spend": "花费 ($)", "roas": "ROAS", "purchases": "购买次数", "purchase_value": "购买总价值",
    "cpa": "CPA ($)", "ctr": "CTR (%)", "cpm": "CPM ($)", "cpc": "CPC ($)", "aov": "客单价",
    "impressions": "展现量", "clicks_all": "点击量 (All)", "clicks": "点击量 (All)", "ctr_all": "点击率 (All)",
    "landing_page_views": "落地页访问量", "add_to_cart": "加购次数", "initiate_checkout": "结账发起数 (IC)",
    "rate_click_to_lp": "点击 → 落地页访问转化率", "rate_lp_to_atc": "落地页 → 加购转化率",
    "rate_atc_to_ic": "加购 → 购买转化率", "rate_ic_to_pur": "购买转化率",
    "cvr_purchase": "点击 → 购买转化率", "cvr_lp_to_pur": "CVR (全站转化率)",
    "date_range": "日期/时段", "campaign_type": "投放模式", "adset_name": "广告组ID", "adset_id": "广告组ID",
    "custom_audience_settings": "自定义受众源", "converting_keywords": "高潜兴趣词", "audience_type": "受众策略",
    "country": "国家", "age_group": "年龄", "gender": "性别", "creative_name": "素材名称", "placement": "版位",
    "landing_page_url": "页面 URL", "mom_change": "环比波动", "anomaly_metric_name": "异常项",
    "converting_countries": "产生成效的国家", "converting_genders": "产生成效的性别", "converting_ages": "产生成效的年龄",
    "dimension_item": "维度名称", "content_item": "内容名称"
}

COMMON_METRICS = {
    "spend": ["花费金额(USD)", "花费金额 （USD）", "花费金额 (USD)", "花费金额", "Amount Spent"],
    "roas": ["广告花费回报 (ROAS) - 购物", "广告花费回报（ROAS）-购物", "ROAS", "Purchase ROAS"],
    "purchases": ["购买次数", "成效数量", "成效", "Purchases", "Results"],
    "cpa": ["单次购买费用", "单次购物成本", "单次成效成本", "单次成效费用", "Cost per Purchase"],
    "ctr": ["链接点击率", "链接点击率（%)", "链接点击率（%）", "CTR"],
    "cpm": ["千次展示费用", "CPM"],
    "clicks": ["点击", "链接点击", "Clicks (All)"],
    "impressions": ["曝光", "展示次数", "Impressions"],
    "purchase_value": ["购买价值", "购物价值", "Purchase Conversion Value"],
    "aov": ["单次购买价值", "单次购物价值"]
}

SHEET_MAPPINGS = {
    "整体数据": {
        **COMMON_METRICS,
        "date_range": ["时间范围", "Date", "Day"],
        "clicks_all": ["点击", "Clicks"],
        "landing_page_views": ["落地页浏览量", "Landing Page Views"],
        "add_to_cart": ["加入购物车", "Adds to Cart"],
        "initiate_checkout": ["结账发起次数", "Initiate Checkout"],
        "rate_click_to_lp": ["点击-落地页浏览转化率"],
        "rate_lp_to_atc": ["落地页浏览-加购转化率"],
        "rate_atc_to_ic": ["加购-结账转化率"],
        "rate_ic_to_pur": ["结账-购买转化率"]
    },
    "分时段数据": {
        **COMMON_METRICS,
        "date_range": ["时间范围", "Time of Day", "Hourly"],
        "landing_page_views": ["落地页浏览量"],
        "add_to_cart": ["加入购物车"],
        "initiate_checkout": ["结账发起次数"],
    },
    "异常指标": {
        "anomaly_metric_name": ["异常指标"],
        "mom_change": ["环比"]
    },
    "广告架构": {**COMMON_METRICS, "dimension_item": ["广告类型", "Campaign Name"]},
    "受众组": {
        **COMMON_METRICS,
        "dimension_item": ["广告组", "广告组Id", "Ad Set Name"],
        "custom_audience_settings": ["设置的自定义受众"],
        "converting_keywords": ["产生成效的关键词"]
    },
    "受众类型": {**COMMON_METRICS, "dimension_item": ["受众类型"]},
    "国家": {**COMMON_METRICS, "dimension_item": ["国家/地区", "国家", "Country", "Region"]},
    "年龄": {**COMMON_METRICS, "dimension_item": ["年龄", "Age"]},
    "性别": {**COMMON_METRICS, "dimension_item": ["性别", "Gender"]},
    "平台&版位": {**COMMON_METRICS, "dimension_item": ["平台&版位", "Placement", "Platform"]},
    "素材": {
        **COMMON_METRICS,
        "content_item": ["素材", "Ad Name", "Creative Name"],
        "cvr_lp_to_pur": ["落地页浏览-购买转化率"]
    },
    "落地页": {
        **COMMON_METRICS,
        "content_item": ["落地页url", "落地页", "Website URL"],
        "ctr_all": ["曝光-点击转化率"],
        "rate_lp_to_atc": ["落地页浏览-加购转化率", "落地页浏览-购物转化率"]
    }
}

GROUP_CONFIG = {
    "Master_Overview": ["整体数据", "分时段数据", "异常指标"],
    "Master_Breakdown": ["广告架构", "受众组", "受众类型", "国家", "年龄", "性别", "平台&版位"],
    "Master_Creative": ["素材", "落地页"]
}

FIELD_ALIASES = {
    "adset_id": ["adset_id", "ad set id", "adset id", "广告组编号", "广告组id", "adset_name", "ad set name"],
    "converting_countries": ["converting_countries", "country", "region", "国家", "地区"],
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
# PART 2: 核心工具函数
# ==========================================

def clean_numeric(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('$', '').replace('¥', '').replace(',', '')
    if '%' in val_str: 
        try: return float(val_str.replace('%', '')) / 100.0
        except: return 0.0
    try: return float(val_str)
    except: return 0.0

def clean_numeric_strict(val):
    if pd.isna(val): return 0.0
    val_str = str(val).strip().replace('$', '').replace('¥', '').replace(',', '')
    if '%' in val_str: val_str = val_str.replace('%', '')
    try: return float(val_str)
    except: return 0.0

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

    eps = 1e-9
    res['spend'] = sums['spend']
    res['roas'] = sums['purchase_value'] / (sums['spend'] + eps)
    res['cpm'] = (sums['spend'] / (sums['impressions'] + eps)) * 1000
    res['cpc'] = sums['spend'] / (sums['clicks'] + eps)
    res['ctr'] = sums['clicks'] / (sums['impressions'] + eps)
    res['cpa'] = sums['spend'] / (sums['purchases'] + eps)
    res['cvr_purchase'] = sums['purchases'] / (sums['clicks'] + eps)
    res['rate_click_to_lp'] = sums['landing_page_views'] / (sums['clicks'] + eps)
    res['rate_lp_to_atc'] = sums['add_to_cart'] / (sums['landing_page_views'] + eps)
    res['rate_atc_to_ic'] = sums['initiate_checkout'] / (sums['add_to_cart'] + eps)
    res['rate_ic_to_pur'] = sums['purchases'] / (sums['initiate_checkout'] + eps)
    res['aov'] = sums['purchase_value'] / (sums['purchases'] + eps)

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
    if any(x in k for x in ['rate', 'ctr', 'cvr', '点击率', '转化率']): return f"{val:.2%}"
    if any(x in k for x in ['spend', 'cpm', 'cpc', 'value', 'aov', 'cpa', '花费', '金额', '客单价', 'gmv']): return f"{val:,.2f}"
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
                if not pd.isna(v): extracted[metric] = [v, higher_better]
            except: pass
    return extracted

def remap_json_keys(obj, mapping):
    """递归将 JSON 对象中的英文 Key 替换为中文展示名"""
    if isinstance(obj, dict):
        new_dict = {}
        for k, v in obj.items():
            new_key = mapping.get(k, k)
            new_dict[new_key] = remap_json_keys(v, mapping)
        return new_dict
    elif isinstance(obj, list):
        return [remap_json_keys(i, mapping) for i in obj]
    else:
        return obj

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
        if any(x in str(col).lower() for x in ["url", "link", "素材", "内容", "content"]):
            link_col_idx = j
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
# PART 3: 主逻辑类
# ==========================================

class AdReportProcessor:
    def __init__(self, raw_file, bench_file=None):
        self.raw_file = raw_file
        self.bench_file = bench_file
        self.processed_dfs = {}
        self.merged_dfs = {}
        self.final_json = {}
        self.doc = Document()

    def process_etl(self):
        xls = pd.ExcelFile(self.raw_file)
        for sheet_name, mapping in SHEET_MAPPINGS.items():
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                final_cols = {}
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
                    text_cols = ['date_range', 'anomaly_metric_name', 'converting_keywords',
                                 'custom_audience_settings', 'dimension_item', 'content_item']
                    for col in df_clean.columns:
                        if col not in text_cols: df_clean[col] = df_clean[col].apply(clean_numeric)

                    if sheet_name in ["素材", "落地页","受众组"]:
                        if "spend" in df_clean.columns:
                            df_clean = df_clean.sort_values("spend", ascending=False).head(10)

                    df_clean["Source_Sheet"] = sheet_name
                    self.processed_dfs[sheet_name] = df_clean

        for master_name, source_sheets in GROUP_CONFIG.items():
            dfs_to_merge = [self.processed_dfs[src] for src in source_sheets if src in self.processed_dfs]
            if dfs_to_merge:
                merged_df = pd.concat(dfs_to_merge, ignore_index=True)
                cols = list(merged_df.columns)
                priority_cols = ['Source_Sheet', 'date_range', 'dimension_item', 'content_item',
                                 'spend', 'roas', 'purchases', 'cpa']
                new_order = [c for c in priority_cols if c in cols] + [c for c in cols if c not in priority_cols]
                self.merged_dfs[master_name] = merged_df[new_order]

    def generate_report(self):
        benchmark_targets = {'roas': [2.0, True], 'cpm': [20.0, False], 'ctr': [0.015, True], 'cpc': [1.5, False], 'cpa': [30.0, False]}
        if self.bench_file:
            try:
                df_b = pd.read_excel(self.bench_file)
                benchmark_targets = extract_benchmark_values(df_b)
            except: pass

        self.doc.add_heading('广告投放深度分析报告', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.final_json = {"report_title": "广告投放深度分析报告", "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d")}

        # 1. 大盘总览
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

                    # Benchmark
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
                    if "受众" in title: req_cols += ["converting_countries", "converting_keywords"]
                    rename_map = {}; valid_cols = []
                    for req in req_cols:
                        aliases = FIELD_ALIASES.get(req, [req])
                        found = find_column_fuzzy(df_curr, aliases)
                        if found: valid_cols.append(found); rename_map[found] = req
                        else: df_curr[req] = 0.0 if "converting" not in req else "-"; valid_cols.append(req)

                    df_final = df_curr[valid_cols].rename(columns=rename_map)
                    if "dimension_item" in df_final.columns:
                          df_final = df_final[~df_final['dimension_item'].astype(str).str.lower().str.contains('unknow', na=False)]
                    if top10 and 'spend' in df_final.columns: df_final = df_final.sort_values('spend', ascending=False).head(10)
                    
                    df_clean = df_final.round(2)
                    df_display = apply_report_labels(df_clean, custom_mapping={'dimension_item': dim_label})
                    add_df_to_word(self.doc, df_display, title, level=2)
                    self.final_json['3_audience_analysis'][title] = df_clean.to_dict(orient='records')

        # 4. 素材与落地页
        if "Master_Creative" in self.merged_dfs:
            df_cr = self.merged_dfs["Master_Creative"]
            for title, keywords, label, json_key in [("4. 素材分析", ["素材", "Creative"], "素材名称", "4_creative_analysis"), ("6. 落地页分析", ["落地页", "Landing"], "落地页 URL", "6_landing_page_analysis")]:
                mask = df_cr['Source_Sheet'].astype(str).apply(lambda x: any(k in x for k in keywords))
                df_curr = df_cr[mask].copy()
                if not df_curr.empty:
                    if not find_column_fuzzy(df_curr, ['cpc']): df_curr['cpc'] = df_curr['spend'] / df_curr['clicks'].replace(0, np.nan) if 'clicks' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['cpa']): df_curr['cpa'] = df_curr['spend'] / df_curr['purchases'].replace(0, np.nan) if 'purchases' in df_curr else 0
                    if not find_column_fuzzy(df_curr, ['ctr']): df_curr['ctr'] = df_curr['clicks'] / df_curr['impressions'].replace(0, np.nan) if 'impressions' in df_curr else 0

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

        # 5. 版位
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

        # 7. 架构诊断
        rows = []
        if "Master_Overview" in self.merged_dfs:
             metrics = calc_metrics_dict(self.merged_dfs["Master_Overview"])
             rows.append({"模块": "预算结构", "当前结构数据表现": f"总花费: ${metrics.get('spend',0):,.2f}\nCPA: ${metrics.get('cpa',0):.2f}\nROAS: {metrics.get('roas',0):.2f}", "存在的问题": ""})

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
        self.final_json['7_structure_audit'] = df_struct.to_dict(orient='records')


# ==========================================
# PART 4: Streamlit UI
# ==========================================
def main():
    st.set_page_config(page_title="Auto-Merge & Analysis V20.10", layout="wide")
    st.title("📊 广告数据清洗与降维合并 (Auto-Merge V20.10)")
    st.markdown("""
    **功能说明：**
    1. **ETL阶段**：自动进行字段映射、数值清洗、以及素材/落地页 Top10 截断。
    2. **报告阶段**：生成架构诊断表、Word 报告及 Gemini 分析用 JSON。
    """)

    col1, col2 = st.columns(2)
    with col1:
        raw_file = st.file_uploader("1. 上传 [数据报表] (Excel)", type=["xlsx", "xls"])
    with col2:
        bench_file = st.file_uploader("2. 上传 [行业Benchmark] (Excel, 可选)", type=["xlsx", "xls"])

    if st.button("🚀 开始执行全流程"):
        if not raw_file:
            st.error("请至少上传数据报表！")
            return

        processor = AdReportProcessor(raw_file, bench_file)

        try:
            with st.spinner("阶段 1/2: 数据清洗、Top10截断、降维合并..."):
                processor.process_etl()
                st.success("✅ 阶段 1 完成：Master Tables 已生成")
                
                with st.expander("查看降维合并后的数据 (Master Tables)"):
                    tabs = st.tabs(processor.merged_dfs.keys())
                    for i, (k, v) in enumerate(processor.merged_dfs.items()):
                        with tabs[i]: st.dataframe(v.head(20))

            with st.spinner("阶段 2/2: 生成架构诊断、Word报告 & JSON..."):
                processor.generate_report()
                st.success("✅ 阶段 2 完成：报告已生成")

            st.divider()

            c1, c2, c3 = st.columns(3)

            # 1. JSON (中文展示名 Key)
            final_json_display = remap_json_keys(processor.final_json, REPORT_MAPPING)
            json_str = json.dumps(final_json_display, indent=4, ensure_ascii=False, default=str)
            c1.download_button("📥 下载 JSON (Gemini用 - 中文Key)", json_str, "Ad_Report_Data.json", "application/json")

            # 2. Excel (中文展示名 Header)
            output_xls = io.BytesIO()
            with pd.ExcelWriter(output_xls, engine='xlsxwriter') as writer:
                for name, df in processor.merged_dfs.items(): 
                    # 重命名列头为中文
                    df_display = df.rename(columns=REPORT_MAPPING)
                    df_display.to_excel(writer, sheet_name=name, index=False)
            c2.download_button("📥 下载 Excel (合并后数据 - 中文头)", output_xls.getvalue(), "Merged_Ad_Report_Final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # 3. Word
            output_doc = io.BytesIO()
            processor.doc.save(output_doc)
            c3.download_button("📥 下载 Word (最终报告)", output_doc.getvalue(), "Ad_Report_Final_V20_10.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        except Exception as e:
            st.error(f"处理过程中发生错误: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()
