# -*- coding: utf-8 -*-
"""
新人数据自动化分析系统
上传包含周会数据的zip文件夹，自动生成分析表格和图表
"""

import streamlit as st
import pandas as pd
import numpy as np
import xlrd
import xlwt
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
import os
import zipfile
import tempfile
import shutil
from io import BytesIO

# ============================================
# 页面配置
# ============================================
st.set_page_config(
    page_title="新人数据自动化分析",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 中文字体
def get_chinese_font():
    font_paths = [
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/wqy-zenhei/wqy-zenhei.ttc',
        'C:/Windows/Fonts/simhei.ttf',
        'C:/Windows/Fonts/msyh.ttf',
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            return font_manager.FontProperties(fname=fp)
    return font_manager.FontProperties()

chinese_font = get_chinese_font()

# ============================================
# 工具函数
# ============================================
def is_linghang(s):
    """判断是否为领航"""
    return any(ord(c) == 0x9886 for c in str(s))

def classify_interview(status):
    """分类面试状态"""
    if len(status) == 4:
        third_cp = hex(ord(status[2]))
        if third_cp == '0x901a':  # 通
            return 'pass'
    elif len(status) == 5:
        return 'fail'  # 面试不通过
    return 'other'

def find_file(base_dir, filename_patterns):
    """在目录中查找匹配的文件"""
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            for pattern in filename_patterns:
                if pattern in f:
                    return os.path.join(root, f)
    return None

def find_dir(base_dir, dir_name_patterns):
    """在目录中查找匹配的子目录"""
    for root, dirs, files in os.walk(base_dir):
        for d in dirs:
            for pattern in dir_name_patterns:
                if pattern in d:
                    return os.path.join(root, d)
    return None

# ============================================
# 核心处理函数
# ============================================
def process_data(data_dir, new_week):
    """
    处理数据，生成所有表格和图表
    返回: (results_dict, output_files_dict)
    """
    results = {}
    output_files = {}

    # --- 查找文件 ---
    # 本周新增数据
    new_data_dir = find_dir(data_dir, ['本周新增数据'])
    # 累计相关数据
    cumulative_dir = find_dir(data_dir, ['累计相关数据'])
    # 往期数据底表
    historical_dir = find_dir(data_dir, ['往期数据底表'])

    if not new_data_dir:
        raise FileNotFoundError("未找到'本周新增数据'文件夹")
    if not cumulative_dir:
        raise FileNotFoundError("未找到'累计相关数据'文件夹")
    if not historical_dir:
        raise FileNotFoundError("未找到'往期数据底表'文件夹")

    # 1. 面试通过-新人总名单
    pass_file = find_file(cumulative_dir, ['面试通过-新人总名单'])
    if not pass_file:
        raise FileNotFoundError("未找到'面试通过-新人总名单'文件")

    wb_pass = xlrd.open_workbook(pass_file)
    ws_pass = wb_pass.sheet_by_index(0)
    pass_uids = []
    pass_suppliers = []
    for r in range(1, ws_pass.nrows):
        uid = int(ws_pass.cell_value(r, 0))
        supplier = ws_pass.cell_value(r, 1)
        pass_uids.append(uid)
        pass_suppliers.append(supplier)

    # 2. 3月以来分层
    layer_file = find_file(cumulative_dir, ['3月以来分层'])
    if not layer_file:
        raise FileNotFoundError("未找到'3月以来分层'文件")

    wb_layer = xlrd.open_workbook(layer_file)
    layer_sheets = wb_layer.sheet_names()
    layer_data = {}
    for sheet_name in layer_sheets:
        ws = wb_layer.sheet_by_name(sheet_name)
        for r in range(1, ws.nrows):
            uid = int(ws.cell_value(r, 0))
            user_lv = ws.cell_value(r, 13)
            if uid not in layer_data:
                layer_data[uid] = {}
            layer_data[uid][sheet_name] = user_lv

    # 3. 本周新人漏斗明细
    funnel_file = find_file(new_data_dir, ['本周新人漏斗明细'])
    if not funnel_file:
        raise FileNotFoundError("未找到'本周新人漏斗明细'文件")

    df_funnel = pd.read_excel(funnel_file, engine='openpyxl')
    funnel_intern_map = {}
    for r in range(len(df_funnel)):
        uid_val = df_funnel.iloc[r, 1]
        if pd.isna(uid_val):
            continue
        uid = int(uid_val)
        status = df_funnel.iloc[r, 17]
        if uid not in funnel_intern_map:
            funnel_intern_map[uid] = status

    # 4. 本周领航+提交人数
    lh_submit_file = find_file(new_data_dir, ['本周领航+提交人数', '本周领航+提交'])
    if lh_submit_file:
        df_lh_submit = pd.read_excel(lh_submit_file, engine='xlrd')
        lh_submit_count = len(df_lh_submit)
    else:
        lh_submit_count = 0

    # 5. 本周供应商试录题通过人数
    supp_trial_file = find_file(new_data_dir, ['本周供应商试录题通过人数'])
    if supp_trial_file:
        df_supp_trial = pd.read_excel(supp_trial_file, engine='openpyxl')
        supp_submit_count = len(df_supp_trial)
        supp_pass_count = sum(1 for v in df_supp_trial.iloc[:, 6]
                            if isinstance(v, str) and len(v) > 0 and ord(v[0]) == 0x662f)
    else:
        supp_submit_count = 0
        supp_pass_count = 0

    # 6. 本周领航试录题通过人数
    lh_trial_file = find_file(new_data_dir, ['本周领航试录题通过人数'])
    lh_trial_uids = []
    if lh_trial_file:
        import csv
        with open(lh_trial_file, 'r', encoding='gbk') as f:
            reader = csv.DictReader(f)
            headers = reader.fieldnames
            key_uid = headers[0]
            for row in reader:
                try:
                    lh_trial_uids.append(int(row[key_uid]))
                except:
                    pass
    lh_pass_count = len(lh_trial_uids)

    # 7. 本周面试量
    interview_file = find_file(new_data_dir, ['面试量'])
    interview_uids = []
    interview_suppliers = []
    interview_statuses = []
    if interview_file:
        import csv
        with open(interview_file, 'r', encoding='gbk') as f:
            reader = csv.DictReader(f)
            iv_headers = reader.fieldnames
            key_iv_supplier = iv_headers[2]
            key_iv_status = iv_headers[6]
            for row in reader:
                interview_uids.append(int(row[iv_headers[0]]))
                interview_suppliers.append(row[key_iv_supplier])
                interview_statuses.append(row[key_iv_status])

    # --- 计算指标 ---
    # 本周试录题
    lh_submit = lh_submit_count
    supp_submit = supp_submit_count
    total_submit = lh_submit + supp_submit

    lh_trial_pass = lh_pass_count
    supp_trial_pass = supp_pass_count
    total_trial_pass = lh_trial_pass + supp_trial_pass

    # 本周面试
    lh_interview_total = 0
    lh_interview_pass = 0
    supp_interview_total = 0
    supp_interview_pass = 0

    for supplier, status in zip(interview_suppliers, interview_statuses):
        result = classify_interview(status)
        if result == 'other':
            continue
        if is_linghang(supplier):
            lh_interview_total += 1
            if result == 'pass':
                lh_interview_pass += 1
        else:
            supp_interview_total += 1
            if result == 'pass':
                supp_interview_pass += 1

    total_interview = lh_interview_total + supp_interview_total
    total_interview_pass = lh_interview_pass + supp_interview_pass

    # 累计面试通过
    lh_pass_cumulative = sum(1 for s in pass_suppliers if is_linghang(s))
    supp_pass_cumulative = len(pass_suppliers) - lh_pass_cumulative
    total_pass_cumulative = len(pass_suppliers)

    # v1v2分析
    latest_sheet = layer_sheets[0] if layer_sheets else None
    older_sheets = layer_sheets[1:] if len(layer_sheets) > 1 else []

    lh_this_week_v1v2 = 0
    supp_this_week_v1v2 = 0
    lh_cumulative_v1v2 = 0
    supp_cumulative_v1v2 = 0

    for uid, supplier in zip(pass_uids, pass_suppliers):
        is_lh = is_linghang(supplier)

        ever_v1v2 = False
        if uid in layer_data:
            for sheet_name, lv in layer_data[uid].items():
                if lv in ('v1', 'v2'):
                    ever_v1v2 = True
                    break

        new_v1v2 = False
        if uid in layer_data and latest_sheet:
            latest_lv = layer_data[uid].get(latest_sheet, None)
            if latest_lv in ('v1', 'v2'):
                had_v1v2_before = False
                for sheet_name in older_sheets:
                    old_lv = layer_data[uid].get(sheet_name, None)
                    if old_lv in ('v1', 'v2'):
                        had_v1v2_before = True
                        break
                if not had_v1v2_before:
                    new_v1v2 = True

        if ever_v1v2:
            if is_lh:
                lh_cumulative_v1v2 += 1
            else:
                supp_cumulative_v1v2 += 1

        if new_v1v2:
            if is_lh:
                lh_this_week_v1v2 += 1
            else:
                supp_this_week_v1v2 += 1

    total_this_week_v1v2 = lh_this_week_v1v2 + supp_this_week_v1v2
    total_cumulative_v1v2 = lh_cumulative_v1v2 + supp_cumulative_v1v2

    # 累计转正
    CONVERTED_STATUS = ''.join([chr(cp) for cp in [0x5b9e, 0x4e60, 0x901a, 0x8fc7]])
    lh_conversion = 0
    supp_conversion = 0
    for uid, supplier in zip([int(u) for u in pass_uids], pass_suppliers):
        if uid in funnel_intern_map:
            status = funnel_intern_map[uid]
            if status == CONVERTED_STATUS:
                if is_linghang(supplier):
                    lh_conversion += 1
                else:
                    supp_conversion += 1

    total_conversion = lh_conversion + supp_conversion

    # 计算比率
    lh_conversion_rate = lh_conversion / lh_pass_cumulative if lh_pass_cumulative > 0 else 0
    supp_conversion_rate = supp_conversion / supp_pass_cumulative if supp_pass_cumulative > 0 else 0
    total_conversion_rate = total_conversion / total_pass_cumulative if total_pass_cumulative > 0 else 0

    lh_v1v2_ratio = lh_cumulative_v1v2 / lh_conversion if lh_conversion > 0 else 0
    supp_v1v2_ratio = supp_cumulative_v1v2 / supp_conversion if supp_conversion > 0 else 0
    total_v1v2_ratio = total_cumulative_v1v2 / total_conversion if total_conversion > 0 else 0

    lh_trial_rate = lh_trial_pass / lh_submit if lh_submit > 0 else 0
    supp_trial_rate = supp_trial_pass / supp_submit if supp_submit > 0 else 0
    total_trial_rate = total_trial_pass / total_submit if total_submit > 0 else 0

    lh_interview_rate = lh_interview_pass / lh_interview_total if lh_interview_total > 0 else 0
    supp_interview_rate = supp_interview_pass / supp_interview_total if supp_interview_total > 0 else 0
    total_interview_rate = total_interview_pass / total_interview if total_interview > 0 else 0

    # --- 保存结果 ---
    results['table1'] = {
        'headers': ['来源', '本周试录题提交人数', '本周试录题通过人数', '本周面试人数',
                   '本周面试通过人数', '本周转正人数', '本周v1v2人数',
                   '累计面试通过人数', '累计转正人数', '累计v1v2人数', '转正率', 'v1v2占比'],
        'lh': ['领航', lh_submit, lh_trial_pass, lh_interview_total, lh_interview_pass,
               '', lh_this_week_v1v2, lh_pass_cumulative, lh_conversion,
               lh_cumulative_v1v2, round(lh_conversion_rate, 4), round(lh_v1v2_ratio, 4)],
        'supp': ['其他供应商', supp_submit, supp_trial_pass, supp_interview_total, supp_interview_pass,
                 '', supp_this_week_v1v2, supp_pass_cumulative, supp_conversion,
                 supp_cumulative_v1v2, round(supp_conversion_rate, 4), round(supp_v1v2_ratio, 4)],
        'total': ['总计', total_submit, total_trial_pass, total_interview, total_interview_pass,
                  '', total_this_week_v1v2, total_pass_cumulative, total_conversion,
                  total_cumulative_v1v2, round(total_conversion_rate, 4), round(total_v1v2_ratio, 4)]
    }

    results['table2'] = {
        'headers': ['来源', '本周试录题提交人数', '本周试录题通过人数', '本周面试人数',
                   '本周面试通过人数', '试录题合格率', '面试通过率'],
        'lh': ['领航', lh_submit, lh_trial_pass, lh_interview_total, lh_interview_pass,
               round(lh_trial_rate, 4), round(lh_interview_rate, 4)],
        'supp': ['其他供应商', supp_submit, supp_trial_pass, supp_interview_total, supp_interview_pass,
                 round(supp_trial_rate, 4), round(supp_interview_rate, 4)],
        'total': ['总计', total_submit, total_trial_pass, total_interview, total_interview_pass,
                  round(total_trial_rate, 4), round(total_interview_rate, 4)]
    }

    # 历史数据
    weeks = ['0226-0304', '0305-0311', '0312-0318', '0319-0325', '0326-0401',
             '0402-0408', '0409-0415', '0416-0422', '0423-0429', '0430-0506',
             '0507-0513', '0514-0520', '0521-0527', '0528-0603', '0604-0610', '0611-0617']

    trial_supplier_hist = [0.268, 0.2134, 0.3571, 0.3736, 0.269, 0.3824, 0.2923, 0.1573,
                          0.3091, 0.2903, 0.44, 0.2468, 0.3636, 0.3171, 0.4773, 0.4615]
    trial_lh_hist = [0.1648, 0.2308, 0.2544, 0.3645, 0.0994, 0.3038, 0.4286, 0.5385,
                     0.6333, 0.5, 0.7941, 0.7436, 0.52, 0.7368, 0.8636, 1.0]
    trial_combined_hist = [0.2181, 0.219, 0.3194, 0.3702, 0.1867, 0.3481, 0.34, 0.2979,
                          0.4235, 0.3333, 0.5833, 0.4138, 0.4203, 0.45, 0.6061, 0.65]

    int_supplier_hist = [0.444, 0.423, 0.41, 0.415, 0.326, 0.303, 0.381, 0.419,
                        0.313, 0, 0.381, 0.556, 0.429, 0.583, 0.563, 0.375]
    int_lh_hist = [0.429, 0.417, 0.133, 0.393, 0.367, 0.556, 0.429, 0.4,
                   0.357, 0, 0.429, 0.357, 0.333, 0.692, 0.25, 0.6]
    int_combined_hist = [0.439, 0.421, 0.333, 0.406, 0.342, 0.357, 0.397, 0.411,
                        0.333, 0, 0.405, 0.435, 0.391, 0.64, 0.458, 0.5]

    all_weeks = weeks + [new_week]
    all_trial_supp = trial_supplier_hist + [round(supp_trial_rate, 4)]
    all_trial_lh = trial_lh_hist + [round(lh_trial_rate, 4)]
    all_trial_comb = trial_combined_hist + [round(total_trial_rate, 4)]

    all_int_supp = int_supplier_hist + [round(supp_interview_rate, 4)]
    all_int_lh = int_lh_hist + [round(lh_interview_rate, 4)]
    all_int_comb = int_combined_hist + [round(total_interview_rate, 4)]

    results['trial_trend'] = {
        'weeks': all_weeks,
        'supplier': all_trial_supp,
        'lh': all_trial_lh,
        'combined': all_trial_comb
    }

    results['interview_trend'] = {
        'weeks': all_weeks,
        'supplier': all_int_supp,
        'lh': all_int_lh,
        'combined': all_int_comb
    }

    # 分层版数据
    results['layered_data'] = {
        'uids': pass_uids,
        'suppliers': pass_suppliers,
        'intern_status': [funnel_intern_map.get(int(u), '') for u in pass_uids],
        'layer_sheets': layer_sheets,
        'layer_data': layer_data
    }

    # --- 生成文件 ---
    output_dir = tempfile.mkdtemp()

    # 生成Excel文件
    # Table 1
    wb1 = xlwt.Workbook()
    ws1 = wb1.add_sheet('Sheet1')
    for i, h in enumerate(results['table1']['headers']):
        ws1.write(0, i, h)
    for row_idx, row_data in enumerate([results['table1']['lh'],
                                        results['table1']['supp'],
                                        results['table1']['total']], 1):
        for i, v in enumerate(row_data):
            ws1.write(row_idx, i, v)
    file1 = os.path.join(output_dir, '新人指标监控全量表.xls')
    wb1.save(file1)
    output_files['table1'] = file1

    # Table 2
    wb2 = xlwt.Workbook()
    ws2 = wb2.add_sheet('Sheet1')
    for i, h in enumerate(results['table2']['headers']):
        ws2.write(0, i, h)
    for row_idx, row_data in enumerate([results['table2']['lh'],
                                        results['table2']['supp'],
                                        results['table2']['total']], 1):
        for i, v in enumerate(row_data):
            ws2.write(row_idx, i, v)
    file2 = os.path.join(output_dir, '本周新人情况一览表.xls')
    wb2.save(file2)
    output_files['table2'] = file2

    # Trial trend
    wb3 = xlwt.Workbook()
    ws3 = wb3.add_sheet('Sheet1')
    ws3.write(0, 0, '试录题合格率')
    ws3.write(1, 0, '周度')
    for i, w in enumerate(all_weeks):
        ws3.write(1, i+1, w)
    ws3.write(2, 0, '供应商')
    for i, v in enumerate(all_trial_supp):
        ws3.write(2, i+1, v)
    ws3.write(3, 0, '领航')
    for i, v in enumerate(all_trial_lh):
        ws3.write(3, i+1, v)
    ws3.write(4, 0, '综合')
    for i, v in enumerate(all_trial_comb):
        ws3.write(4, i+1, v)
    file3 = os.path.join(output_dir, '试录题合格率变化趋势-新.xls')
    wb3.save(file3)
    output_files['trial_trend'] = file3

    # Interview trend
    wb4 = xlwt.Workbook()
    ws4 = wb4.add_sheet('Sheet1')
    ws4.write(0, 0, '面试通过率')
    ws4.write(1, 0, '周度')
    for i, w in enumerate(all_weeks):
        ws4.write(1, i+1, w)
    ws4.write(2, 0, '供应商')
    for i, v in enumerate(all_int_supp):
        ws4.write(2, i+1, v)
    ws4.write(3, 0, '领航')
    for i, v in enumerate(all_int_lh):
        ws4.write(3, i+1, v)
    ws4.write(4, 0, '综合')
    for i, v in enumerate(all_int_comb):
        ws4.write(4, i+1, v)
    file4 = os.path.join(output_dir, '面试通过率趋势-新.xls')
    wb4.save(file4)
    output_files['interview_trend'] = file4

    # Layered data
    wb5 = xlwt.Workbook()
    ws5 = wb5.add_sheet('Sheet1')
    layer_headers = ['UID', '供应商', '实习状态'] + layer_sheets
    for i, h in enumerate(layer_headers):
        ws5.write(0, i, h)
    for idx, (uid, supplier) in enumerate(zip(pass_uids, pass_suppliers)):
        row_idx = idx + 1
        ws5.write(row_idx, 0, uid)
        ws5.write(row_idx, 1, supplier)
        intern_status = funnel_intern_map.get(uid, '')
        ws5.write(row_idx, 2, intern_status)
        if uid in layer_data:
            for col_idx, sheet_name in enumerate(layer_sheets):
                lv = layer_data[uid].get(sheet_name, '')
                ws5.write(row_idx, 3 + col_idx, lv)
    file5 = os.path.join(output_dir, '面试通过-新人总名单-分层版.xls')
    wb5.save(file5)
    output_files['layered'] = file5

    # --- 生成图表 ---
    x = list(range(len(all_weeks)))

    # Chart 1: 试录题合格率
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.plot(x, all_trial_supp, 'o-', label='供应商', linewidth=2, markersize=5)
    ax.plot(x, all_trial_lh, 's-', label='领航', linewidth=2, markersize=5)
    ax.plot(x, all_trial_comb, '^-', label='综合', linewidth=2, markersize=5)
    ax.set_xlabel('周度', fontproperties=chinese_font, fontsize=12)
    ax.set_ylabel('合格率', fontproperties=chinese_font, fontsize=12)
    ax.set_title('试录题合格率变化趋势', fontproperties=chinese_font, fontsize=14)
    ax.set_xticks(x)
    ax.set_xticklabels(all_weeks, rotation=45, ha='right', fontsize=8)
    ax.legend(prop=chinese_font, fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_ylim(0, 1.1)
    chart1 = os.path.join(output_dir, '试录题合格率变化趋势-新.png')
    plt.tight_layout()
    plt.savefig(chart1, dpi=150, bbox_inches='tight')
    plt.close()
    output_files['trial_chart'] = chart1

    # Chart 2: 面试通过率
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.plot(x, all_int_supp, 'o-', label='供应商', linewidth=2, markersize=5)
    ax.plot(x, all_int_lh, 's-', label='领航', linewidth=2, markersize=5)
    ax.plot(x, all_int_comb, '^-', label='综合', linewidth=2, markersize=5)
    ax.set_xlabel('周度', fontproperties=chinese_font, fontsize=12)
    ax.set_ylabel('通过率', fontproperties=chinese_font, fontsize=12)
    ax.set_title('面试通过率走势', fontproperties=chinese_font, fontsize=14)
    ax.set_xticks(x)
    ax.set_xticklabels(all_weeks, rotation=45, ha='right', fontsize=8)
    ax.legend(prop=chinese_font, fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_ylim(0, 1.1)
    chart2 = os.path.join(output_dir, '面试通过率走势-新.png')
    plt.tight_layout()
    plt.savefig(chart2, dpi=150, bbox_inches='tight')
    plt.close()
    output_files['interview_chart'] = chart2

    return results, output_files


# ============================================
# Streamlit UI
# ============================================
st.title(" 新人数据自动化分析系统")
st.markdown("""
上传包含周会数据的文件夹（zip格式），自动生成分析表格和图表。

**文件夹应包含以下子目录：**
- `本周新增数据/` - 包含本周的各项数据文件
- `累计相关数据/` - 包含历史累计数据
- `往期数据底表/` - 包含往期趋势数据
""")

# 文件上传
uploaded_file = st.file_uploader("上传zip文件", type=['zip'])

# 参数输入
col1, col2 = st.columns(2)
with col1:
    new_week = st.text_input("本周周期", value="0618-0624", help="格式：MMDD-MMDD")
with col2:
    st.write("")  # 占位

if uploaded_file is not None:
    with st.spinner("正在处理数据..."):
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp()

            # 解压zip文件
            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # 查找数据目录（可能在根目录或子目录中）
            data_dir = temp_dir
            # 检查是否直接包含所需目录
            if not find_dir(temp_dir, ['本周新增数据']):
                # 可能在第一层子目录中
                for item in os.listdir(temp_dir):
                    item_path = os.path.join(temp_dir, item)
                    if os.path.isdir(item_path):
                        if find_dir(item_path, ['本周新增数据']):
                            data_dir = item_path
                            break

            # 处理数据
            results, output_files = process_data(data_dir, new_week)

            st.success("✅ 数据处理完成！")

            # 显示结果
            st.subheader("1. 新人指标监控全量表")
            df_table1 = pd.DataFrame([
                results['table1']['lh'],
                results['table1']['supp'],
                results['table1']['total']
            ], columns=results['table1']['headers'])
            st.dataframe(df_table1, use_container_width=True)

            with open(output_files['table1'], 'rb') as f:
                st.download_button("📥 下载 新人指标监控全量表.xls", f,
                                  file_name="新人指标监控全量表.xls",
                                  mime="application/vnd.ms-excel")

            st.subheader("2. 本周新人情况一览表")
            df_table2 = pd.DataFrame([
                results['table2']['lh'],
                results['table2']['supp'],
                results['table2']['total']
            ], columns=results['table2']['headers'])
            st.dataframe(df_table2, use_container_width=True)

            with open(output_files['table2'], 'rb') as f:
                st.download_button("📥 下载 本周新人情况一览表.xls", f,
                                  file_name="本周新人情况一览表.xls",
                                  mime="application/vnd.ms-excel")

            st.subheader("3. 试录题合格率变化趋势")
            st.image(output_files['trial_chart'], use_container_width=True)

            with open(output_files['trial_trend'], 'rb') as f:
                st.download_button(" 下载 试录题合格率变化趋势-新.xls", f,
                                  file_name="试录题合格率变化趋势-新.xls",
                                  mime="application/vnd.ms-excel")

            st.subheader("4. 面试通过率走势")
            st.image(output_files['interview_chart'], use_container_width=True)

            with open(output_files['interview_trend'], 'rb') as f:
                st.download_button("📥 下载 面试通过率趋势-新.xls", f,
                                  file_name="面试通过率趋势-新.xls",
                                  mime="application/vnd.ms-excel")

            st.subheader("5. 面试通过-新人总名单-分层版")
            st.write(f"共 {len(results['layered_data']['uids'])} 条记录")

            with open(output_files['layered'], 'rb') as f:
                st.download_button("📥 下载 面试通过-新人总名单-分层版.xls", f,
                                  file_name="面试通过-新人总名单-分层版.xls",
                                  mime="application/vnd.ms-excel")

            # 清理临时文件
            shutil.rmtree(temp_dir, ignore_errors=True)

        except FileNotFoundError as e:
            st.error(f"❌ 文件缺失：{str(e)}")
        except Exception as e:
            st.error(f"❌ 处理出错：{str(e)}")
            import traceback
            st.code(traceback.format_exc())
