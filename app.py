import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 智能补货 (最终版)")
st.title("📦 Coupang 智能补货 (定制导出版)")
st.markdown("### 核心逻辑：分表极简展示 + 统一视觉标准")

# ==========================================
# 2. 列号配置 (请确认 Excel 实际位置)
# ==========================================
# A=0, B=1, C=2, D=3, E=4, F=5, G=6 ... M=12, N=13 ... R=17

# --- 1. 基础信息表 (Master) ---
IDX_M_CODE    = 0    # A列: 产品编码
IDX_M_SHOP    = 1    # B列: 店铺
IDX_M_COL_E   = 4    # E列: 基础信息E
IDX_M_COL_F   = 5    # F列: SKU名称
IDX_M_COST    = 6    # G列: 采购单价
IDX_M_ORANGE  = 3    # D列: 橙火ID
IDX_M_INBOUND = 12   # M列: 入库码
IDX_M_ACTIVE  = 13   # N列: 是否在做 (有Y=在做，大小写容错)

# --- 2. 销售表 (近7天) ---
IDX_7D_SKU    = 0    # A列: SKU/ID
IDX_7D_QTY    = 8    # I列: 销售数量

# --- 3. 火箭仓/橙火库存表 ---
IDX_INV_R_SKU = 2    # C列: SKU/ID
IDX_INV_R_QTY = 7    # H列: 数量
IDX_INV_R_FEE = 17   # R列: 本月仓储费

# --- 4. 极风库存表 ---
IDX_INV_J_BAR = 2    # C列: 条码/入库码
IDX_INV_J_QTY = 10   # K列: 数量

# ==========================================
# 3. 工具函数
# ==========================================
def clean_match_key(series):
    s = series.astype(str).str.upper()
    s = s.str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip()
    s = s.replace('NAN', '')
    return s

def clean_num(series):
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def clean_str(series):
    return series.astype(str).str.replace('nan', '', case=False).str.strip()

def read_file(file):
    if file is None:
        return pd.DataFrame()
    if file.name.endswith(('.xlsx', '.xls', '.xlsm')):
        try:
            file.seek(0)
            return pd.read_excel(file, dtype=str, engine='openpyxl')
        except:
            return pd.DataFrame()
    encodings = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'gbk', 'latin1']
    for enc in encodings:
        try:
            file.seek(0)
            return pd.read_csv(file, dtype=str, encoding=enc)
        except:
            continue
    return pd.DataFrame()

def blank_repeat_like_merge(df: pd.DataFrame, group_col: str, cols_to_blank: list):
    """让指定列在同一group内重复行置空，达到“视觉合并”效果（兼容Excel Table）"""
    out = df.copy()
    if group_col not in out.columns:
        return out
    first = out[group_col].astype(str).fillna('').ne(out[group_col].astype(str).shift())
    for c in cols_to_blank:
        if c in out.columns:
            out.loc[~first, c] = ''
    return out

def estimate_col_widths(df: pd.DataFrame, fixed_col_names=None, fixed_width=26,
                        min_w=6, max_w=22, header_pad=2, cell_pad=2):
    """xlsxwriter无真正autofit，用字符长度估算列宽"""
    fixed_col_names = set(fixed_col_names or [])
    widths = []
    for col in df.columns:
        if col in fixed_col_names:
            widths.append(fixed_width)
            continue
        best = len(str(col)) + header_pad
        ser = df[col].fillna('')
        lens = ser.astype(str).map(len)
        if len(lens) > 0:
            best = max(best, int(lens.max()) + cell_pad)
        best = max(min_w, min(max_w, best))
        widths.append(best)
    return widths

def col_to_excel(col_idx: int) -> str:
    """0->A, 1->B ..."""
    n = col_idx + 1
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def make_work_order_html(df: pd.DataFrame, title: str, subtitle: str) -> bytes:
    """A4灰白可打印HTML，在线加载Noto字体；斑马纹按产品编码分组"""
    df2 = df.copy()
    for c in df2.columns:
        df2[c] = df2[c].fillna('').astype(str)

    if '产品编码' in df2.columns and len(df2) > 0:
        codes = df2['产品编码'].astype(str).fillna('')
        gid = (codes.ne(codes.shift()).cumsum() % 2).astype(int)
    else:
        gid = pd.Series([0] * len(df2), dtype=int)

    left_cols = set([c for c in df2.columns if c in ('基础信息', 'SKU名称')])

    thead = "<thead><tr>" + "".join([f"<th>{col}</th>" for col in df2.columns]) + "</tr></thead>"
    rows = []
    for i in range(len(df2)):
        zebra = "z1" if int(gid.iloc[i]) == 1 else "z0"
        tds = []
        for col in df2.columns:
            cls = "td-left" if col in left_cols else "td-center"
            tds.append(f'<td class="{cls}">{df2.iloc[i][col]}</td>')
        rows.append(f'<tr class="{zebra}">' + "".join(tds) + "</tr>")
    tbody = "<tbody>" + "".join(rows) + "</tbody>"

    today = datetime.now().strftime("%Y-%m-%d")

    html = f"""<!doctype html>
<html lang="zh">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;600&family=Noto+Sans+KR:wght@400;600&display=swap" rel="stylesheet">
  <title>{title}</title>
  <style>
    :root {{
      --fg:#111; --muted:#555; --line:#d8d8d8; --head:#3f3f3f; --zebra:#f1f1f1; --white:#fff;
    }}
    html,body {{
      margin:0; padding:0; background:var(--white); color:var(--fg);
      font-family:"Noto Sans SC","Noto Sans KR",system-ui,-apple-system,"Segoe UI",Arial,sans-serif;
      -webkit-print-color-adjust:exact; print-color-adjust:exact;
    }}
    .page {{ width:210mm; margin:0 auto; padding:14mm 12mm; box-sizing:border-box; }}
    .header {{
      display:flex; justify-content:space-between; align-items:flex-end;
      border-bottom:2px solid var(--line); padding-bottom:10px; margin-bottom:12px; gap:12px;
    }}
    .title {{ font-size:18px; font-weight:600; margin:0; }}
    .sub {{ margin:6px 0 0 0; font-size:12px; color:var(--muted); }}
    .meta {{ font-size:12px; color:var(--muted); text-align:right; white-space:nowrap; }}
    table {{ width:100%; border-collapse:collapse; font-size:11px; }}
    th {{
      background:var(--head); color:#fff; padding:8px 6px; border:1px solid var(--line);
      text-align:center; font-weight:600;
    }}
    td {{ border:1px solid var(--line); padding:6px 6px; vertical-align:middle; word-break:break-word; }}
    .td-center {{ text-align:center; }}
    .td-left {{ text-align:left; }}
    tr.z1 td {{ background:var(--zebra); }}
    @media print {{
      thead {{ display:table-header-group; }}
      tr {{ page-break-inside:avoid; }}
      .page {{ padding:10mm 10mm; }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <div class="header">
      <div>
        <h1 class="title">{title}</h1>
        <div class="sub">{subtitle}</div>
      </div>
      <div class="meta">
        生成日期：{today}<br/>
        打印：A4 / 黑白灰
      </div>
    </div>
    <table>{thead}{tbody}</table>
  </div>
</body>
</html>
"""
    return html.encode("utf-8")

# ==========================================
# 4. 侧边栏
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数设置")

    st.subheader("🛡️ 总安全库存 (采购)")
    safety_weeks = st.number_input("安全周数 (倍数)", min_value=1, max_value=20, value=3, step=1)
    min_safety_qty = st.number_input(
        "最低库存基数 (保底)",
        min_value=0,
        max_value=100,
        value=5,
        step=1,
        help="仅对【在做】且【有入库码】的产品生效"
    )

    st.divider()
    orange_safety_weeks = st.number_input("🚚 橙火安全周数 (调拨预警)", min_value=1, max_value=10, value=2, step=1)

    st.divider()
    redundancy_weeks = st.number_input("⚠️ 库存冗余周数 (滞销标准)", min_value=4, max_value=52, value=8, step=1)

    st.divider()
    st.subheader("🔍 单品库存查询")
    search_key = st.text_input("输入产品编码 (A列)", placeholder="输入后按回车查询，留空看全部")

    st.divider()
    st.info("📂 请上传文件 (保持Master顺序)")
    file_master = st.file_uploader("1. 基础信息表 (Master) *必传", type=['xlsx', 'csv'])
    files_sales = st.file_uploader("2. 销售表 (近7天) *多选", type=['xlsx', 'csv'], accept_multiple_files=True)
    files_inv_r = st.file_uploader("3. 橙火/火箭仓库存 *多选", type=['xlsx', 'csv'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("4. 极风库存 *多选", type=['xlsx', 'csv'], accept_multiple_files=True)

# ==========================================
# 5. 主逻辑
# ==========================================
if file_master and files_sales and files_inv_r and files_inv_j:
    if st.button("🚀 生成定制报表", type="primary", use_container_width=True):
        with st.spinner("正在按指定列顺序匹配数据..."):

            # --- A. 读取 Master ---
            df_m = read_file(file_master)
            if df_m.empty:
                st.stop()

            df_base = pd.DataFrame()
            try:
                df_base['Shop'] = clean_str(df_m.iloc[:, IDX_M_SHOP])
                df_base['Code'] = clean_match_key(df_m.iloc[:, IDX_M_CODE])
                df_base['Info_E'] = clean_str(df_m.iloc[:, IDX_M_COL_E])
                df_base['Info_F'] = clean_str(df_m.iloc[:, IDX_M_COL_F])
                df_base['Cost'] = clean_num(df_m.iloc[:, IDX_M_COST])
                df_base['Orange_ID'] = clean_match_key(df_m.iloc[:, IDX_M_ORANGE])
                df_base['Inbound_Code'] = clean_match_key(df_m.iloc[:, IDX_M_INBOUND])

                active_raw = clean_str(df_m.iloc[:, IDX_M_ACTIVE])
                df_base['Active'] = active_raw.astype(str).str.contains('Y', case=False, na=False).map(lambda x: 'Y' if x else '')
            except IndexError:
                st.error("❌ 基础表列数不足，请检查列配置！")
                st.stop()

            # --- B. 销售汇总 ---
            s_list = [read_file(f) for f in files_sales]
            if not s_list:
                st.stop()
            df_sales = pd.concat(s_list, ignore_index=True)
            df_sales['Key'] = clean_match_key(df_sales.iloc[:, IDX_7D_SKU])
            df_sales['Qty'] = clean_num(df_sales.iloc[:, IDX_7D_QTY])
            agg_sales = df_sales.groupby('Key')['Qty'].sum().reset_index()

            # --- C. 橙火库存 ---
            r_list = [read_file(f) for f in files_inv_r]
            if r_list:
                df_r = pd.concat(r_list, ignore_index=True)
                df_r['Key'] = clean_match_key(df_r.iloc[:, IDX_INV_R_SKU])
                df_r['Qty'] = clean_num(df_r.iloc[:, IDX_INV_R_QTY])
                try:
                    df_r['Fee'] = clean_num(df_r.iloc[:, IDX_INV_R_FEE])
                except:
                    df_r['Fee'] = 0
                agg_orange = df_r.groupby('Key')[['Qty', 'Fee']].sum().reset_index()
            else:
                agg_orange = pd.DataFrame(columns=['Key', 'Qty', 'Fee'])

            # --- D. 极风库存 ---
            j_list = [read_file(f) for f in files_inv_j]
            if j_list:
                df_j = pd.concat(j_list, ignore_index=True)
                df_j['Key'] = clean_match_key(df_j.iloc[:, IDX_INV_J_BAR])
                df_j['Qty'] = clean_num(df_j.iloc[:, IDX_INV_J_QTY])
                agg_jifeng = df_j.groupby('Key')['Qty'].sum().reset_index()
            else:
                agg_jifeng = pd.DataFrame(columns=['Key', 'Qty'])

            # --- E. 匹配合并 ---
            df_final = pd.merge(df_base, agg_sales, left_on='Orange_ID', right_on='Key', how='left')
            df_final.rename(columns={'Qty': 'Sales_7d'}, inplace=True)

            df_final = pd.merge(df_final, agg_orange, left_on='Orange_ID', right_on='Key', how='left', suffixes=('', '_R'))
            df_final.rename(columns={'Qty': 'Stock_Orange', 'Fee': 'Storage_Fee'}, inplace=True)

            df_final = pd.merge(df_final, agg_jifeng, left_on='Inbound_Code', right_on='Key', how='left', suffixes=('', '_J'))
            df_final.rename(columns={'Qty': 'Stock_Jifeng'}, inplace=True)

            # --- F. 计算逻辑 ---
            df_final['Sales_7d'] = df_final['Sales_7d'].fillna(0)
            df_final['Stock_Orange'] = df_final['Stock_Orange'].fillna(0)
            df_final['Stock_Jifeng'] = df_final['Stock_Jifeng'].fillna(0)
            df_final['Storage_Fee'] = df_final['Storage_Fee'].fillna(0)

            df_final['Total_Stock'] = df_final['Stock_Orange'] + df_final['Stock_Jifeng']
            df_final['Safety_Calc'] = df_final['Sales_7d'] * safety_weeks

            def apply_safety_floor(row):
                base_val = row['Safety_Calc']
                is_active = (str(row.get('Active', '')).strip().upper() == 'Y')
                has_inbound = bool(str(row.get('Inbound_Code', '')).strip())
                if is_active and has_inbound:
                    return max(base_val, min_safety_qty)
                return base_val

            df_final['Safety'] = df_final.apply(apply_safety_floor, axis=1)
            df_final['Redundancy_Std'] = df_final['Sales_7d'] * redundancy_weeks

            df_final['Restock_Qty'] = (df_final['Safety'] - df_final['Total_Stock']).apply(lambda x: int(x) if x > 0 else 0)
            inactive_mask = df_final['Active'].astype(str).str.upper().ne('Y')
            df_final.loc[inactive_mask, 'Restock_Qty'] = 0
            df_final['Restock_Money'] = df_final['Restock_Qty'] * df_final['Cost']

            df_final['Redundancy_Qty'] = (df_final['Total_Stock'] - df_final['Redundancy_Std']).apply(lambda x: int(x) if x > 0 else 0)
            df_final['Redundancy_Money'] = df_final['Redundancy_Qty'] * df_final['Cost']

            df_final['Orange_Safety_Calc'] = df_final['Sales_7d'] * orange_safety_weeks

            def apply_orange_floor(row):
                base_val = row['Orange_Safety_Calc']
                has_inbound = bool(str(row.get('Inbound_Code', '')).strip())
                if has_inbound:
                    return max(base_val, min_safety_qty)
                return base_val

            df_final['Orange_Safety_Std'] = df_final.apply(apply_orange_floor, axis=1)
            df_final['Orange_Transfer_Qty'] = (df_final['Orange_Safety_Std'] - df_final['Stock_Orange']).apply(lambda x: int(x) if x > 0 else 0)

            # --- G. 整理输出 ---
            cols_export_base = [
                'Shop', 'Code', 'Info_E', 'Info_F', 'Cost', 'Orange_ID', 'Inbound_Code',
                'Sales_7d', 'Stock_Orange', 'Stock_Jifeng', 'Total_Stock', 'Safety',
                'Restock_Qty', 'Restock_Money', 'Redundancy_Std', 'Redundancy_Qty',
                'Redundancy_Money', 'Orange_Safety_Std', 'Orange_Transfer_Qty', 'Storage_Fee'
            ]
            df_out_base = df_final[cols_export_base].copy()

            header_map_base = {
                'Shop': '店铺名称',
                'Code': '产品编码',
                'Info_E': '基础信息',
                'Info_F': 'SKU名称',
                'Cost': '采购单价',
                'Orange_ID': '橙火ID',
                'Inbound_Code': '入库码',
                'Sales_7d': '7天销量',
                'Stock_Orange': '橙火库存',
                'Stock_Jifeng': '极风库存',
                'Total_Stock': '库存合计',
                'Safety': f'总安全库存(有码>{min_safety_qty})',
                'Restock_Qty': '建议采购数',
                'Restock_Money': '预计采购总额(RMB)',
                'Redundancy_Std': f'冗余标准({redundancy_weeks}周)',
                'Redundancy_Qty': '冗余数量',
                'Redundancy_Money': '冗余资金',
                'Orange_Safety_Std': f'橙火安全库存(有码>{min_safety_qty})',
                'Orange_Transfer_Qty': '建议调拨数量',
                'Storage_Fee': '本月仓储费(预警)'
            }
            df_out_base.rename(columns=header_map_base, inplace=True)

            # Sheet1/展示专用：插入“在做(Y)”列
            df_sheet1 = df_out_base.copy()
            df_sheet1.insert(1, '在做(Y)', df_final['Active'].values)

            # ==========================================
            # ✅ H. 搜索与KPI + 高亮看板（按产品编码分组斑马纹）
            # ==========================================
            if search_key:
                df_display = df_sheet1[df_sheet1['产品编码'].astype(str).str.contains(search_key, case=False, na=False)].copy()
            else:
                df_display = df_sheet1.copy()

            # KPI（按当前筛选后的df_display）
            buy_mask = df_display.get('建议采购数', pd.Series([0] * len(df_display))).astype(float) > 0
            k1_cnt = int(buy_mask.sum())
            k1_val = float(df_display.loc[buy_mask, '预计采购总额(RMB)'].sum()) if '预计采购总额(RMB)' in df_display.columns else 0.0

            red_mask = df_display.get('冗余数量', pd.Series([0] * len(df_display))).astype(float) > 0
            k2_cnt = int(red_mask.sum())
            k2_val = float(df_display.loc[red_mask, '冗余资金'].sum()) if '冗余资金' in df_display.columns else 0.0

            trans_mask = df_display.get('建议调拨数量', pd.Series([0] * len(df_display))).astype(float) > 0
            k3_cnt = int(trans_mask.sum())
            k3_val = float(df_display.loc[trans_mask, '建议调拨数量'].sum()) if '建议调拨数量' in df_display.columns else 0.0

            fee_mask = df_display.get('本月仓储费(预警)', pd.Series([0] * len(df_display))).astype(float) > 0
            k4_cnt = int(fee_mask.sum())
            k4_val = float(df_display.loc[fee_mask, '本月仓储费(预警)'].sum()) if '本月仓储费(预警)' in df_display.columns else 0.0

            st.divider()
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("**📦 需采购 SKU / 金额**", f"{k1_cnt} 个", f"¥ {k1_val:,.0f}")
            m2.metric("**⚠️ 冗余 SKU / 资金**", f"{k2_cnt} 个", f"¥ {k2_val:,.0f}", delta_color="inverse")
            m3.metric("**🚚 需调拨 SKU / 数量**", f"{k3_cnt} 个", f"{k3_val:,.0f} 件")
            m4.metric("**🚨 库龄预警 SKU / 总仓储费**", f"{k4_cnt} 个", f"₩ {k4_val:,.0f}", delta_color="inverse")

            # 视觉置空（店铺名称/产品编码同一产品只显示一次）
            df_display_vis = df_display.copy()
            if '产品编码' in df_display_vis.columns:
                first_in_group = df_display_vis['产品编码'].astype(str).fillna('').ne(df_display_vis['产品编码'].astype(str).shift())
                for col in ['店铺名称', '产品编码']:
                    if col in df_display_vis.columns:
                        df_display_vis.loc[~first_in_group, col] = ''

                # 按产品编码分组斑马纹
                zebra_group_ids = (df_display['产品编码'].astype(str).fillna('').ne(df_display['产品编码'].astype(str).shift()).cumsum() % 2)
            else:
                zebra_group_ids = pd.Series([0] * len(df_display_vis), index=df_display_vis.index)

            def highlight_zebra(row):
                try:
                    gid = zebra_group_ids.loc[row.name]
                    if int(gid) == 1:
                        return ['background-color: #f7f7f7'] * len(row)
                except:
                    pass
                return [''] * len(row)

            def highlight_bold_cols(row):
                # 用于“产品编码、SKU名称”整列加粗
                return ['font-weight: bold'] * len(row)

            def highlight_restock_qty(s):
                return ['background-color: #ffcccc; color: #b71c1c; font-weight: bold' if float(v) > 0 else '' for v in s]

            def highlight_restock_money(s):
                return ['background-color: #ffcccc; color: #b71c1c' if float(v) > 0 else '' for v in s]

            def highlight_redundancy_qty(s):
                return ['background-color: #ffe0b2; color: #e65100; font-weight: bold' if float(v) > 0 else '' for v in s]

            def highlight_redundancy_money(s):
                return ['background-color: #ffe0b2; color: #e65100' if float(v) > 0 else '' for v in s]

            def highlight_transfer(s):
                return ['background-color: #e3f2fd; color: #0d47a1; font-weight: bold' if float(v) > 0 else '' for v in s]

            def highlight_fee(s):
                return ['background-color: #e1bee7; color: #4a148c; font-weight: bold' if float(v) > 0 else '' for v in s]

            st_df = df_display_vis.style.apply(highlight_zebra, axis=1)

            # 加粗列（存在才加）
            for c in ['产品编码', 'SKU名称']:
                if c in df_display_vis.columns:
                    st_df = st_df.apply(highlight_bold_cols, subset=[c])

            # 条件高亮列（存在才加）
            if '建议采购数' in df_display_vis.columns:
                st_df = st_df.apply(highlight_restock_qty, subset=['建议采购数'])
            if '预计采购总额(RMB)' in df_display_vis.columns:
                st_df = st_df.apply(highlight_restock_money, subset=['预计采购总额(RMB)'])
            if '冗余数量' in df_display_vis.columns:
                st_df = st_df.apply(highlight_redundancy_qty, subset=['冗余数量'])
            if '冗余资金' in df_display_vis.columns:
                st_df = st_df.apply(highlight_redundancy_money, subset=['冗余资金'])
            if '建议调拨数量' in df_display_vis.columns:
                st_df = st_df.apply(highlight_transfer, subset=['建议调拨数量'])
            if '本月仓储费(预警)' in df_display_vis.columns:
                st_df = st_df.apply(highlight_fee, subset=['本月仓储费(预警)'])

            # 格式化（存在才format）
            fmt_map = {}
            for k in ['橙火库存', '极风库存', '库存合计', '7天销量', '建议采购数', '冗余数量', '建议调拨数量']:
                if k in df_display_vis.columns:
                    fmt_map[k] = '{:.0f}'
            if f'总安全库存(有码>{min_safety_qty})' in df_display_vis.columns:
                fmt_map[f'总安全库存(有码>{min_safety_qty})'] = '{:.0f}'
            if f'冗余标准({redundancy_weeks}周)' in df_display_vis.columns:
                fmt_map[f'冗余标准({redundancy_weeks}周)'] = '{:.0f}'
            if f'橙火安全库存(有码>{min_safety_qty})' in df_display_vis.columns:
                fmt_map[f'橙火安全库存(有码>{min_safety_qty})'] = '{:.0f}'
            if '预计采购总额(RMB)' in df_display_vis.columns:
                fmt_map['预计采购总额(RMB)'] = '{:,.0f}'
            if '采购单价' in df_display_vis.columns:
                fmt_map['采购单价'] = '{:,.0f}'
            if '冗余资金' in df_display_vis.columns:
                fmt_map['冗余资金'] = '{:,.0f}'
            if '本月仓储费(预警)' in df_display_vis.columns:
                fmt_map['本月仓储费(预警)'] = '{:,.0f}'

            if fmt_map:
                st_df = st_df.format(fmt_map)

            st.dataframe(st_df, use_container_width=True, height=600, hide_index=True)

            # ==========================================
            # Excel 导出（Table + 按产品编码分组斑马纹 + 两列左对齐）
            # ==========================================
            out_io = io.BytesIO()
            with pd.ExcelWriter(out_io, engine='xlsxwriter') as writer:
                wb = writer.book

                fmt_center = wb.add_format({'align': 'center', 'valign': 'vcenter'})
                fmt_left = wb.add_format({'align': 'left', 'valign': 'vcenter'})

                fmt_zebra_group = wb.add_format({'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter'})
                fmt_zebra_left = wb.add_format({'bg_color': '#F2F2F2', 'align': 'left', 'valign': 'vcenter'})

                fmt_red_bold = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True, 'align': 'center', 'valign': 'vcenter'})
                fmt_red_norm = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': False, 'align': 'center', 'valign': 'vcenter'})
                fmt_orange_bold = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': True, 'align': 'center', 'valign': 'vcenter'})
                fmt_orange_norm = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': False, 'align': 'center', 'valign': 'vcenter'})
                fmt_blue = wb.add_format({'bg_color': '#C5D9F1', 'font_color': '#1F497D', 'bold': True, 'align': 'center', 'valign': 'vcenter'})
                fmt_purple = wb.add_format({'bg_color': '#E1BEE7', 'font_color': '#4A148C', 'bold': True, 'align': 'center', 'valign': 'vcenter'})
                fmt_bold = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})

                def apply_conditional(ws, df_curr):
                    col_idx = {c: i for i, c in enumerate(df_curr.columns)}
                    nrows = len(df_curr)
                    if nrows <= 0:
                        return

                    def cf(colname, rule):
                        if colname not in col_idx:
                            return
                        c = col_idx[colname]
                        if rule == 'bold':
                            ws.conditional_format(1, c, nrows, c, {'type': 'formula', 'criteria': '=TRUE', 'format': fmt_bold})
                        elif rule == 'red_bold':
                            ws.conditional_format(1, c, nrows, c, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red_bold})
                        elif rule == 'red_norm':
                            ws.conditional_format(1, c, nrows, c, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red_norm})
                        elif rule == 'orange_bold':
                            ws.conditional_format(1, c, nrows, c, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange_bold})
                        elif rule == 'orange_norm':
                            ws.conditional_format(1, c, nrows, c, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange_norm})
                        elif rule == 'blue':
                            ws.conditional_format(1, c, nrows, c, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_blue})
                        elif rule == 'purple':
                            ws.conditional_format(1, c, nrows, c, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_purple})

                    cf('产品编码', 'bold')
                    cf('SKU名称', 'bold')
                    cf('建议采购数', 'red_bold')
                    cf('预计采购总额(RMB)', 'red_norm')
                    cf('冗余数量', 'orange_bold')
                    cf('冗余资金', 'orange_norm')
                    cf('建议调拨数量', 'blue')
                    cf('本月仓储费(预警)', 'purple')

                def build_table_sheet(sheet_name, df_curr, fixed_width_cols=None, fixed_width=26, hide_cols=None):
                    hide_cols = hide_cols or []
                    fixed_width_cols = fixed_width_cols or []

                    if '产品编码' in df_curr.columns and len(df_curr) > 0:
                        codes = df_curr['产品编码'].astype(str).fillna('')
                        gid = (codes.ne(codes.shift()).cumsum() % 2).astype(int)
                    else:
                        gid = pd.Series([0] * len(df_curr), dtype=int)

                    df_write = df_curr.copy()
                    if '产品编码' in df_write.columns and '店铺名称' in df_write.columns:
                        df_write = blank_repeat_like_merge(df_write, '产品编码', ['店铺名称', '产品编码'])

                    df_write.to_excel(writer, index=False, sheet_name=sheet_name)
                    ws = writer.sheets[sheet_name]

                    nrows, ncols = len(df_write), len(df_write.columns)

                    columns = [{'header': c} for c in df_write.columns]
                    if ncols > 0:
                        last_row = nrows if nrows > 0 else 0
                        ws.add_table(0, 0, last_row, ncols - 1, {
                            'columns': columns,
                            'style': 'Table Style Medium 9',
                            'autofilter': True if nrows > 0 else False,
                            'banded_rows': False
                        })

                    widths = estimate_col_widths(df_write, fixed_col_names=fixed_width_cols, fixed_width=fixed_width)
                    for i, w in enumerate(widths):
                        ws.set_column(i, i, w, fmt_center)

                    for colname in ['基础信息', 'SKU名称']:
                        if colname in df_write.columns:
                            cidx = list(df_write.columns).index(colname)
                            ws.set_column(cidx, cidx, widths[cidx], fmt_left)

                    helper_col = ncols
                    ws.write(0, helper_col, "_gid")
                    for i in range(nrows):
                        ws.write(i + 1, helper_col, int(gid.iloc[i]), fmt_center)
                    ws.set_column(helper_col, helper_col, None, None, {'hidden': True})

                    if nrows > 0 and ncols > 0:
                        helper_letter = col_to_excel(helper_col)
                        formula = f'=${helper_letter}2=1'
                        ws.conditional_format(1, 0, nrows, ncols - 1, {
                            'type': 'formula',
                            'criteria': formula,
                            'format': fmt_zebra_group
                        })
                        for colname in ['基础信息', 'SKU名称']:
                            if colname in df_write.columns:
                                cidx = list(df_write.columns).index(colname)
                                ws.conditional_format(1, cidx, nrows, cidx, {
                                    'type': 'formula',
                                    'criteria': formula,
                                    'format': fmt_zebra_left
                                })

                    for c in hide_cols:
                        if 0 <= c < ncols:
                            ws.set_column(c, c, None, None, {'hidden': True})

                    apply_conditional(ws, df_write)

                build_table_sheet('补货计算表', df_sheet1, fixed_width_cols=['基础信息'], fixed_width=26)

                df_buy = df_out_base[df_out_base['建议采购数'] > 0].copy()
                build_table_sheet('采购单(找工厂)', df_buy, fixed_width_cols=['基础信息'], fixed_width=26,
                                  hide_cols=[14, 15, 16, 17, 18, 19])

                df_trans = df_out_base[df_out_base['建议调拨数量'] > 0].copy()
                build_table_sheet('调拨单(发橙火)', df_trans, fixed_width_cols=['基础信息'], fixed_width=26,
                                  hide_cols=[12, 13, 14, 15, 16, 17, 19])

                df_fee = df_out_base[df_out_base['本月仓储费(预警)'] > 0].copy()
                build_table_sheet('库龄预警单(需重入库)', df_fee, fixed_width_cols=['基础信息'], fixed_width=26,
                                  hide_cols=[11, 12, 13, 14, 15, 16, 17, 18])

            excel_bytes = out_io.getvalue()

            # ZIP打包：Excel + 3个HTML工单
            stamp = pd.Timestamp.now().strftime('%Y%m%d')
            excel_name = f"Coupang_Restock_Full_v18_{stamp}.xlsx"

            html_buy = make_work_order_html(df_buy, "采购工单（找工厂）", "范围：建议采购数 > 0")
            html_trans = make_work_order_html(df_trans, "调拨工单（发橙火）", "范围：建议调拨数量 > 0")
            html_fee = make_work_order_html(df_fee, "库龄预警工单（需重入库）", "范围：本月仓储费(预警) > 0")

            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', compression=zipfile.ZIP_DEFLATED) as z:
                z.writestr(excel_name, excel_bytes)
                z.writestr(f"WorkOrder_Buy_{stamp}.html", html_buy)
                z.writestr(f"WorkOrder_Transfer_{stamp}.html", html_trans)
                z.writestr(f"WorkOrder_Fee_{stamp}.html", html_fee)

            st.download_button(
                "📦 下载压缩包（Excel + 3个工单HTML）",
                data=zip_buf.getvalue(),
                file_name=f"Coupang_Restock_Pack_{stamp}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
else:
    st.info("👈 请在左侧上传文件")
