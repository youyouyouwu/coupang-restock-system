import streamlit as st
import pandas as pd
import io

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
        help="仅对有入库码的产品生效"
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

                # ✅ 新增：N列是否在做（Y大小写容错；包含Y即可判定）
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

            # 安全库存 (保底逻辑)
            df_final['Safety_Calc'] = df_final['Sales_7d'] * safety_weeks

            def apply_safety_floor(row):
                base_val = row['Safety_Calc']
                if row['Inbound_Code']:
                    return max(base_val, min_safety_qty)
                else:
                    return base_val

            df_final['Safety'] = df_final.apply(apply_safety_floor, axis=1)

            df_final['Redundancy_Std'] = df_final['Sales_7d'] * redundancy_weeks

            df_final['Restock_Qty'] = (df_final['Safety'] - df_final['Total_Stock']).apply(lambda x: int(x) if x > 0 else 0)
            df_final['Restock_Money'] = df_final['Restock_Qty'] * df_final['Cost']

            df_final['Redundancy_Qty'] = (df_final['Total_Stock'] - df_final['Redundancy_Std']).apply(lambda x: int(x) if x > 0 else 0)
            df_final['Redundancy_Money'] = df_final['Redundancy_Qty'] * df_final['Cost']

            df_final['Orange_Safety_Calc'] = df_final['Sales_7d'] * orange_safety_weeks

            def apply_orange_floor(row):
                base_val = row['Orange_Safety_Calc']
                if row['Inbound_Code']:
                    return max(base_val, min_safety_qty)
                else:
                    return base_val

            df_final['Orange_Safety_Std'] = df_final.apply(apply_orange_floor, axis=1)

            df_final['Orange_Transfer_Qty'] = (df_final['Orange_Safety_Std'] - df_final['Stock_Orange']).apply(
                lambda x: int(x) if x > 0 else 0
            )

            # --- G. 整理输出 ---
            # ✅ 基础导出（不加“在做”列）：用于 sheet2/3/4，保证其它任何地方不变
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

            # ✅ 仅用于：Streamlit显示 + Sheet1(补货计算表) 的表（插入“在做(Y)”列在店铺名称与产品编码之间）
            df_sheet1 = df_out_base.copy()
            # 从 df_final 拿到 Active，并插入到第2列（索引1位置）
            active_col = df_final['Active'].copy()
            df_sheet1.insert(1, '在做(Y)', active_col.values)

            # --- H. 搜索与看板（仅基于 df_sheet1 显示） ---
            if search_key:
                df_display = df_sheet1[df_sheet1['产品编码'].astype(str).str.contains(search_key, case=False, na=False)]
            else:
                df_display = df_sheet1

            zebra_group_ids = (df_display['产品编码'] != df_display['产品编码'].shift()).cumsum() % 2

            st.divider()

            buy_mask = df_display['建议采购数'] > 0
            k1_cnt = len(df_display[buy_mask])
            k1_val = df_display.loc[buy_mask, '预计采购总额(RMB)'].sum()

            red_mask = df_display['冗余数量'] > 0
            k2_cnt = len(df_display[red_mask])
            k2_val = df_display.loc[red_mask, '冗余资金'].sum()

            trans_mask = df_display['建议调拨数量'] > 0
            k3_cnt = len(df_display[trans_mask])
            k3_val = df_display.loc[trans_mask, '建议调拨数量'].sum()

            fee_mask = df_display['本月仓储费(预警)'] > 0
            k4_cnt = len(df_display[fee_mask])
            k4_val = df_display.loc[fee_mask, '本月仓储费(预警)'].sum()

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("**📦 需采购 SKU / 金额**", f"{k1_cnt} 个", f"¥ {k1_val:,.0f}")
            m2.metric("**⚠️ 冗余 SKU / 资金**", f"{k2_cnt} 个", f"¥ {k2_val:,.0f}", delta_color="inverse")
            m3.metric("**🚚 需调拨 SKU / 数量**", f"{k3_cnt} 个", f"{k3_val:,.0f} 件")
            m4.metric("**🚨 库龄预警 SKU / 总仓储费**", f"{k4_cnt} 个", f"₩ {k4_val:,.0f}", delta_color="inverse")

            # === 样式定义 ===
            def highlight_zebra(row):
                try:
                    gid = zebra_group_ids.loc[row.name]
                    if gid == 1:
                        return ['background-color: #f7f7f7'] * len(row)
                except:
                    pass
                return [''] * len(row)

            def highlight_bold_info(s):
                return ['font-weight: bold'] * len(s)

            def highlight_restock_qty(s):
                return ['background-color: #ffcccc; color: #b71c1c; font-weight: bold' if v > 0 else '' for v in s]

            def highlight_restock_money(s):
                return ['background-color: #ffcccc; color: #b71c1c' if v > 0 else '' for v in s]

            def highlight_redundancy_qty(s):
                return ['background-color: #ffe0b2; color: #e65100; font-weight: bold' if v > 0 else '' for v in s]

            def highlight_redundancy_money(s):
                return ['background-color: #ffe0b2; color: #e65100' if v > 0 else '' for v in s]

            def highlight_transfer(s):
                return ['background-color: #e3f2fd; color: #0d47a1; font-weight: bold' if v > 0 else '' for v in s]

            def highlight_fee(s):
                return ['background-color: #e1bee7; color: #4a148c; font-weight: bold' if v > 0 else '' for v in s]

            st_df = (
                df_display.style
                .apply(highlight_zebra, axis=1)
                .apply(highlight_bold_info, subset=['产品编码', 'SKU名称'])
                .apply(highlight_restock_qty, subset=['建议采购数'])
                .apply(highlight_restock_money, subset=['预计采购总额(RMB)'])
                .apply(highlight_redundancy_qty, subset=['冗余数量'])
                .apply(highlight_redundancy_money, subset=['冗余资金'])
                .apply(highlight_transfer, subset=['建议调拨数量'])
                .apply(highlight_fee, subset=['本月仓储费(预警)'])
                .format({
                    '橙火库存': '{:.0f}',
                    '极风库存': '{:.0f}',
                    '库存合计': '{:.0f}',
                    f'总安全库存(有码>{min_safety_qty})': '{:.0f}',
                    f'冗余标准({redundancy_weeks}周)': '{:.0f}',
                    f'橙火安全库存(有码>{min_safety_qty})': '{:.0f}',
                    '建议采购数': '{:.0f}',
                    '预计采购总额(RMB)': '{:,.0f}',
                    '7天销量': '{:.0f}',
                    '采购单价': '{:,.0f}',
                    '冗余数量': '{:.0f}',
                    '冗余资金': '{:,.0f}',
                    '建议调拨数量': '{:.0f}',
                    '本月仓储费(预警)': '{:,.0f}',
                })
            )

            st.dataframe(st_df, use_container_width=True, height=600, hide_index=True)

            # Excel 导出
            out_io = io.BytesIO()
            with pd.ExcelWriter(out_io, engine='xlsxwriter') as writer:

                wb = writer.book

                fmt_header = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                fmt_header_dark = wb.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': 'white', 'border': 1})

                fmt_zebra = wb.add_format({'bg_color': '#F2F2F2'})
                fmt_bold_col = wb.add_format({'bold': True})

                fmt_red_bold = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True})
                fmt_red_norm = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': False})
                fmt_orange_bold = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': True})
                fmt_orange_norm = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': False})
                fmt_blue = wb.add_format({'bg_color': '#C5D9F1', 'font_color': '#1F497D', 'bold': True})
                fmt_purple = wb.add_format({'bg_color': '#E1BEE7', 'font_color': '#4A148C', 'bold': True})

                # --- 封装格式化函数（给原结构：20列）---
                def format_sheet(ws, df_curr, hide_cols=[]):
                    # 1. 斑马纹
                    curr_zebra_ids = (df_curr['产品编码'] != df_curr['产品编码'].shift()).cumsum() % 2
                    for i, gid in enumerate(curr_zebra_ids):
                        if gid == 1:
                            ws.set_row(i + 1, None, fmt_zebra)

                    # 2. 表头
                    ws.set_row(0, None, fmt_header)
                    target_headers = {
                        1: '产品编码',
                        3: 'SKU名称',
                        12: '建议采购数',
                        15: '冗余数量',
                        18: '建议调拨数量',
                        19: '本月仓储费(预警)'
                    }
                    for col_idx, text in target_headers.items():
                        ws.write(0, col_idx, text, fmt_header_dark)
                    ws.set_column('A:T', 13)

                    # 3. 隐藏列
                    for c_idx in hide_cols:
                        ws.set_column(c_idx, c_idx, None, None, {'hidden': True})

                    # 4. 条件格式
                    nrows = len(df_curr)
                    ws.conditional_format(1, 1, nrows, 1, {'type': 'formula', 'criteria': '=TRUE', 'format': fmt_bold_col})
                    ws.conditional_format(1, 3, nrows, 3, {'type': 'formula', 'criteria': '=TRUE', 'format': fmt_bold_col})
                    ws.conditional_format(1, 12, nrows, 12, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red_bold})
                    ws.conditional_format(1, 13, nrows, 13, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red_norm})
                    ws.conditional_format(1, 15, nrows, 15, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange_bold})
                    ws.conditional_format(1, 16, nrows, 16, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange_norm})
                    ws.conditional_format(1, 18, nrows, 18, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_blue})
                    ws.conditional_format(1, 19, nrows, 19, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_purple})

                # --- ✅ Sheet1 专用格式化（新增1列：在做(Y)，总21列 A:U）---
                def format_sheet_with_active(ws, df_curr):
                    # 1. 斑马纹（仍按产品编码分组）
                    curr_zebra_ids = (df_curr['产品编码'] != df_curr['产品编码'].shift()).cumsum() % 2
                    for i, gid in enumerate(curr_zebra_ids):
                        if gid == 1:
                            ws.set_row(i + 1, None, fmt_zebra)

                    # 2. 表头
                    ws.set_row(0, None, fmt_header)
                    target_headers = {
                        2: '产品编码',            # 原1 -> 2
                        4: 'SKU名称',            # 原3 -> 4
                        13: '建议采购数',         # 原12 -> 13
                        16: '冗余数量',           # 原15 -> 16
                        19: '建议调拨数量',       # 原18 -> 19
                        20: '本月仓储费(预警)'    # 原19 -> 20
                    }
                    for col_idx, text in target_headers.items():
                        ws.write(0, col_idx, text, fmt_header_dark)
                    ws.set_column('A:U', 13)

                    # 3. 条件格式（索引整体+1）
                    nrows = len(df_curr)
                    ws.conditional_format(1, 2, nrows, 2, {'type': 'formula', 'criteria': '=TRUE', 'format': fmt_bold_col})  # 产品编码
                    ws.conditional_format(1, 4, nrows, 4, {'type': 'formula', 'criteria': '=TRUE', 'format': fmt_bold_col})  # SKU名称
                    ws.conditional_format(1, 13, nrows, 13, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red_bold})
                    ws.conditional_format(1, 14, nrows, 14, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red_norm})
                    ws.conditional_format(1, 16, nrows, 16, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange_bold})
                    ws.conditional_format(1, 17, nrows, 17, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange_norm})
                    ws.conditional_format(1, 19, nrows, 19, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_blue})
                    ws.conditional_format(1, 20, nrows, 20, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_purple})

                # --- 写入 Sheet1: 补货计算表（✅含“在做(Y)”列） ---
                df_sheet1.to_excel(writer, index=False, sheet_name='补货计算表')
                format_sheet_with_active(writer.sheets['补货计算表'], df_sheet1)

                # --- 写入 Sheet2: 采购单 (保持原结构不变) ---
                df_buy = df_out_base[df_out_base['建议采购数'] > 0].copy()
                df_buy.to_excel(writer, index=False, sheet_name='采购单(找工厂)')
                format_sheet(writer.sheets['采购单(找工厂)'], df_buy, hide_cols=[14, 15, 16, 17, 18, 19])

                # --- 写入 Sheet3: 调拨单 (保持原结构不变) ---
                df_trans = df_out_base[df_out_base['建议调拨数量'] > 0].copy()
                df_trans.to_excel(writer, index=False, sheet_name='调拨单(发橙火)')
                format_sheet(writer.sheets['调拨单(发橙火)'], df_trans, hide_cols=[12, 13, 14, 15, 16, 17, 19])

                # --- 写入 Sheet4: 预警单 (保持原结构不变) ---
                df_fee = df_out_base[df_out_base['本月仓储费(预警)'] > 0].copy()
                df_fee.to_excel(writer, index=False, sheet_name='库龄预警单(需重入库)')
                format_sheet(writer.sheets['库龄预警单(需重入库)'], df_fee, hide_cols=[11, 12, 13, 14, 15, 16, 17, 18])

            st.download_button(
                "📥 下载最终 Excel (包含全量数据)",
                data=out_io.getvalue(),
                file_name=f"Coupang_Restock_Full_v18_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.ms-excel",
                type="primary"
            )
else:
    st.info("👈 请在左侧上传文件")
