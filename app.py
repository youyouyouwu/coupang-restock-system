import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 智能补货 (最终版)")
st.title("📦 Coupang 智能补货 (定制导出版)")
st.markdown("### 核心逻辑：基于Master表顺序，定制列排序与库存匹配规则")

# ==========================================
# 2. 列号配置 (请确认 Excel 实际位置)
# ==========================================
# A=0, B=1, C=2, D=3, E=4, F=5, G=6 ... M=12 ... R=17

# --- 1. 基础信息表 (Master) ---
IDX_M_CODE    = 0    # A列: 产品编码
IDX_M_SHOP    = 1    # B列: 店铺
IDX_M_COL_E   = 4    # E列: 基础信息E
IDX_M_COL_F   = 5    # F列: SKU名称
IDX_M_COST    = 6    # G列: 采购单价 (第5列)

IDX_M_ORANGE  = 3    # D列: 橙火ID (匹配橙火)
IDX_M_INBOUND = 12   # M列: 入库码 (匹配极风)

# --- 2. 销售表 (近7天) ---
IDX_7D_SKU    = 0    # A列: SKU/ID (默认匹配D列)
IDX_7D_QTY    = 8    # I列: 销售数量

# --- 3. 火箭仓/橙火库存表 ---
IDX_INV_R_SKU = 2    # C列: SKU/ID (与Master D列匹配)
IDX_INV_R_QTY = 7    # H列: 数量
IDX_INV_R_FEE = 17   # R列: 本月仓储费 (新增预警)

# --- 4. 极风库存表 ---
IDX_INV_J_BAR = 2    # C列: 条码/入库码 (与Master M列匹配)
IDX_INV_J_QTY = 10   # K列: 数量

# ==========================================
# 3. 工具函数
# ==========================================
def clean_match_key(series):
    """清洗匹配键"""
    s = series.astype(str).str.upper()
    s = s.str.replace(r'\.0$', '', regex=True)
    s = s.str.replace('"', '').str.strip()
    s = s.replace('NAN', '')
    return s

def clean_num(series):
    """清洗数值"""
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def clean_str(series):
    """清洗普通字符串"""
    return series.astype(str).str.replace('nan', '', case=False).str.strip()

def read_file(file):
    """读取文件"""
    if file is None: return pd.DataFrame()
    if file.name.endswith(('.xlsx', '.xls', '.xlsm')):
        try:
            file.seek(0)
            return pd.read_excel(file, dtype=str, engine='openpyxl')
        except: return pd.DataFrame()
    
    encodings = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'gbk', 'latin1']
    for enc in encodings:
        try:
            file.seek(0)
            return pd.read_csv(file, dtype=str, encoding=enc)
        except: continue
    return pd.DataFrame()

# ==========================================
# 4. 侧边栏
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数设置")
    
    # 1. 总补货设置 (采购)
    safety_weeks = st.number_input("🛡️ 总安全库存周数 (采购标准)", min_value=1, max_value=20, value=3, step=1)
    
    # 2. 橙火调拨设置 (内部发货)
    st.divider()
    orange_safety_weeks = st.number_input("🚚 橙火安全周数 (调拨预警)", min_value=1, max_value=10, value=2, step=1)
    
    # 3. 冗余设置 (滞销)
    st.divider()
    redundancy_weeks = st.number_input("⚠️ 库存冗余周数 (滞销标准)", min_value=4, max_value=52, value=8, step=1)
    
    st.divider()
    st.info("📂 请上传文件 (保持Master顺序)")
    file_master = st.file_uploader("1. 基础信息表 (Master) *必传", type=['xlsx','csv'])
    files_sales = st.file_uploader("2. 销售表 (近7天) *多选", type=['xlsx','csv'], accept_multiple_files=True)
    files_inv_r = st.file_uploader("3. 橙火/火箭仓库存 *多选", type=['xlsx','csv'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("4. 极风库存 *多选", type=['xlsx','csv'], accept_multiple_files=True)

# ==========================================
# 5. 主逻辑
# ==========================================
if file_master and files_sales and files_inv_r and files_inv_j:
    if st.button("🚀 生成定制报表", type="primary", use_container_width=True):
        with st.spinner("正在按指定列顺序匹配数据..."):
            
            # --- A. 读取 Master ---
            df_m = read_file(file_master)
            if df_m.empty: st.stop()
            
            df_base = pd.DataFrame()
            try:
                # 1. 提取展示列
                df_base['Shop'] = clean_str(df_m.iloc[:, IDX_M_SHOP])          
                df_base['Code'] = clean_match_key(df_m.iloc[:, IDX_M_CODE])    
                df_base['Info_E'] = clean_str(df_m.iloc[:, IDX_M_COL_E])       
                df_base['Info_F'] = clean_str(df_m.iloc[:, IDX_M_COL_F]) 
                df_base['Cost']   = clean_num(df_m.iloc[:, IDX_M_COST]) 
                
                df_base['Orange_ID'] = clean_match_key(df_m.iloc[:, IDX_M_ORANGE]) 
                df_base['Inbound_Code'] = clean_match_key(df_m.iloc[:, IDX_M_INBOUND]) 
                
            except IndexError:
                st.error("❌ 基础表列数不足，请检查列配置！"); st.stop()

            # --- B. 销售汇总 ---
            s_list = [read_file(f) for f in files_sales]
            if not s_list: st.stop()
            df_sales = pd.concat(s_list, ignore_index=True)
            df_sales['Key'] = clean_match_key(df_sales.iloc[:, IDX_7D_SKU])
            df_sales['Qty'] = clean_num(df_sales.iloc[:, IDX_7D_QTY])
            agg_sales = df_sales.groupby('Key')['Qty'].sum().reset_index()

            # --- C. 橙火库存 (包含 R列 仓储费) ---
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
                agg_orange = pd.DataFrame(columns=['Key','Qty','Fee'])

            # --- D. 极风库存 ---
            j_list = [read_file(f) for f in files_inv_j]
            if j_list:
                df_j = pd.concat(j_list, ignore_index=True)
                df_j['Key'] = clean_match_key(df_j.iloc[:, IDX_INV_J_BAR])
                df_j['Qty'] = clean_num(df_j.iloc[:, IDX_INV_J_QTY])
                agg_jifeng = df_j.groupby('Key')['Qty'].sum().reset_index()
            else:
                agg_jifeng = pd.DataFrame(columns=['Key','Qty'])

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
            
            # 1. 库存合计
            df_final['Total_Stock'] = df_final['Stock_Orange'] + df_final['Stock_Jifeng']
            
            # 2. 安全库存 & 冗余标准
            df_final['Safety'] = df_final['Sales_7d'] * safety_weeks
            df_final['Redundancy_Std'] = df_final['Sales_7d'] * redundancy_weeks
            
            # 3. 建议补货数 & 采购总额
            df_final['Restock_Qty'] = (df_final['Safety'] - df_final['Total_Stock']).apply(lambda x: int(x) if x > 0 else 0)
            df_final['Restock_Money'] = df_final['Restock_Qty'] * df_final['Cost']
            
            # 4. 冗余数量 & 冗余资金
            df_final['Redundancy_Qty'] = (df_final['Total_Stock'] - df_final['Redundancy_Std']).apply(lambda x: int(x) if x > 0 else 0)
            df_final['Redundancy_Money'] = df_final['Redundancy_Qty'] * df_final['Cost']
            
            # 5. 橙火调拨
            df_final['Orange_Safety_Std'] = df_final['Sales_7d'] * orange_safety_weeks
            df_final['Orange_Transfer_Qty'] = (df_final['Orange_Safety_Std'] - df_final['Stock_Orange']).apply(lambda x: int(x) if x > 0 else 0)

            # --- G. 整理输出 ---
            cols_export = [
                'Shop',           # 1
                'Code',           # 2
                'Info_E',         # 3
                'Info_F',         # 4
                'Cost',           # 5
                'Orange_ID',      # 6
                'Inbound_Code',   # 7
                'Sales_7d',       # 8
                'Stock_Orange',   # 9
                'Stock_Jifeng',   # 10
                'Total_Stock',    # 11
                'Safety',         # 12
                'Restock_Qty',    # 13 
                'Restock_Money',  # 14
                'Redundancy_Std', # 15
                'Redundancy_Qty', # 16 
                'Redundancy_Money', # 17
                'Orange_Safety_Std', # 18
                'Orange_Transfer_Qty', # 19
                'Storage_Fee'     # 20
            ]
            
            df_out = df_final[cols_export].copy()
            
            header_map = {
                'Shop': '店铺名称',
                'Code': '产品编码',
                'Info_E': '基础信息E列',
                'Info_F': 'SKU名称',
                'Cost': '采购单价',  
                'Orange_ID': '橙火ID (D列)',
                'Inbound_Code': '入库码 (M列)',
                'Sales_7d': '7天销量',
                'Stock_Orange': '橙火库存',
                'Stock_Jifeng': '极风库存',
                'Total_Stock': '库存合计',
                'Safety': f'总安全库存({safety_weeks}周)', 
                'Restock_Qty': '建议采购数',
                'Restock_Money': '预计采购总额(RMB)',
                'Redundancy_Std': f'冗余标准({redundancy_weeks}周)',
                'Redundancy_Qty': '冗余数量',
                'Redundancy_Money': '冗余资金',
                'Orange_Safety_Std': f'橙火安全库存({orange_safety_weeks}周)',
                'Orange_Transfer_Qty': '建议调拨数量',
                'Storage_Fee': '本月仓储费(预警)'
            }
            df_out.rename(columns=header_map, inplace=True)

            # --- H. 展示 ---
            st.divider()
            
            # === 1. 核心看板 (4个维度) ===
            # 计算逻辑：
            # A. 采购：SKU数量 + 金额
            buy_mask = df_out['建议采购数'] > 0
            k1_cnt = len(df_out[buy_mask])
            k1_val = df_out.loc[buy_mask, '预计采购总额(RMB)'].sum()
            
            # B. 冗余：SKU数量 + 资金
            red_mask = df_out['冗余数量'] > 0
            k2_cnt = len(df_out[red_mask])
            k2_val = df_out.loc[red_mask, '冗余资金'].sum()
            
            # C. 调拨：SKU数量 + 调拨总数
            trans_mask = df_out['建议调拨数量'] > 0
            k3_cnt = len(df_out[trans_mask])
            k3_val = df_out.loc[trans_mask, '建议调拨数量'].sum()
            
            # D. 库龄：SKU数量 + 仓储费总额 (韩币)
            fee_mask = df_out['本月仓储费(预警)'] > 0
            k4_cnt = len(df_out[fee_mask])
            k4_val = df_out.loc[fee_mask, '本月仓储费(预警)'].sum() # 求和费用

            m1, m2, m3, m4 = st.columns(4)
            
            # 指标1：采购
            m1.metric(
                label="📦 需采购 SKU / 金额",
                value=f"{k1_cnt} 个",
                delta=f"¥ {k1_val:,.0f}"
            )
            
            # 指标2：冗余 (反向颜色)
            m2.metric(
                label="⚠️ 冗余 SKU / 资金",
                value=f"{k2_cnt} 个",
                delta=f"¥ {k2_val:,.0f}",
                delta_color="inverse"
            )
            
            # 指标3：调拨
            m3.metric(
                label="🚚 需调拨 SKU / 数量",
                value=f"{k3_cnt} 个",
                delta=f"{k3_val:,.0f} 件"
            )
            
            # 指标4：库龄 (反向颜色)
            m4.metric(
                label="🚨 库龄预警 SKU / 总仓储费",
                value=f"{k4_cnt} 个",
                delta=f"₩ {k4_val:,.0f}", # 韩币
                delta_color="inverse"
            )

            # === 2. 表格展示 ===
            # 样式设置
            def highlight_restock(s):
                return ['background-color: #ffcccc; color: #b71c1c; font-weight: bold' if v > 0 else '' for v in s]
            
            def highlight_redundancy(s):
                return ['background-color: #ffe0b2; color: #e65100; font-weight: bold' if v > 0 else '' for v in s]

            def highlight_transfer(s):
                return ['background-color: #e3f2fd; color: #0d47a1; font-weight: bold' if v > 0 else '' for v in s]
            
            def highlight_fee(s):
                return ['background-color: #e1bee7; color: #4a148c; font-weight: bold' if v > 0 else '' for v in s]

            st.dataframe(
                df_out.style.apply(highlight_restock, subset=['建议采购数', '预计采购总额(RMB)']) 
                      .apply(highlight_redundancy, subset=['冗余数量', '冗余资金']) 
                      .apply(highlight_transfer, subset=['建议调拨数量'])
                      .apply(highlight_fee, subset=['本月仓储费(预警)'])
                      .format({
                          '橙火库存': '{:.0f}', '极风库存': '{:.0f}', '库存合计': '{:.0f}', 
                          f'总安全库存({safety_weeks}周)': '{:.0f}',
                          f'冗余标准({redundancy_weeks}周)': '{:.0f}',
                          f'橙火安全库存({orange_safety_weeks}周)': '{:.0f}',
                          '建议采购数': '{:.0f}', '预计采购总额(RMB)': '{:,.0f}', 
                          '7天销量': '{:.0f}', '采购单价': '{:,.0f}',
                          '冗余数量': '{:.0f}', '冗余资金': '{:,.0f}',
                          '建议调拨数量': '{:.0f}',
                          '本月仓储费(预警)': '{:,.0f}'
                      }),
                use_container_width=True, 
                height=600,
                hide_index=True
            )

            # Excel 导出
            out_io = io.BytesIO()
            with pd.ExcelWriter(out_io, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False, sheet_name='补货计算表')
                
                # Sheet2: 采购单
                df_buy = df_out[df_out['建议采购数'] > 0].copy()
                df_buy.to_excel(writer, index=False, sheet_name='采购单(找工厂)')
                
                # Sheet3: 调拨单
                df_trans = df_out[df_out['建议调拨数量'] > 0].copy()
                df_trans.to_excel(writer, index=False, sheet_name='调拨单(发橙火)')
                
                # Sheet4: 库龄预警单
                df_fee = df_out[df_out['本月仓储费(预警)'] > 0].copy()
                df_fee.to_excel(writer, index=False, sheet_name='库龄预警单(需重入库)')
                
                wb = writer.book
                ws = writer.sheets['补货计算表']
                
                fmt_red = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True})
                fmt_orange = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': True})
                fmt_blue = wb.add_format({'bg_color': '#C5D9F1', 'font_color': '#1F497D', 'bold': True})
                fmt_purple = wb.add_format({'bg_color': '#E1BEE7', 'font_color': '#4A148C', 'bold': True})
                
                ws.conditional_format(1, 12, len(df_out), 13, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red})
                ws.conditional_format(1, 15, len(df_out), 16, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange})
                ws.conditional_format(1, 18, len(df_out), 18, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_blue})
                ws.conditional_format(1, 19, len(df_out), 19, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_purple})
                
                fmt_head = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                ws.set_row(0, None, fmt_head)
                ws.set_column('A:T', 13)

            st.download_button(
                "📥 下载最终 Excel",
                data=out_io.getvalue(),
                file_name=f"Coupang_Restock_Full_v5_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.ms-excel",
                type="primary"
            )
else:
    st.info("👈 请在左侧上传文件")
