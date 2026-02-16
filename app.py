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
# A=0, B=1, C=2, D=3, E=4, F=5, G=6 ... M=12

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

# --- 4. 极风库存表 ---
IDX_INV_J_BAR = 2    # C列: 条码/入库码 (与Master M列匹配)
IDX_INV_J_QTY = 10   # K列: 数量

# ==========================================
# 3. 工具函数
# ==========================================
def clean_match_key(series):
    """清洗匹配键: 去空格、转大写、去.0、去nan"""
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
    safety_weeks = st.number_input("🛡️ 安全库存周数 (Weeks)", min_value=1, max_value=20, value=3, step=1)
    
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
                df_base['Cost']   = clean_num(df_m.iloc[:, IDX_M_COST])  # 第5列：采购单价
                
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

            # --- C. 橙火库存 ---
            r_list = [read_file(f) for f in files_inv_r]
            if r_list:
                df_r = pd.concat(r_list, ignore_index=True)
                df_r['Key'] = clean_match_key(df_r.iloc[:, IDX_INV_R_SKU])
                df_r['Qty'] = clean_num(df_r.iloc[:, IDX_INV_R_QTY])
                agg_orange = df_r.groupby('Key')['Qty'].sum().reset_index()
            else:
                agg_orange = pd.DataFrame(columns=['Key','Qty'])

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
            df_final.rename(columns={'Qty': 'Stock_Orange'}, inplace=True)
            
            df_final = pd.merge(df_final, agg_jifeng, left_on='Inbound_Code', right_on='Key', how='left', suffixes=('', '_J'))
            df_final.rename(columns={'Qty': 'Stock_Jifeng'}, inplace=True)

            # --- F. 计算逻辑 ---
            df_final['Sales_7d'] = df_final['Sales_7d'].fillna(0)
            df_final['Stock_Orange'] = df_final['Stock_Orange'].fillna(0)
            df_final['Stock_Jifeng'] = df_final['Stock_Jifeng'].fillna(0)
            
            # 1. 库存合计
            df_final['Total_Stock'] = df_final['Stock_Orange'] + df_final['Stock_Jifeng']
            
            # 2. 安全库存
            df_final['Safety'] = df_final['Sales_7d'] * safety_weeks
            
            # 3. 建议补货数
            df_final['Restock_Qty'] = (df_final['Safety'] - df_final['Total_Stock']).apply(lambda x: int(x) if x > 0 else 0)
            
            # 4. 采购总额 = 补货数 * 采购单价(Cost)
            df_final['Restock_Money'] = df_final['Restock_Qty'] * df_final['Cost']

            # --- G. 整理输出 ---
            cols_export = [
                'Shop',           # 1
                'Code',           # 2
                'Info_E',         # 3
                'Info_F',         # 4
                'Cost',           # 5. 采购金额 (单价)
                'Orange_ID',      # 6
                'Inbound_Code',   # 7
                'Sales_7d',       # 8
                'Stock_Orange',   # 9
                'Stock_Jifeng',   # 10
                'Total_Stock',    # 11
                'Safety',         # 12
                'Restock_Qty',    # 13. 建议补货数
                'Restock_Money',  # 14. 采购总额 (补货数 * 单价)
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
                'Safety': f'安全库存({safety_weeks}周)', 
                'Restock_Qty': '建议补货数',
                'Restock_Money': '预计采购总额(RMB)' # 明确改为RMB
            }
            df_out.rename(columns=header_map, inplace=True)

            # --- H. 展示 ---
            st.divider()
            c1, c2 = st.columns(2)
            c1.metric("📦 总需补货件数", f"{df_out['建议补货数'].sum():,.0f}")
            c2.metric("💰 预计采购总额 (RMB)", f"¥ {df_out['预计采购总额(RMB)'].sum():,.0f}") # 使用人民币符号

            def highlight_restock(s):
                return ['background-color: #ffcccc; color: red; font-weight: bold' if v > 0 else '' for v in s]

            st.dataframe(
                df_out.style.apply(highlight_restock, subset=['建议补货数'])
                      .format({
                          '橙火库存': '{:.0f}', '极风库存': '{:.0f}', '库存合计': '{:.0f}', f'安全库存({safety_weeks}周)': '{:.0f}',
                          '建议补货数': '{:.0f}', '预计采购总额(RMB)': '{:,.0f}', 
                          '7天销量': '{:.0f}', '采购单价': '{:,.0f}'
                      }),
                use_container_width=True, 
                height=600,
                hide_index=True
            )

            # Excel 导出
            out_io = io.BytesIO()
            with pd.ExcelWriter(out_io, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False, sheet_name='补货计算表')
                
                df_buy = df_out[df_out['建议补货数'] > 0].copy()
                df_buy.to_excel(writer, index=False, sheet_name='采购单')
                
                wb = writer.book
                ws = writer.sheets['补货计算表']
                
                # 红色高亮 (建议补货数在第13列，索引12)
                fmt_red = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True})
                ws.conditional_format(1, 12, len(df_out), 12, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red})
                
                fmt_head = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                ws.set_row(0, None, fmt_head)
                ws.set_column('A:N', 13)

            st.download_button(
                "📥 下载最终 Excel",
                data=out_io.getvalue(),
                file_name=f"Coupang_Restock_RMB_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.ms-excel",
                type="primary"
            )
else:
    st.info("👈 请在左侧上传文件")
