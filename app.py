import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 智能补货 (定制列版)")
st.title("📦 Coupang 智能补货 (定制导出版)")
st.markdown("### 核心逻辑：基于Master表顺序，定制列排序与库存匹配规则")

# ==========================================
# 2. 列号配置 (请确认 Excel 实际位置)
# ==========================================
# A=0, B=1, C=2, D=3, E=4, F=5 ... M=12

# --- 1. 基础信息表 (Master) ---
# 您指定的关键列：
IDX_M_SHOP    = 1    # B列: 店铺 (放在第1列)
IDX_M_COL_E   = 4    # E列: 基础信息E (放在第2列)
IDX_M_COL_F   = 5    # F列: 基础信息F (放在第3列)
IDX_M_ORANGE  = 3    # D列: 橙火ID (放在第4列 & 匹配橙火库存)
IDX_M_INBOUND = 12   # M列: 入库码 (放在第5列 & 匹配极风库存)

# 其他辅助列 (用于计算)
IDX_M_COST    = 6    # G列: 采购成本
IDX_M_PROFIT  = 10   # K列: 单品毛利

# --- 2. 销售表 (近7天) ---
IDX_7D_SKU    = 0    # A列: SKU/ID (默认匹配D列橙火ID)
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
    """清洗匹配键: 去空格、转大写、去.0"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    """清洗数值"""
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def clean_str(series):
    """清洗普通字符串"""
    return series.astype(str).str.replace('nan', '', case=False).str.strip()

def read_file(file):
    """读取文件 (支持多种编码)"""
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
    safety_days = st.number_input("🛡️ 安全库存天数", 7, 60, 20)
    
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
            
            # --- A. 读取 Master (保留原始顺序) ---
            df_m = read_file(file_master)
            if df_m.empty: st.stop()
            
            df_base = pd.DataFrame()
            try:
                # 1. 提取用于展示的列 (按您要求的顺序)
                df_base['Shop'] = clean_str(df_m.iloc[:, IDX_M_SHOP])          # 第1列: 店铺 (B)
                df_base['Info_E'] = clean_str(df_m.iloc[:, IDX_M_COL_E])       # 第2列: E列
                df_base['Info_F'] = clean_str(df_m.iloc[:, IDX_M_COL_F])       # 第3列: F列
                df_base['Orange_ID'] = clean_match_key(df_m.iloc[:, IDX_M_ORANGE]) # 第4列: 橙火ID (D)
                df_base['Inbound_Code'] = clean_match_key(df_m.iloc[:, IDX_M_INBOUND]) # 第5列: 入库码 (M)
                
                # 2. 提取计算用数据
                df_base['Cost'] = clean_num(df_m.iloc[:, IDX_M_COST])
                
                # 3. 设置匹配键 (Key)
                # 橙火库存 & 销量 -> 匹配 D列 (Orange_ID)
                # 极风库存 -> 匹配 M列 (Inbound_Code)
            except IndexError:
                st.error("❌ 基础表列数不足，请检查列配置！"); st.stop()

            # --- B. 销售汇总 (假设销量匹配橙火ID/D列) ---
            s_list = [read_file(f) for f in files_sales]
            if not s_list: st.stop()
            df_sales = pd.concat(s_list, ignore_index=True)
            # 清洗
            df_sales['Key'] = clean_match_key(df_sales.iloc[:, IDX_7D_SKU])
            df_sales['Qty'] = clean_num(df_sales.iloc[:, IDX_7D_QTY])
            agg_sales = df_sales.groupby('Key')['Qty'].sum().reset_index()

            # --- C. 橙火库存 (匹配 D列) ---
            r_list = [read_file(f) for f in files_inv_r]
            if r_list:
                df_r = pd.concat(r_list, ignore_index=True)
                df_r['Key'] = clean_match_key(df_r.iloc[:, IDX_INV_R_SKU])
                df_r['Qty'] = clean_num(df_r.iloc[:, IDX_INV_R_QTY])
                agg_orange = df_r.groupby('Key')['Qty'].sum().reset_index()
            else:
                agg_orange = pd.DataFrame(columns=['Key','Qty'])

            # --- D. 极风库存 (匹配 M列) ---
            j_list = [read_file(f) for f in files_inv_j]
            if j_list:
                df_j = pd.concat(j_list, ignore_index=True)
                df_j['Key'] = clean_match_key(df_j.iloc[:, IDX_INV_J_BAR])
                df_j['Qty'] = clean_num(df_j.iloc[:, IDX_INV_J_QTY])
                agg_jifeng = df_j.groupby('Key')['Qty'].sum().reset_index()
            else:
                agg_jifeng = pd.DataFrame(columns=['Key','Qty'])

            # --- E. 匹配合并 (Left Join 保留顺序) ---
            # 1. 匹配销量 (用 D列 Orange_ID)
            df_final = pd.merge(df_base, agg_sales, left_on='Orange_ID', right_on='Key', how='left')
            df_final.rename(columns={'Qty': 'Sales_7d'}, inplace=True)
            
            # 2. 匹配橙火库存 (用 D列 Orange_ID)
            df_final = pd.merge(df_final, agg_orange, left_on='Orange_ID', right_on='Key', how='left', suffixes=('', '_R'))
            df_final.rename(columns={'Qty': 'Stock_Orange'}, inplace=True)
            
            # 3. 匹配极风库存 (用 M列 Inbound_Code)
            df_final = pd.merge(df_final, agg_jifeng, left_on='Inbound_Code', right_on='Key', how='left', suffixes=('', '_J'))
            df_final.rename(columns={'Qty': 'Stock_Jifeng'}, inplace=True)

            # --- F. 计算补货 ---
            df_final['Sales_7d'] = df_final['Sales_7d'].fillna(0)
            df_final['Stock_Orange'] = df_final['Stock_Orange'].fillna(
