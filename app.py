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
IDX_M_COST    = 6    # G列: 采购单价

IDX_M_ORANGE  = 3    # D列: 橙火ID
IDX_M_INBOUND = 12   # M列: 入库码

# --- 2. 销售表 (近7天) ---
IDX_7D_SKU    = 0    # A列: SKU
IDX_7D_QTY    = 8    # I列: 销量

# --- 3. 火箭仓/橙火库存表 ---
IDX_INV_R_SKU = 2    # C列: SKU
IDX_INV_R_QTY = 7    # H列: 数量

# --- 4. 极风库存表 ---
IDX_INV_J_BAR = 2    # C列: 条码
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
    
    # 补货标准
    safety_weeks = st.number_input("🛡️ 安全库存周数 (补货标准)", min_value=1, max_value=20, value=3, step=1)
    
    # 滞销标准
    redundancy_weeks = st.number_input("⚠️ 库存冗余周数 (滞销标准)", min_value=4, max_value=52, value=8, step=1, help="超过此标准的库存量将被标记为冗余")
    
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
        with st.spinner("正在计算补货与滞销数据..."):
            
            # --- A. 读取 Master ---
            df_m = read_file(file_master)
            if df_m.empty: st.stop()
            
            df_base = pd.DataFrame()
            try:
                df_base['Shop'] = clean_str(df_m.iloc[:, IDX_M_SHOP])          
                df_base['Code'] = clean_match_key(df_m.iloc[:, IDX_M_CODE])    
                df_base['Info_E'] = clean_str(df_m.iloc[:, IDX_M_COL_E])       
                df_base['Info_F'] = clean_str(df_m.iloc[:, IDX_M_COL_F]) 
                df_base['Cost']   = clean_num(df_m.iloc[:, IDX_M_COST])  
                df_base['Orange_ID'] = clean_match_key(df_m.iloc[:, IDX_M_ORANGE]) 
                df_base['Inbound_Code'] = clean_match_key(df_m.iloc[:, IDX_M_INBOUND]) 
            except IndexError:
                st.error("❌ 基础表列数不足"); st.stop()

            # --- B. 销售汇总
