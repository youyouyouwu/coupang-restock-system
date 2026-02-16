import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 智能补货系统 (Master顺序版)")
st.title("📦 Coupang 智能补货系统 (Master顺序版)")
st.markdown("### 核心逻辑：顺序完全参照【基础信息表】，自动高亮需补货行")

# ==========================================
# 2. 列号配置 (!!! 请务必根据您的Excel实际列号修改 !!!)
# ==========================================
# A列=0, B列=1, C列=2, D列=3 ... 以此类推

# --- 1. 基础信息表 (Master) ---
IDX_M_CODE    = 0    # A列: 内部编码
IDX_M_SHOP    = 1    # B列: 登品店铺
IDX_M_NAME    = 2    # C列: 产品名称 (★请确认您的Excel位置)
IDX_M_SKU     = 3    # D列: SKU ID (关联键)
IDX_M_ORANGE  = 4    # E列: 橙火ID (★请确认您的Excel位置)
IDX_M_INBOUND = 5    # F列: 产品入库码 (★请确认您的Excel位置)
IDX_M_COST    = 6    # G列: 采购成本
IDX_M_PROFIT  = 10   # K列: 单品毛利
IDX_M_BAR     = 12   # M列: 条码/自发货ID

# --- 2. 销售表 (近7天) ---
IDX_7D_SKU    = 0    # A列: SKU
IDX_7D_QTY    = 8    # I列: 销售数量

# --- 3. 火箭仓库存表 ---
IDX_INV_R_SKU = 2    # C列: SKU所在列
IDX_INV_R_QTY = 7    # H列: 数量所在列

# --- 4. 极风/自发货库存表 ---
IDX_INV_J_BAR = 2    # C列: 条码所在列
IDX_INV_J_QTY = 10   # K列: 数量所在列

# ==========================================
# 3. 工具函数
# ==========================================
def clean_match_key(series):
    """清洗用于匹配的Key"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    """清洗数值列"""
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def clean_str(series):
    """普通清洗字符串"""
    return series.astype(str).str.replace('nan', '', case=False).str.strip()

def read_file(file):
    """通用读取函数 (防乱码)"""
    if file is None: return pd.DataFrame()
    
    if file.name.endswith(('.xlsx', '.xls', '.xlsm')):
        try:
            file.seek(0)
            return pd.read_excel(file, dtype=str, engine='openpyxl')
        except Exception as e:
            st.error(f"Excel读取失败: {file.name}, 错误: {e}")
            return pd.DataFrame()

    encodings_to_try = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'gbk', 'latin1']
    for encoding in encodings_to_try:
        try:
            file.seek(0)
            return pd.read_csv(file, dtype=str, encoding=encoding)
        except:
            continue
            
    st.error(f"❌ 无法读取文件: {file.name}")
    return pd.DataFrame()

# ==========================================
# 4. 侧边栏 & 参数设置
# ==========================================
with st.sidebar:
    st.header("⚙️ 补货参数设置")
    safety_days = st.number_input("🛡️ 安全库存天数", min_value=7, max_value=60, value=20, step=1)
    
    st.divider()
    st.header("📂 数据上传")
    st.info("⚠️ 注意：输出结果将严格按照【基础信息表】的顺序排列")
    
    file_master = st.file_uploader("1. 基础信息表 (Master) *必传", type=['xlsx', 'csv', 'xls'])
    files_sales_7d = st.file_uploader("2. 销售表 (近7天) *多选", type=['xlsx', 'csv', 'xls'], accept_multiple_files=True)
    files_inv_r = st.file_uploader("3. 火箭仓库存 *多选", type=['xlsx', 'csv', 'xls'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("4. 极风/自发库存 *多选", type=['xlsx', 'csv', 'xls'], accept_multiple_files=True)

# ==========================================
# 5. 主逻辑
# ==========================================
if file_master and files_sales_7d and files_inv_r and files_inv_j:
    if st.button("🚀 开始计算 (保持Master顺序)", type="primary", use_container_width=True):
        with st.spinner("正在计算并匹配 Master 顺序..."):
            
            # --- A. 读取 Master (保持原索引) ---
            df_m = read_file(file_master)
            if df_m.empty: st.stop()

            # 构造基础数据 (保留原始索引以确保顺序)
            df_base = pd.DataFrame()
            try:
                # 匹配键
                df_base['SKU_ID'] = clean_match_key(df_m.iloc[:, IDX_M_SKU])
                df_base['Barcode'] = clean_match_key(df_m.iloc[:, IDX_M_BAR])
                
                # 展示字段
                df_base['Code'] = clean_match_key(df_m.iloc[:, IDX_M_CODE])
                df_base['Shop'] = clean_str(df_m.iloc[:, IDX_M_SHOP])
                df_base['Name'] = clean_str(df_m.iloc[:, IDX_M_NAME])
                df_base['Orange_ID'] = clean_str(df_m.iloc[:, IDX_M_ORANGE])
                df_base['Inbound_Code'] = clean_str(df_m.iloc[:, IDX_M_INBOUND])
                
                # 计算字段
                df_base['Cost'] = clean_num(df_m.iloc[:, IDX_M_COST])
            except IndexError:
                st.error("❌ 基础表列数不足！请检查列号配置。")
                st.stop()
            
            # --- B. 销售汇总 ---
            s_list = [read_file(f) for f in files_sales_7d]
            if not s_list: st.error("❌ 销售表为空"); st.stop()
            df_7d_all = pd.concat(s_list, ignore_index=True)
            df_7d_all['Match_SKU'] = clean_match_key(df_7d_all.iloc[:, IDX_7D_SKU])
            df_7d_all['Qty_7Days'] = clean_num(df_7d_all.iloc[:, IDX_7D_QTY])
            sales_agg = df_7d_all.groupby('Match_SKU')['Qty_7Days'].sum().reset_index()
            
            # --- C. 库存汇总 ---
            # 火箭
            r_list = [read_file(f) for f in files_inv_r]
            if r_list:
                df_r = pd.concat(r_list, ignore_index=True)
                df_r['Match_SKU'] = clean_match_key(df_r.iloc[:, IDX_INV_R_SKU])
                df_r['Rocket_Stock'] = clean_num(df_r.iloc[:, IDX_INV_R_QTY])
                inv_r_agg = df_r.groupby('Match_SKU')['Rocket_Stock'].sum().reset_index()
            else:
                inv_r_agg = pd.DataFrame(columns=['Match_SKU', 'Rocket_Stock'])
            
            # 极风
            j_list = [read_file(f) for f in files_inv_j]
            if j_list:
                df_j = pd.concat(j_list, ignore_index=True)
                df_j['Match_Bar'] = clean_match_key(df_j.iloc[:, IDX_INV_J_BAR])
                df_j['Jifeng_Stock'] = clean_num(df_j.iloc[:, IDX_INV_J_QTY])
                inv_j_agg = df_j.groupby('Match_Bar')['Jifeng_Stock'].sum().reset_index()
            else:
                inv_j_agg = pd.DataFrame(columns=['Match_Bar', 'Jifeng_Stock'])

            # --- D. 合并 (关键：使用 Left Join 保持 df_base 的顺序和行数) ---
            df_final = pd.merge(df_base, sales_agg, left_on='SKU_ID', right_on='Match_SKU', how='left')
            df_final = pd.merge(df_final, inv_r_agg, left_on='SKU_ID', right_on='Match_SKU', how='left')
            df_final = pd.merge(df_final, inv_j_agg, left_on='Barcode', right_on='Match_Bar', how='left')
            
            # 填充0
            fill_cols = ['Qty_7Days', 'Rocket_Stock', 'Jifeng_Stock']
            df_final[fill_cols] = df_final[fill_cols].fillna(0)
            
            # 计算
            df_final['Daily_Avg'] = df_final['Qty_7Days'] / 7
            df_final['Safety_Line'] = df_final['Daily_Avg'] * safety_days
            df_final['Total_Stock'] = df_final['Rocket_Stock'] + df_final['Jifeng_Stock']
            
            df_final['Restock_Qty'] = (df_final['Safety_Line'] - df_final['Total_Stock']).apply(lambda x: int(x) if x > 0 else 0)
            df_final['Restock_Cost'] = df_final['Restock_Qty'] * df_final['Cost']

            # --- E. 样式与展示 ---
            
            # 统计
            total_money = df_final['Restock_Cost'].sum()
            total_items = df_final['Restock_Qty'].sum()
            need_restock_count = len(df_final[df_final['Restock_Qty'] > 0])
            
            st.divider()
            c1, c2, c3 = st.columns(3)
            c1.metric("💰 补货总金额", f"₩ {total_money:,.0f}")
            c2.metric("📦 需补货SKU", f"{need_restock_count} / {len(df_final)} 个")
            c3.metric("🚛 补货总件数", f"{total_items:,.0f} 件")
            
            st.subheader("📋 全量数据预览 (已高亮需补货行)")
            
            # 定义展示列
            cols_show = ['Code', 'Name', 'Shop', 'Orange_ID', 'Inbound_Code', 'Restock_Qty', 'Restock_Cost', 'Qty_7Days', 'Total_Stock']
            df_display = df_final[cols_show].copy()

            # --- 样式函数 (核心) ---
            def highlight_rows(row):
                # 如果补货数 > 0，整行浅红背景
                if row['Restock_Qty'] > 0:
                    return ['background-color: #ffe6e6'] * len(row)
                return [''] * len(row)

            def highlight_col(s):
                # 单独给 Restock_Qty 列加粗变红
                return ['color: #d32f2f; font-weight: bold' if v > 0 else '' for v in s]

            # 应用样式
            st_df = df_display.style.apply(highlight_rows, axis=1)\
                                    .apply(highlight_col, subset=['Restock_Qty'])\
                                    .format({'Restock_Cost': '{:,.0f}', 'Qty_7Days': '{:.0f}', 'Total_Stock': '{:.0f}'})
            
            st.dataframe(st_df, use_container_width=True, height=600)

            # --- F. Excel 导出 ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 1. 完整顺序表 (Sheet1)
                final_output_cols = ['Code', 'Name', 'Shop', 'SKU_ID', 'Barcode', 'Orange_ID', 'Inbound_Code', 
                                     'Qty_7Days', 'Safety_Line', 'Total_Stock', 
                                     'Rocket_Stock', 'Jifeng_Stock', 
                                     'Restock_Qty', 'Restock_Cost']
                
                df_export = df_final[final_output_cols].copy()
                df_export.columns = ['产品编号', '产品名称', '店铺', 'SKU', '条码', '橙火ID', '入库码', 
                                     '7天销量', '安全库存线', '总库存', 
                                     'Rocket库存', '极风库存', 
                                     '建议补货数', '补货金额']
                
                df_export.to_excel(writer, index=False, sheet_name='全量补货表(Master顺序)')
                
                # 2. 纯补货单 (Sheet2 - 仅需补货的)
                df_buy_only = df_export[df_export['建议补货数'] > 0].copy()
