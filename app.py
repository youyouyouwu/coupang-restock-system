import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 智能补货系统 (Pro)")
st.title("📦 Coupang 智能补货系统 (Pro版)")
st.markdown("### 核心逻辑：基于【近7天销量】预测安全库存，生成含【橙火ID/入库码】的详细工单")

# ==========================================
# 2. 列号配置 (!!! 请务必根据您的Excel实际列号修改 !!!)
# ==========================================
# A列=0, B列=1, C列=2, D列=3 ... 以此类推

# --- 1. 基础信息表 (Master) ---
IDX_M_CODE    = 0    # A列: 内部编码 (关联键)
IDX_M_SHOP    = 1    # B列: 登品店铺
IDX_M_NAME    = 2    # C列: 产品名称 (★请确认您的Excel是不是在C列，不是请修改数字)
IDX_M_SKU     = 3    # D列: SKU ID (关联键)
IDX_M_ORANGE  = 4    # E列: 橙火ID (★请确认您的Excel此列位置)
IDX_M_INBOUND = 5    # F列: 产品入库码 (★请确认您的Excel此列位置)
IDX_M_COST    = 6    # G列: 采购成本
IDX_M_PROFIT  = 10   # K列: 单品毛利
IDX_M_BAR     = 12   # M列: 条码/自发货ID

# --- 2. 销售表 (近7天) ---
IDX_7D_SKU    = 0    # A列: 注册商品ID / SKU
IDX_7D_QTY    = 8    # I列: 销售数量

# --- 3. 火箭仓库存表 ---
IDX_INV_R_SKU = 2    # C列: SKU所在列
IDX_INV_R_QTY = 7    # H列: 数量所在列

# --- 4. 极风/自发货库存表 ---
IDX_INV_J_BAR = 2    # C列: 条码所在列
IDX_INV_J_QTY = 10   # K列: 数量所在列

# ==========================================
# 3. 工具函数 (已修复编码问题)
# ==========================================
def clean_match_key(series):
    """清洗用于匹配的Key"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    """清洗数值列"""
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def clean_str(series):
    """普通清洗字符串(保留原样，只去空格)"""
    return series.astype(str).str.replace('nan', '', case=False).str.strip()

def read_file(file):
    """通用读取函数 (防乱码增强版)"""
    if file is None: return pd.DataFrame()
    
    # Excel
    if file.name.endswith(('.xlsx', '.xls', '.xlsm')):
        try:
            file.seek(0)
            return pd.read_excel(file, dtype=str, engine='openpyxl')
        except Exception as e:
            st.error(f"Excel读取失败: {file.name}, 错误: {e}")
            return pd.DataFrame()

    # CSV
    encodings_to_try = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'gbk', 'gb18030', 'latin1']
    for encoding in encodings_to_try:
        try:
            file.seek(0)
            return pd.read_csv(file, dtype=str, encoding=encoding)
        except:
            continue
            
    st.error(f"❌ 无法读取文件: {file.name}。请另存为标准Excel(.xlsx)后重试。")
    return pd.DataFrame()

# ==========================================
# 4. 侧边栏 & 参数设置
# ==========================================
with st.sidebar:
    st.header("⚙️ 补货参数设置")
    safety_days = st.number_input("🛡️ 安全库存天数", min_value=7, max_value=60, value=20, step=1)
    
    st.divider()
    st.header("📂 数据上传")
    st.info("⚠️ 注意：除基础表外，其他支持多文件上传")
    
    file_master = st.file_uploader("1. 基础信息表 (Master) *必传", type=['xlsx', 'csv', 'xls'])
    files_sales_7d = st.file_uploader("2. 销售表 (近7天) *多选", type=['xlsx', 'csv', 'xls'], accept_multiple_files=True)
    files_inv_r = st.file_uploader("3. 火箭仓库存 *多选", type=['xlsx', 'csv', 'xls'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("4. 极风/自发库存 *多选", type=['xlsx', 'csv', 'xls'], accept_multiple_files=True)

# ==========================================
# 5. 主逻辑
# ==========================================
if file_master and files_sales_7d and files_inv_r and files_inv_j:
    if st.button("🚀 开始计算补货工单", type="primary", use_container_width=True):
        with st.spinner("正在提取橙火ID、入库码并计算补货量..."):
            
            # ----------------------------------
            # A. 读取基础信息 (Master)
            # ----------------------------------
            df_m = read_file(file_master)
            if df_m.empty: st.stop()

            # 提取所有需要的字段 (SKU层级)
            df_base = pd.DataFrame()
            try:
                # 关键匹配键
                df_base['SKU_ID'] = clean_match_key(df_m.iloc[:, IDX_M_SKU])
                df_base['Barcode'] = clean_match_key(df_m.iloc[:, IDX_M_BAR])
                
                # 基础信息 (完全引用Master)
                df_base['Code'] = clean_match_key(df_m.iloc[:, IDX_M_CODE])     # 产品编号
                df_base['Shop'] = clean_str(df_m.iloc[:, IDX_M_SHOP])           # 登品店铺
                df_base['Name'] = clean_str(df_m.iloc[:, IDX_M_NAME])           # 产品名称
                df_base['Orange_ID'] = clean_str(df_m.iloc[:, IDX_M_ORANGE])    # 橙火ID
                df_base['Inbound_Code'] = clean_str(df_m.iloc[:, IDX_M_INBOUND])# 入库码
                
                # 数值
                df_base['Cost'] = clean_num(df_m.iloc[:, IDX_M_COST]) 
                df_base['Unit_Profit'] = clean_num(df_m.iloc[:, IDX_M_PROFIT])
            except IndexError:
                st.error("❌ 基础表列数不足！请检查代码第2部分【列号配置】中的数字是否正确。")
                st.stop()
            
            # ----------------------------------
            # B. 读取销售表 (近7天) - 汇总
            # ----------------------------------
            s_list = []
            for f in files_sales_7d:
                df = read_file(f)
                if not df.empty: s_list.append(df)
            
            if not s_list:
                st.error("❌ 销售表为空"); st.stop()

            df_7d_all = pd.concat(s_list, ignore_index=True)
            df_7d_all['Match_SKU'] = clean_match_key(df_7d_all.iloc[:, IDX_7D_SKU])
            df_7d_all['Qty_7Days'] = clean_num(df_7d_all.iloc[:, IDX_7D_QTY])
            
            # SKU层级销量汇总
            sales_velocity = df_7d_all.groupby('Match_SKU')['Qty_7Days'].sum().reset_index()
            
            # ----------------------------------
            # C. 读取库存 (Stock) - 汇总
            # ----------------------------------
            # 火箭仓
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

            # ----------------------------------
            # D. 合并计算
            # ----------------------------------
            # 左连接：以Master表为准，保留所有Master里的SKU信息
            df_final = pd.merge(df_base, sales_velocity, left_on='SKU_ID', right_on='Match_SKU', how='left')
            df_final = pd.merge(df_final, inv_r_agg, left_on='SKU_ID', right_on='Match_SKU', how='left')
            df_final = pd.merge(df_final, inv_j_agg, left_on='Barcode', right_on='Match_Bar', how='left')
            
            # 填充缺失值
            fill_cols = ['Qty_7Days', 'Rocket_Stock', 'Jifeng_Stock']
            df_final[fill_cols] = df_final[fill_cols].fillna(0)
            
            # 计算逻辑
            df_final['Daily_Avg'] = df_final['Qty_7Days'] / 7
            df_final['Safety_Line'] = df_final['Daily_Avg'] * safety_days
            df_final['Total_Stock'] = df_final['Rocket_Stock'] + df_final['Jifeng_Stock']
            
            df_final['Restock_Qty'] = (df_final['Safety_Line'] - df_final['Total_Stock']).apply(lambda x: int(x) if x > 0 else 0)
            df_final['Restock_Cost'] = df_final['Restock_Qty'] * df_final['Cost']
            
            # ----------------------------------
            # E. 展示与导出
            # ----------------------------------
            df_restock = df_final[df_final['Restock_Qty'] > 0].sort_values(by='Restock_Cost', ascending=False)
            
            st.divider()
            c1, c2, c3 = st.columns(3)
            c1.metric("💰 预计补货总金额", f"₩ {df_restock['Restock_Cost'].sum():,.0f}")
            c2.metric("📦 需补货SKU数", f"{len(df_restock)} 个")
            c3.metric("🚛 总补货件数", f"{df_restock['Restock_Qty'].sum():,.0f} 件")
            
            st.subheader("📋 建议补货清单 (预览前50条)")
            
            # 展示用的列 (包含新加的列)
            preview_cols = ['Code', 'Name', 'Shop', 'Orange_ID', 'Inbound_Code', 'Restock_Qty', 'Restock_Cost', 'Rocket_Stock', 'Jifeng_Stock']
            # 防止列名不存在
            valid_preview = [c for c in preview_cols if c in df_restock.columns]
            
            # 使用原生dataframe展示，避免Style报错
            st.dataframe(df_restock[valid_preview].head(50), use_container_width=True)
            
            # Excel 导出
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 1. 补货工单 (给采购/仓库看，包含详细ID)
                cols_order = ['Code', 'Name', 'Shop', 'SKU_ID', 'Barcode', 'Orange_ID', 'Inbound_Code', 'Cost', 'Restock_Qty', 'Restock_Cost']
                df_order = df_restock[cols_order].copy()
                df_order.columns = ['产品编号', '产品名称', '店铺', 'SKU', '条码', '橙火ID', '入库码', '采购单价', '补货数量', '预计金额']
                df_order.to_excel(writer, index=False, sheet_name='补货工单')
                
                # 2. 全量数据 (分析用)
                cols_full = ['Code', 'Name', 'Shop', 'SKU_ID', 'Barcode', 'Orange_ID', 'Inbound_Code', 
                             'Qty_7Days', 'Safety_Line', 'Total_Stock', 'Restock_Qty']
                df_full = df_final[cols_full].copy()
                df_full.columns = ['产品编号', '产品名称', '店铺', 'SKU', '条码', '橙火ID', '入库码', 
                                   '7天销量', '安全库存线', '总库存', '建议补货']
                df_full.to_excel(writer, index=False, sheet_name='全量数据')
                
                # 格式设置
                wb = writer.book
                ws = writer.sheets['补货工单']
                fmt_header = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1})
                ws.set_row(0, None, fmt_header)
                ws.set_column('A:J', 15)

            st.download_button(
                label="📥 下载详细补货工单 (Excel)",
                data=output.getvalue(),
                file_name=f"Restock_Order_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.ms-excel",
                type="primary"
            )

else:
    st.info("👈 请在左侧上传文件 (所有文件支持多选)")
