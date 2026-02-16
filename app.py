import streamlit as st
import pandas as pd
import io
import re

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 智能补货系统 (精简版)")
st.title("📦 Coupang 智能补货系统 (纯净版)")
st.markdown("### 核心逻辑：基于【近7天销量】预测安全库存，直接生成补货工单")

# ==========================================
# 2. 列号配置 (请务必根据您的新Excel核对列号!!!)
# ==========================================
# --- Master 基础表 ---
IDX_M_CODE   = 0    # A列: 内部编码 (关联键)
IDX_M_SHOP   = 1    # B列: 店铺名
IDX_M_SKU    = 3    # D列: SKU ID (关联键)
IDX_M_COST   = 6    # G列: 采购成本 (用于算补货金额)
IDX_M_PROFIT = 10   # K列: 单品毛利 (理论值)
IDX_M_BAR    = 12   # M列: 条码/自发货ID (关联Jifeng)

# --- 7天销量表 (新) ---
# 假设您导出的7天销量表格式。
IDX_7D_SKU   = 0    # 假设 A列是SKU/注册商品ID
IDX_7D_QTY   = 8    # 假设 I列是数量 (请根据实际情况修改!)

# --- 库存表 ---
IDX_INV_R_SKU  = 2  # 火箭仓表 SKU所在列
IDX_INV_R_QTY  = 7  # 火箭仓表 数量所在列

IDX_INV_J_BAR  = 2  # 极风表 条码所在列
IDX_INV_J_QTY  = 10 # 极风表 数量所在列

# ==========================================
# 3. 工具函数
# ==========================================
def clean_match_key(series):
    """清洗用于匹配的Key (SKU/编码/条码)"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    """清洗数值列"""
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def read_file(file):
    """通用读取函数"""
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file, dtype=str)
        else:
            return pd.read_excel(file, dtype=str, engine='openpyxl')
    except:
        return pd.read_csv(file, dtype=str, encoding='gbk')

# ==========================================
# 4. 侧边栏 & 参数设置
# ==========================================
with st.sidebar:
    st.header("⚙️ 补货参数设置")
    
    safety_days = st.number_input("🛡️ 安全库存天数", min_value=7, max_value=60, value=20, step=1, 
                                  help="系统将按照：(近7天销量 ÷ 7) × 此天数，来计算您需要的安全库存量。")
    
    st.divider()
    
    st.header("📂 数据上传")
    st.info("请按顺序上传以下表格：")
    
    file_master = st.file_uploader("1. 清洗后的综合管理表 (Master) *必传", type=['xlsx', 'csv'])
    file_sales_7d = st.file_uploader("2. 近7天销售数据表 *必传 (用于算周转)", type=['xlsx', 'csv'])
    files_inv_r = st.file_uploader("3. 火箭仓库存 (Rocket) *必传", type=['xlsx', 'csv'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("4. 极风/自发货库存 (Jifeng) *必传", type=['xlsx', 'csv'], accept_multiple_files=True)

# ==========================================
# 5. 主逻辑
# ==========================================
if file_master and file_sales_7d and files_inv_r and files_inv_j:
    if st.button("🚀 开始计算补货工单", type="primary", use_container_width=True):
        with st.spinner("正在分析库存周转..."):
            
            # ----------------------------------
            # A. 读取基础信息 (Master)
            # ----------------------------------
            df_m = read_file(file_master)
            df_base = pd.DataFrame()
            df_base['Code'] = clean_match_key(df_m.iloc[:, IDX_M_CODE])
            df_base['Shop'] = df_m.iloc[:, IDX_M_SHOP].astype(str)
            df_base['SKU_ID'] = clean_match_key(df_m.iloc[:, IDX_M_SKU])
            df_base['Barcode'] = clean_match_key(df_m.iloc[:, IDX_M_BAR])
            df_base['Cost'] = clean_num(df_m.iloc[:, IDX_M_COST]) 
            df_base['Unit_Profit'] = clean_num(df_m.iloc[:, IDX_M_PROFIT])
            
            # ----------------------------------
            # B. 读取近7天销量 (Velocity)
            # ----------------------------------
            df_7d = read_file(file_sales_7d)
            df_7d['Match_SKU'] = clean_match_key(df_7d.iloc[:, IDX_7D_SKU])
            df_7d['Qty_7Days'] = clean_num(df_7d.iloc[:, IDX_7D_QTY])
            sales_velocity = df_7d.groupby('Match_SKU')['Qty_7Days'].sum().reset_index()
            
            # ----------------------------------
            # C. 读取库存 (Stock)
            # ----------------------------------
            # 1. 火箭仓 (按SKU匹配)
            r_list = [read_file(f) for f in files_inv_r]
            df_r = pd.concat(r_list, ignore_index=True)
            df_r['Match_SKU'] = clean_match_key(df_r.iloc[:, IDX_INV_R_SKU])
            df_r['Rocket_Stock'] = clean_num(df_r.iloc[:, IDX_INV_R_QTY])
            inv_r_agg = df_r.groupby('Match_SKU')['Rocket_Stock'].sum().reset_index()
            
            # 2. 极风/自发 (按条码/Bar匹配)
            j_list = [read_file(f) for f in files_inv_j]
            df_j = pd.concat(j_list, ignore_index=True)
            df_j['Match_Bar'] = clean_match_key(df_j.iloc[:, IDX_INV_J_BAR])
            df_j['Jifeng_Stock'] = clean_num(df_j.iloc[:, IDX_INV_J_QTY])
            inv_j_agg = df_j.groupby('Match_Bar')['Jifeng_Stock'].sum().reset_index()

            # ----------------------------------
            # D. 数据合并与计算
            # ----------------------------------
            # 1. 合并7天销量
            df_final = pd.merge(df_base, sales_velocity, left_on='SKU_ID', right_on='Match_SKU', how='left')
            
            # 2. 合并库存
            df_final = pd.merge(df_final, inv_r_agg, left_on='SKU_ID', right_on='Match_SKU', how='left')
            df_final = pd.merge(df_final, inv_j_agg, left_on='Barcode', right_on='Match_Bar', how='left')
            
            # 填充0
            df_final['Qty_7Days'] = df_final['Qty_7Days'].fillna(0)
            df_final['Rocket_Stock'] = df_final['Rocket_Stock'].fillna(0)
            df_final['Jifeng_Stock'] = df_final['Jifeng_Stock'].fillna(0)
            
            # 3. 核心计算公式
            # 日均销量 = 7天销量 / 7
            df_final['Daily_Avg_Sales'] = df_final['Qty_7Days'] / 7
            
            # 安全库存线 = 日均 * 设置天数
            df_final['Safety_Line'] = df_final['Daily_Avg_Sales'] * safety_days
            
            # 总库存
            df_final['Total_Stock'] = df_final['Rocket_Stock'] + df_final['Jifeng_Stock']
            
            # 建议补货量 = 安全库存线 - 总库存 (小于0则为0)
            df_final['Restock_Qty'] = (df_final['Safety_Line'] - df_final['Total_Stock']).apply(lambda x: x if x > 0 else 0).astype(int)
            
            # 补货金额 = 建议补货量 * 成本
            df_final['Restock_Cost'] = df_final['Restock_Qty'] * df_final['Cost']
            
            # ----------------------------------
            # E. 展示与下载
            # ----------------------------------
            
            # 筛选出需要补货的
            df_restock = df_final[df_final['Restock_Qty'] > 0].sort_values(by='Restock_Cost', ascending=False)
            
            # 看板指标
            total_need_money = df_restock['Restock_Cost'].sum()
            total_skus = len(df_restock)
            total_qty_needed = df_restock['Restock_Qty'].sum()
            
            st.divider()
            st.subheader("📊 补货概览")
            c1, c2, c3 = st.columns(3)
            c1.metric("💰 预计补货总金额", f"₩ {total_need_money:,.0f}")
            c2.metric("📦 需补货SKU数", f"{total_skus} 个")
            c3.metric("🚛 总补货件数", f"{total_qty_needed:,.0f} 件")
            
            st.warning(f"当前计算基于：近7天日均销量 × {safety_days}天安全库存。")
            
            # 展示表格
            st.subheader("📋 建议补货清单 (Top 50)")
            
            # 格式化显示列 (移除了净利相关列)
            show_cols = ['Code', 'Shop', 'SKU_ID', 'Barcode', 'Qty_7Days', 'Rocket_Stock', 'Jifeng_Stock', 'Restock_Qty', 'Restock_Cost']
            df_show = df_restock[show_cols].head(50).copy()
            
            st.dataframe(
                df_show.style.background_gradient(subset=['Restock_Qty'], cmap='Reds')
                             .format({'Restock_Cost': '{:,.0f}', 'Qty_7Days': '{:.0f}'}),
                use_container_width=True
            )
            
            # ----------------------------------
            # F. 生成Excel下载
            # ----------------------------------
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Sheet 1: 补货工单
                df_order = df_restock[['Code', 'Shop', 'SKU_ID', 'Barcode', 'Cost', 'Restock_Qty', 'Restock_Cost']].copy()
                df_order.columns = ['内部编码', '店铺', 'SKU', '条码(自发ID)', '采购单价', '建议补货数', '预计金额']
                df_order.to_excel(writer, index=False, sheet_name='补货工单_发采购')
                
                # Sheet 2: 详细分析
                df_final_out = df_final[['Code', 'Shop', 'SKU_ID', 'Barcode', 'Cost', 'Unit_Profit', 
                                         'Qty_7Days', 'Daily_Avg_Sales', 'Safety_Line', 
                                         'Rocket_Stock', 'Jifeng_Stock', 'Total_Stock', 
                                         'Restock_Qty', 'Restock_Cost']]
                df_final_out.columns = ['内部编码', '店铺', 'SKU', '条码', '成本', '单品毛利', 
                                        '近7天销量', '日均销量', '安全库存线', 
                                        '火箭仓库存', '极风库存', '总库存', 
                                        '建议补货数', '补货金额']
                df_final_out.to_excel(writer, index=False, sheet_name='全量数据分析')
                
                # 设置格式
                wb = writer.book
                ws1 = writer.sheets['补货工单_发采购']
                fmt_header = wb.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                ws1.set_row(0, None, fmt_header)
                ws1.set_column('A:G', 15)
                
            st.download_button(
                label="📥 下载补货工单 (Excel)",
                data=output.getvalue(),
                file_name=f"Restock_Order_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.ms-excel",
                type="primary"
            )

else:
    st.info("👈 请在左侧侧边栏上传必要的文件以开始。")
