import streamlit as st
import pandas as pd
import io
import os
import glob
from zipfile import ZipFile, ZIP_DEFLATED

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
# 5. PDF 生成（修复 CJK 方框）
# ==========================================
def _find_cjk_font_path():
    """
    依次在常见路径/项目目录寻找可用的中日韩字体（ttf/otf）。
    找到就返回路径，否则返回 None
    """
    candidates = []

    # 1) 项目内 fonts 目录（推荐你放这里）
    candidates += glob.glob(os.path.join("fonts", "*.ttf"))
    candidates += glob.glob(os.path.join("fonts", "*.otf"))
    candidates += glob.glob(os.path.join(".", "*.ttf"))
    candidates += glob.glob(os.path.join(".", "*.otf"))

    # 2) Linux 常见字体目录（Streamlit Cloud 可能存在）
    linux_dirs = [
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        "/usr/share/fonts/truetype",
        "/usr/share/fonts/opentype",
    ]
    for d in linux_dirs:
        candidates += glob.glob(os.path.join(d, "**", "*.ttf"), recursive=True)
        candidates += glob.glob(os.path.join(d, "**", "*.otf"), recursive=True)

    # 优先级关键字：NotoCJK / Nanum / SourceHan
    priority_keys = [
        "NotoSansCJK", "NotoSansKR", "NotoSansSC", "NotoSansTC",
        "SourceHanSans", "NanumGothic", "AppleGothic", "MalgunGothic", "SimHei"
    ]

    def score(p):
        name = os.path.basename(p)
        s = 0
        for i, k in enumerate(priority_keys):
            if k.lower() in name.lower():
                s += (len(priority_keys) - i) * 10
        # 常规：Regular 优先
        if "regular" in name.lower():
            s += 3
        return s

    # 过滤掉明显无用字体（比如符号字体）
    filtered = [p for p in candidates if os.path.isfile(p)]
    if not filtered:
        return None

    filtered.sort(key=score, reverse=True)
    # 返回最高分
    return filtered[0] if score(filtered[0]) > 0 else filtered[0]

def generate_pdf_bytes(df: pd.DataFrame, title: str, subtitle: str = "") -> bytes:
    """
    生成 A4 灰白工单 PDF（黑白打印友好）。
    依赖 reportlab：requirements.txt 需要加入 reportlab
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except Exception as e:
        raise RuntimeError(
            "缺少 reportlab 依赖。请在 requirements.txt 增加一行：reportlab"
        ) from e

    font_path = _find_cjk_font_path()
    if not font_path:
        raise RuntimeError(
            "未找到可用的中/韩文字体，PDF 会变成方框。\n"
            "✅ 解决：在项目根目录创建 fonts/ 并放入 NotoSansCJK 或 NanumGothic 的 .ttf/.otf 文件。"
        )

    # 注册字体（名字固定）
    font_name = "CJKFont"
    try:
        pdfmetrics.registerFont(TTFont(font_name, font_path))
    except Exception as e:
        raise RuntimeError(
            f"找到字体但注册失败：{font_path}\n"
            "请优先放 .ttf（有些 .otf 在特定环境会失败）。"
        ) from e

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24
    )

    styles = getSampleStyleSheet()
    styles["Title"].fontName = font_name
    styles["Normal"].fontName = font_name

    story = []
    story.append(Paragraph(title, styles["Title"]))
    if subtitle:
        story.append(Spacer(1, 6))
        story.append(Paragraph(subtitle, styles["Normal"]))
    story.append(Spacer(1, 12))

    # 表格内容（限制列数避免过宽：这里保留你常用字段）
    df_use = df.copy()

    # 转字符串避免 None
    df_use = df_use.fillna("")
    data = [list(df_use.columns)] + df_use.astype(str).values.tolist()

    tbl = Table(data, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("FONTSIZE", (0, 1), (-1, -1), 9),

        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#444444")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),

        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#999999")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
    ]))

    story.append(tbl)
    doc.build(story)
    return buf.getvalue()

# ==========================================
# 6. 主逻辑
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

            # --- H. 搜索与看板（基于 df_sheet1） ---
            if search_key:
                df_display = df_sheet1[df_sheet1['产品编码'].astype(str).str.contains(search_key, case=False, na=False)].copy()
            else:
                df_display = df_sheet1.copy()

            first_in_group = df_display['产品编码'].ne(df_display['产品编码'].shift())
            df_display_vis = df_display.copy()
            for col in ['店铺名称', '产品编码']:
                df_display_vis.loc[~first_in_group, col] = ''

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
                df_display_vis.style
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

            # ==========================================
            # Excel 导出：保持你现有“超级表”方案，但修复 banded_rows 干扰斑马纹
            # ==========================================
            def col_to_excel(col_idx: int) -> str:
                n = col_idx + 1
                s = ""
                while n:
                    n, r = divmod(n - 1, 26)
                    s = chr(65 + r) + s
                return s

            def estimate_col_widths(df: pd.DataFrame, fixed_col_names=None, fixed_width=24,
                                    min_w=6, max_w=22, header_pad=2, cell_pad=2):
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

            out_io = io.BytesIO()
            with pd.ExcelWriter(out_io, engine='xlsxwriter') as writer:
                wb = writer.book

                fmt_center = wb.add_format({'align': 'center', 'valign': 'vcenter'})
                fmt_header = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                fmt_header_dark = wb.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                fmt_zebra = wb.add_format({'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter'})

                fmt_bold_col = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
                fmt_red_bold = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True, 'align': 'center', 'valign': 'vcenter'})
                fmt_red_norm = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': False, 'align': 'center', 'valign': 'vcenter'})
                fmt_orange_bold = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': True, 'align': 'center', 'valign': 'vcenter'})
                fmt_orange_norm = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': False, 'align': 'center', 'valign': 'vcenter'})
                fmt_blue = wb.add_format({'bg_color': '#C5D9F1', 'font_color': '#1F497D', 'bold': True, 'align': 'center', 'valign': 'vcenter'})
                fmt_purple = wb.add_format({'bg_color': '#E1BEE7', 'font_color': '#4A148C', 'bold': True, 'align': 'center', 'valign': 'vcenter'})

                fmt_merge = wb.add_format({'align': 'center', 'valign': 'vcenter'})
                fmt_merge_zebra = wb.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': '#F2F2F2'})

                def get_group_ids(df_curr: pd.DataFrame) -> pd.Series:
                    codes = df_curr['产品编码'].astype(str).fillna('')
                    return (codes != codes.shift()).cumsum() % 2

                def write_headers(ws, headers, dark_map):
                    for c, h in enumerate(headers):
                        ws.write(0, c, h, fmt_header)
                    for c_idx, txt in dark_map.items():
                        if 0 <= c_idx < len(headers):
                            ws.write(0, c_idx, txt, fmt_header_dark)

                def apply_width_and_center(ws, widths):
                    for i, w in enumerate(widths):
                        ws.set_column(i, i, w, fmt_center)

                def add_zebra_conditional(ws, nrows, ncols, helper_col_idx):
                    if nrows <= 0 or ncols <= 0:
                        return
                    helper_col_letter = col_to_excel(helper_col_idx)
                    formula = f'=${helper_col_letter}2=1'
                    ws.conditional_format(1, 0, nrows, ncols - 1, {'type': 'formula', 'criteria': formula, 'format': fmt_zebra})

                def merge_visual_columns(ws, df_curr, col_indices_to_merge, helper_gid: pd.Series):
                    codes = df_curr['产品编码'].astype(str).fillna('')
                    start = 0
                    for i in range(1, len(codes) + 1):
                        is_break = (i == len(codes)) or (codes.iloc[i] != codes.iloc[i - 1])
                        if is_break:
                            end = i - 1
                            if end > start:
                                gid = int(helper_gid.iloc[start])
                                merge_fmt = fmt_merge_zebra if gid == 1 else fmt_merge
                                r1 = start + 1
                                r2 = end + 1
                                for c_idx in col_indices_to_merge:
                                    ws.merge_range(r1, c_idx, r2, c_idx, df_curr.iloc[start, c_idx], merge_fmt)
                            start = i

                def apply_common_conditionals(ws, nrows, mapping):
                    if nrows <= 0:
                        return
                    for col, kind in mapping:
                        if kind == 'bold':
                            ws.conditional_format(1, col, nrows, col, {'type': 'formula', 'criteria': '=TRUE', 'format': fmt_bold_col})
                        elif kind == 'red_bold':
                            ws.conditional_format(1, col, nrows, col, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red_bold})
                        elif kind == 'red_norm':
                            ws.conditional_format(1, col, nrows, col, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_red_norm})
                        elif kind == 'orange_bold':
                            ws.conditional_format(1, col, nrows, col, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange_bold})
                        elif kind == 'orange_norm':
                            ws.conditional_format(1, col, nrows, col, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_orange_norm})
                        elif kind == 'blue':
                            ws.conditional_format(1, col, nrows, col, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_blue})
                        elif kind == 'purple':
                            ws.conditional_format(1, col, nrows, col, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_purple})

                def add_excel_table(ws, df_curr, dark_headers_map, nrows, ncols):
                    """
                    ✅ 关键修复：banded_rows=False 禁用超级表自带隔行底色
                    只保留：筛选按钮、表结构
                    """
                    if nrows <= 0 or ncols <= 0:
                        return

                    columns = []
                    headers = list(df_curr.columns)
                    for i, h in enumerate(headers):
                        col_def = {'header': h}
                        # 让关键列表头维持深色
                        if i in dark_headers_map:
                            col_def['header_format'] = fmt_header_dark
                        else:
                            col_def['header_format'] = fmt_header
                        columns.append(col_def)

                    ws.add_table(
                        0, 0, nrows, ncols - 1,
                        {
                            'columns': columns,
                            'style': None,          # 不使用内置样式（避免覆盖我们的 conditional_format）
                            'banded_rows': False,   # ✅ 禁用隔行灰（你要的修复点）
                            'autofilter': True
                        }
                    )

                def build_sheet(ws, df_curr, dark_headers_map, merge_cols, fixed_width_cols=None,
                                fixed_width=24, min_w=6, max_w=22, hide_cols=None, make_table=True):
                    hide_cols = hide_cols or []
                    headers = list(df_curr.columns)
                    nrows = len(df_curr)
                    ncols = len(headers)

                    widths = estimate_col_widths(
                        df_curr,
                        fixed_col_names=fixed_width_cols or [],
                        fixed_width=fixed_width,
                        min_w=min_w,
                        max_w=max_w
                    )
                    apply_width_and_center(ws, widths)

                    # 写表头（逐格写入，避免整行染蓝）
                    write_headers(ws, headers, dark_headers_map)

                    # zebra_id 辅助列（隐藏）
                    gid = get_group_ids(df_curr) if nrows > 0 else pd.Series(dtype=int)
                    helper_col_idx = ncols
                    ws.write(0, helper_col_idx, '_zebra', fmt_header)
                    for i in range(nrows):
                        ws.write(i + 1, helper_col_idx, int(gid.iloc[i]), fmt_center)
                    ws.set_column(helper_col_idx, helper_col_idx, None, None, {'hidden': True})

                    # 斑马纹（按“产品编码组”）
                    add_zebra_conditional(ws, nrows, ncols, helper_col_idx)

                    # 合并视觉列
                    if nrows > 0 and merge_cols:
                        merge_visual_columns(ws, df_curr, merge_cols, gid)

                    # 隐藏列
                    for c in hide_cols:
                        if 0 <= c < ncols:
                            ws.set_column(c, c, None, None, {'hidden': True})

                    # ✅ Excel 超级表（禁用 banded_rows，避免覆盖我们的斑马纹）
                    if make_table:
                        add_excel_table(ws, df_curr, dark_headers_map, nrows, ncols)

                    return nrows, ncols

                # =========================
                # Sheet1：补货计算表（含在做列）
                # =========================
                name1 = '补货计算表'
                df_sheet1.to_excel(writer, index=False, sheet_name=name1)
                ws1 = writer.sheets[name1]

                dark1 = {2: '产品编码', 4: 'SKU名称', 13: '建议采购数', 16: '冗余数量', 19: '建议调拨数量', 20: '本月仓储费(预警)'}
                nrows1, ncols1 = build_sheet(
                    ws1, df_sheet1, dark1,
                    merge_cols=[0, 2],
                    fixed_width_cols=['基础信息'],
                    fixed_width=26,
                    min_w=6, max_w=22,
                    make_table=True
                )

                apply_common_conditionals(ws1, nrows1, [
                    (2, 'bold'),
                    (4, 'bold'),
                    (13, 'red_bold'),
                    (14, 'red_norm'),
                    (16, 'orange_bold'),
                    (17, 'orange_norm'),
                    (19, 'blue'),
                    (20, 'purple'),
                ])

                # =========================
                # Sheet2：采购单(找工厂)
                # =========================
                df_buy = df_out_base[df_out_base['建议采购数'] > 0].copy()
                name2 = '采购单(找工厂)'
                df_buy.to_excel(writer, index=False, sheet_name=name2)
                ws2 = writer.sheets[name2]

                dark2 = {1: '产品编码', 3: 'SKU名称', 12: '建议采购数', 15: '冗余数量', 18: '建议调拨数量', 19: '本月仓储费(预警)'}
                nrows2, ncols2 = build_sheet(
                    ws2, df_buy, dark2,
                    merge_cols=[0, 1],
                    fixed_width_cols=['基础信息'],
                    fixed_width=26,
                    min_w=6, max_w=22,
                    hide_cols=[14, 15, 16, 17, 18, 19],
                    make_table=True
                )

                apply_common_conditionals(ws2, nrows2, [
                    (1, 'bold'),
                    (3, 'bold'),
                    (12, 'red_bold'),
                    (13, 'red_norm'),
                    (15, 'orange_bold'),
                    (16, 'orange_norm'),
                    (18, 'blue'),
                    (19, 'purple'),
                ])

                # =========================
                # Sheet3：调拨单(发橙火)
                # =========================
                df_trans = df_out_base[df_out_base['建议调拨数量'] > 0].copy()
                name3 = '调拨单(发橙火)'
                df_trans.to_excel(writer, index=False, sheet_name=name3)
                ws3 = writer.sheets[name3]

                nrows3, ncols3 = build_sheet(
                    ws3, df_trans, dark2,
                    merge_cols=[0, 1],
                    fixed_width_cols=['基础信息'],
                    fixed_width=26,
                    min_w=6, max_w=22,
                    hide_cols=[12, 13, 14, 15, 16, 17, 19],
                    make_table=True
                )

                apply_common_conditionals(ws3, nrows3, [
                    (1, 'bold'),
                    (3, 'bold'),
                    (12, 'red_bold'),
                    (13, 'red_norm'),
                    (15, 'orange_bold'),
                    (16, 'orange_norm'),
                    (18, 'blue'),
                    (19, 'purple'),
                ])

                # =========================
                # Sheet4：库龄预警单(需重入库)
                # =========================
                df_fee = df_out_base[df_out_base['本月仓储费(预警)'] > 0].copy()
                name4 = '库龄预警单(需重入库)'
                df_fee.to_excel(writer, index=False, sheet_name=name4)
                ws4 = writer.sheets[name4]

                nrows4, ncols4 = build_sheet(
                    ws4, df_fee, dark2,
                    merge_cols=[0, 1],
                    fixed_width_cols=['基础信息'],
                    fixed_width=26,
                    min_w=6, max_w=22,
                    hide_cols=[11, 12, 13, 14, 15, 16, 17, 18],
                    make_table=True
                )

                apply_common_conditionals(ws4, nrows4, [
                    (1, 'bold'),
                    (3, 'bold'),
                    (12, 'red_bold'),
                    (13, 'red_norm'),
                    (15, 'orange_bold'),
                    (16, 'orange_norm'),
                    (18, 'blue'),
                    (19, 'purple'),
                ])

            excel_bytes = out_io.getvalue()

            # ==========================================
            # PDF（3个工单）+ ZIP 打包下载
            # ==========================================
            pdf_errors = []
            pdf_buy = pdf_trans = pdf_fee = None

            # 你可以按需在这里定制 PDF 需要显示的列（尽量别太宽）
            def pick_pdf_cols(df, mode):
                if df.empty:
                    return df
                base_cols = ['店铺名称', '产品编码', 'SKU名称']
                if mode == "buy":
                    extra = ['建议采购数', '预计采购总额(RMB)', '7天销量', '库存合计']
                elif mode == "trans":
                    extra = ['建议调拨数量', '橙火库存', '极风库存', '7天销量']
                else:
                    extra = ['本月仓储费(预警)', '橙火库存', '7天销量', '库存合计']
                cols = [c for c in base_cols + extra if c in df.columns]
                return df[cols].copy()

            try:
                if not df_buy.empty:
                    pdf_buy = generate_pdf_bytes(
                        pick_pdf_cols(df_buy, "buy"),
                        title="采购工单（找工厂）",
                        subtitle=f"生成日期：{pd.Timestamp.now().strftime('%Y-%m-%d')}"
                    )
            except Exception as e:
                pdf_errors.append(f"采购工单PDF：{e}")

            try:
                if not df_trans.empty:
                    pdf_trans = generate_pdf_bytes(
                        pick_pdf_cols(df_trans, "trans"),
                        title="调拨工单（发橙火）",
                        subtitle=f"生成日期：{pd.Timestamp.now().strftime('%Y-%m-%d')}"
                    )
            except Exception as e:
                pdf_errors.append(f"调拨工单PDF：{e}")

            try:
                if not df_fee.empty:
                    pdf_fee = generate_pdf_bytes(
                        pick_pdf_cols(df_fee, "fee"),
                        title="库龄预警工单（需重入库）",
                        subtitle=f"生成日期：{pd.Timestamp.now().strftime('%Y-%m-%d')}"
                    )
            except Exception as e:
                pdf_errors.append(f"库龄预警PDF：{e}")

            if pdf_errors:
                st.error("PDF 生成失败原因（按你截图的方框问题已在代码中处理，但需要字体/依赖齐全）：\n\n- " + "\n- ".join(pdf_errors))

            zip_buf = io.BytesIO()
            with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
                xlsx_name = f"Coupang_Restock_Full_v18_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"
                zf.writestr(xlsx_name, excel_bytes)

                if pdf_buy:
                    zf.writestr(f"工单_采购单_{pd.Timestamp.now().strftime('%Y%m%d')}.pdf", pdf_buy)
                if pdf_trans:
                    zf.writestr(f"工单_调拨单_{pd.Timestamp.now().strftime('%Y%m%d')}.pdf", pdf_trans)
                if pdf_fee:
                    zf.writestr(f"工单_库龄预警_{pd.Timestamp.now().strftime('%Y%m%d')}.pdf", pdf_fee)

            st.download_button(
                "📥 下载最终 ZIP（Excel + 3个PDF工单）",
                data=zip_buf.getvalue(),
                file_name=f"Coupang_Restock_Pack_v18_{pd.Timestamp.now().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                type="primary"
            )

            st.caption("⚠️ 提醒：如需PDF正常显示中文/韩文，请在 requirements.txt 增加 reportlab，并建议在项目 fonts/ 放入 NotoSansCJK 或 NanumGothic 字体文件。")

else:
    st.info("👈 请在左侧上传文件")
