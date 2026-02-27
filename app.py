import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime

# ==========================================
# ✅ 0. ReportLab 可选依赖（Streamlit Cloud 常见缺失）
# ==========================================
REPORTLAB_OK = True
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import mm
except Exception:
    REPORTLAB_OK = False

# ==========================================
# 1. 页面配置
# ==========================================
st.set_page_config(layout="wide", page_title="Coupang 智能补货 (最终版)")
st.title("📦 Coupang 智能补货 (定制导出版)")
st.markdown("### 核心逻辑：分表极简展示 + 统一视觉标准")

# ==========================================
# 2. 列号配置 (请确认 Excel 实际位置)
# ==========================================
IDX_M_CODE    = 0
IDX_M_SHOP    = 1
IDX_M_COL_E   = 4
IDX_M_COL_F   = 5
IDX_M_COST    = 6
IDX_M_ORANGE  = 3
IDX_M_INBOUND = 12
IDX_M_ACTIVE  = 13

IDX_7D_SKU    = 0
IDX_7D_QTY    = 8

IDX_INV_R_SKU = 2
IDX_INV_R_QTY = 7
IDX_INV_R_FEE = 17

IDX_INV_J_BAR = 2
IDX_INV_J_QTY = 10

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
    n = col_idx + 1
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

# ==========================================
# 4. PDF/HTML 工单生成
# ==========================================
def df_to_pdf_bytes(df: pd.DataFrame, title: str, subtitle_lines: list[str], group_by_col="产品编码") -> bytes:
    """只有 reportlab 存在时才会调用"""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=12*mm, rightMargin=12*mm,
        topMargin=12*mm, bottomMargin=12*mm
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("t", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=14, leading=16)
    meta_style = ParagraphStyle("m", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=11)

    elements = [Paragraph(title, title_style)]
    for line in subtitle_lines:
        elements.append(Paragraph(line, meta_style))
    elements.append(Spacer(1, 6))

    if df is None or df.empty:
        elements.append(Paragraph("（无数据）", meta_style))
        doc.build(elements)
        return buf.getvalue()

    df_work = df.copy()
    cols = list(df_work.columns)
    data = [cols] + df_work.fillna("").astype(str).values.tolist()

    page_w, _ = A4
    usable_w = page_w - doc.leftMargin - doc.rightMargin

    weights = []
    for c in cols:
        series = df_work[c].fillna("").astype(str)
        m = max(len(c), int(series.map(len).max() if len(series) else 0))
        weights.append(max(6, min(28, m)))
    total = sum(weights) if sum(weights) > 0 else 1
    col_widths = [usable_w * (w / total) for w in weights]

    tbl = Table(data, colWidths=col_widths, repeatRows=1)

    # 分组斑马纹：按产品编码
    if group_by_col in df_work.columns:
        key = df_work[group_by_col].astype(str).fillna("")
        flags = (key.ne(key.shift()).cumsum() % 2).astype(int).tolist()
    else:
        flags = [i % 2 for i in range(len(df_work))]

    ts = TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4d4d4d")),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#bfbfbf")),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ])
    for i, f in enumerate(flags, start=1):
        ts.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f2f2f2") if f == 1 else colors.white)

    tbl.setStyle(ts)
    elements.append(tbl)
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("负责人签收：____________________    日期：____________________", meta_style))
    elements.append(Paragraph("备注：__________________________________________________________________________________", meta_style))
    elements.append(Paragraph("      __________________________________________________________________________________", meta_style))
    doc.build(elements)
    return buf.getvalue()

def df_to_html_bytes(df: pd.DataFrame, title: str, subtitle_lines: list[str], group_by_col="产品编码") -> bytes:
    """reportlab 不存在时降级：生成灰白打印友好的 HTML 工单"""
    if df is None or df.empty:
        body = "<p>（无数据）</p>"
    else:
        df_work = df.copy().fillna("").astype(str)
        if group_by_col in df_work.columns:
            key = df_work[group_by_col].astype(str)
            flags = (key.ne(key.shift()).cumsum() % 2).astype(int)
        else:
            flags = pd.Series([i % 2 for i in range(len(df_work))])

        rows = []
        header = "".join([f"<th>{c}</th>" for c in df_work.columns])
        rows.append(f"<tr class='thead'>{header}</tr>")
        for i, (_, r) in enumerate(df_work.iterrows()):
            cls = "zebra" if int(flags.iloc[i]) == 1 else ""
            tds = "".join([f"<td>{r[c]}</td>" for c in df_work.columns])
            rows.append(f"<tr class='{cls}'>{tds}</tr>")
        body = f"<table>{''.join(rows)}</table>"

    html = f"""
    <html>
    <head>
    <meta charset="utf-8" />
    <style>
      body {{ font-family: Arial, Helvetica, sans-serif; margin: 18px; color:#111; }}
      h1 {{ font-size: 18px; margin: 0 0 6px 0; }}
      .meta {{ font-size: 12px; margin: 2px 0; }}
      table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 11px; }}
      th, td {{ border: 1px solid #bbb; padding: 4px 6px; vertical-align: middle; }}
      .thead th {{ background: #4d4d4d; color: #fff; font-weight: 700; }}
      tr.zebra td {{ background: #f2f2f2; }}
      .sign {{ margin-top: 14px; font-size: 12px; }}
      @media print {{
        body {{ margin: 10mm; }}
      }}
    </style>
    </head>
    <body>
      <h1>{title}</h1>
      {''.join([f"<div class='meta'>{line}</div>" for line in subtitle_lines])}
      {body}
      <div class="sign">负责人签收：____________________　日期：____________________</div>
      <div class="sign">备注：__________________________________________________________________________________</div>
      <div class="sign">　　　__________________________________________________________________________________</div>
    </body>
    </html>
    """
    return html.encode("utf-8")

# ==========================================
# 5. 侧边栏
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数设置")
    safety_weeks = st.number_input("安全周数 (倍数)", min_value=1, max_value=20, value=3, step=1)
    min_safety_qty = st.number_input(
        "最低库存基数 (保底)", min_value=0, max_value=100, value=5, step=1,
        help="仅对【在做】且【有入库码】的产品生效"
    )
    st.divider()
    orange_safety_weeks = st.number_input("🚚 橙火安全周数 (调拨预警)", min_value=1, max_value=10, value=2, step=1)
    st.divider()
    redundancy_weeks = st.number_input("⚠️ 库存冗余周数 (滞销标准)", min_value=4, max_value=52, value=8, step=1)
    st.divider()
    search_key = st.text_input("输入产品编码 (A列)", placeholder="输入后按回车查询，留空看全部")
    st.divider()
    st.info("📂 请上传文件 (保持Master顺序)")
    file_master = st.file_uploader("1. 基础信息表 (Master) *必传", type=['xlsx', 'csv'])
    files_sales = st.file_uploader("2. 销售表 (近7天) *多选", type=['xlsx', 'csv'], accept_multiple_files=True)
    files_inv_r = st.file_uploader("3. 橙火/火箭仓库存 *多选", type=['xlsx', 'csv'], accept_multiple_files=True)
    files_inv_j = st.file_uploader("4. 极风库存 *多选", type=['xlsx', 'csv'], accept_multiple_files=True)

# ==========================================
# 6. 主逻辑
# ==========================================
if file_master and files_sales and files_inv_r and files_inv_j:
    if st.button("🚀 生成定制报表", type="primary", use_container_width=True):
        with st.spinner("正在按指定列顺序匹配数据..."):
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

            # Sales
            s_list = [read_file(f) for f in files_sales]
            df_sales = pd.concat(s_list, ignore_index=True)
            df_sales['Key'] = clean_match_key(df_sales.iloc[:, IDX_7D_SKU])
            df_sales['Qty'] = clean_num(df_sales.iloc[:, IDX_7D_QTY])
            agg_sales = df_sales.groupby('Key')['Qty'].sum().reset_index()

            # Orange/Rocket
            r_list = [read_file(f) for f in files_inv_r]
            df_r = pd.concat(r_list, ignore_index=True)
            df_r['Key'] = clean_match_key(df_r.iloc[:, IDX_INV_R_SKU])
            df_r['Qty'] = clean_num(df_r.iloc[:, IDX_INV_R_QTY])
            try:
                df_r['Fee'] = clean_num(df_r.iloc[:, IDX_INV_R_FEE])
            except:
                df_r['Fee'] = 0
            agg_orange = df_r.groupby('Key')[['Qty', 'Fee']].sum().reset_index()

            # Jifeng
            j_list = [read_file(f) for f in files_inv_j]
            df_j = pd.concat(j_list, ignore_index=True)
            df_j['Key'] = clean_match_key(df_j.iloc[:, IDX_INV_J_BAR])
            df_j['Qty'] = clean_num(df_j.iloc[:, IDX_INV_J_QTY])
            agg_jifeng = df_j.groupby('Key')['Qty'].sum().reset_index()

            # Merge
            df_final = pd.merge(df_base, agg_sales, left_on='Orange_ID', right_on='Key', how='left')
            df_final.rename(columns={'Qty': 'Sales_7d'}, inplace=True)

            df_final = pd.merge(df_final, agg_orange, left_on='Orange_ID', right_on='Key', how='left', suffixes=('', '_R'))
            df_final.rename(columns={'Qty': 'Stock_Orange', 'Fee': 'Storage_Fee'}, inplace=True)

            df_final = pd.merge(df_final, agg_jifeng, left_on='Inbound_Code', right_on='Key', how='left', suffixes=('', '_J'))
            df_final.rename(columns={'Qty': 'Stock_Jifeng'}, inplace=True)

            # Calc
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

            # Output
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

            df_sheet1 = df_out_base.copy()
            df_sheet1.insert(1, '在做(Y)', df_final['Active'].values)

            # 展示
            if search_key:
                df_display = df_sheet1[df_sheet1['产品编码'].astype(str).str.contains(search_key, case=False, na=False)].copy()
            else:
                df_display = df_sheet1.copy()
            st.dataframe(df_display, use_container_width=True, height=600, hide_index=True)

            # Sheet2/3/4
            df_buy = df_out_base[df_out_base['建议采购数'] > 0].copy()
            df_trans = df_out_base[df_out_base['建议调拨数量'] > 0].copy()
            df_fee = df_out_base[df_out_base['本月仓储费(预警)'] > 0].copy()

            # ==========================
            # Excel 导出（Table + 按产品编码分组斑马纹 + 指定列左对齐）
            # ==========================
            out_excel_io = io.BytesIO()
            with pd.ExcelWriter(out_excel_io, engine='xlsxwriter') as writer:
                wb = writer.book
                fmt_center = wb.add_format({'align': 'center', 'valign': 'vcenter'})
                fmt_left = wb.add_format({'align': 'left', 'valign': 'vcenter'})
                fmt_zebra_center = wb.add_format({'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter'})
                fmt_zebra_left = wb.add_format({'bg_color': '#F2F2F2', 'align': 'left', 'valign': 'vcenter'})

                def apply_group_zebra(ws, df_write, left_align_cols):
                    nrows = len(df_write)
                    ncols = len(df_write.columns)
                    if nrows <= 0:
                        return

                    # helper gid
                    codes = df_write['产品编码'].astype(str).fillna('') if '产品编码' in df_write.columns else pd.Series([''] * nrows)
                    gid = (codes.ne(codes.shift()).cumsum() % 2).astype(int)

                    helper_col = ncols
                    ws.write(0, helper_col, "_gid")
                    for i in range(nrows):
                        ws.write(i + 1, helper_col, int(gid.iloc[i]), fmt_center)
                    ws.set_column(helper_col, helper_col, None, None, {'hidden': True})

                    helper_letter = col_to_excel(helper_col)
                    formula = f'=${helper_letter}2=1'
                    for c in range(ncols):
                        colname = df_write.columns[c]
                        zebra_fmt = fmt_zebra_left if colname in left_align_cols else fmt_zebra_center
                        ws.conditional_format(1, c, nrows, c, {'type': 'formula', 'criteria': formula, 'format': zebra_fmt})

                def build_table_sheet(sheet_name, df_curr, fixed_width_cols=None, fixed_width=26, hide_cols=None, left_align_cols=None):
                    hide_cols = hide_cols or []
                    fixed_width_cols = fixed_width_cols or []
                    left_align_cols = set(left_align_cols or [])

                    df_write = df_curr.copy()
                    if '产品编码' in df_write.columns and '店铺名称' in df_write.columns:
                        df_write = blank_repeat_like_merge(df_write, '产品编码', ['店铺名称', '产品编码'])

                    df_write.to_excel(writer, index=False, sheet_name=sheet_name)
                    ws = writer.sheets[sheet_name]
                    nrows, ncols = len(df_write), len(df_write.columns)

                    columns = [{'header': c} for c in df_write.columns]
                    # 关键：Table + banded_rows 关闭（不使用 Table 自带斑马纹）
                    ws.add_table(0, 0, max(nrows, 1), ncols - 1, {
                        'columns': columns,
                        'style': 'Table Style Medium 9',
                        'autofilter': True,
                        'banded_rows': False
                    })

                    widths = estimate_col_widths(df_write, fixed_col_names=fixed_width_cols, fixed_width=fixed_width)
                    for i, (col_name, w) in enumerate(zip(df_write.columns, widths)):
                        base_fmt = fmt_left if col_name in left_align_cols else fmt_center
                        ws.set_column(i, i, w, base_fmt)

                    # 按产品编码分组斑马纹
                    apply_group_zebra(ws, df_write, left_align_cols)

                    for c in hide_cols:
                        if 0 <= c < ncols:
                            ws.set_column(c, c, None, None, {'hidden': True})

                # ✅ 你要求：基础信息 & SKU名称 内容左对齐
                left_cols = ['基础信息', 'SKU名称']

                build_table_sheet('补货计算表', df_sheet1, fixed_width_cols=['基础信息'], fixed_width=26, left_align_cols=left_cols)
                build_table_sheet('采购单(找工厂)', df_buy, fixed_width_cols=['基础信息'], fixed_width=26,
                                  hide_cols=[14, 15, 16, 17, 18, 19], left_align_cols=left_cols)
                build_table_sheet('调拨单(发橙火)', df_trans, fixed_width_cols=['基础信息'], fixed_width=26,
                                  hide_cols=[12, 13, 14, 15, 16, 17, 19], left_align_cols=left_cols)
                build_table_sheet('库龄预警单(需重入库)', df_fee, fixed_width_cols=['基础信息'], fixed_width=26,
                                  hide_cols=[11, 12, 13, 14, 15, 16, 17, 18], left_align_cols=left_cols)

            excel_bytes = out_excel_io.getvalue()

            # ==========================
            # 工单（PDF优先，否则HTML）
            # ==========================
            today = datetime.now().strftime("%Y-%m-%d")

            pdf_buy_cols = [c for c in ['店铺名称','产品编码','SKU名称','采购单价','建议采购数','预计采购总额(RMB)','橙火ID','入库码'] if c in df_buy.columns]
            pdf_buy = df_buy[pdf_buy_cols].copy() if not df_buy.empty else df_buy.copy()

            pdf_trans_cols = [c for c in ['店铺名称','产品编码','SKU名称','建议调拨数量','橙火库存','7天销量','橙火ID','入库码'] if c in df_trans.columns]
            pdf_trans = df_trans[pdf_trans_cols].copy() if not df_trans.empty else df_trans.copy()

            fee_cols = [c for c in ['店铺名称','产品编码','SKU名称','本月仓储费(预警)','橙火库存','库存合计','冗余数量'] if c in df_fee.columns]
            pdf_fee_show = df_fee[fee_cols].copy() if (not df_fee.empty and fee_cols) else df_fee.copy()

            # bytes
            if REPORTLAB_OK:
                buy_bytes = df_to_pdf_bytes(
                    pdf_buy, "采购工单（找工厂）",
                    [f"日期：{today}", "负责人：____________________    联系方式：____________________",
                     "说明：按建议采购数生成，黑白打印后分发给采购负责人执行。"],
                    group_by_col="产品编码"
                )
                trans_bytes = df_to_pdf_bytes(
                    pdf_trans, "调拨工单（发橙火）",
                    [f"日期：{today}", "负责人：____________________    联系方式：____________________",
                     "说明：按建议调拨数量生成，供仓库/发货负责人执行。"],
                    group_by_col="产品编码"
                )
                fee_bytes = df_to_pdf_bytes(
                    pdf_fee_show, "库龄预警工单（需重入库）",
                    [f"日期：{today}", "负责人：____________________    联系方式：____________________",
                     "说明：按仓储费预警生成，供负责人核查并决定重入库/清理策略。"],
                    group_by_col="产品编码"
                )
                ext = "pdf"
            else:
                buy_bytes = df_to_html_bytes(
                    pdf_buy, "采购工单（找工厂）",
                    [f"日期：{today}", "负责人：____________________    联系方式：____________________",
                     "说明：按建议采购数生成，黑白打印后分发给采购负责人执行。"],
                    group_by_col="产品编码"
                )
                trans_bytes = df_to_html_bytes(
                    pdf_trans, "调拨工单（发橙火）",
                    [f"日期：{today}", "负责人：____________________    联系方式：____________________",
                     "说明：按建议调拨数量生成，供仓库/发货负责人执行。"],
                    group_by_col="产品编码"
                )
                fee_bytes = df_to_html_bytes(
                    pdf_fee_show, "库龄预警工单（需重入库）",
                    [f"日期：{today}", "负责人：____________________    联系方式：____________________",
                     "说明：按仓储费预警生成，供负责人核查并决定重入库/清理策略。"],
                    group_by_col="产品编码"
                )
                ext = "html"

            # ==========================
            # ZIP打包：Excel + 3个工单
            # ==========================
            zip_io = io.BytesIO()
            with zipfile.ZipFile(zip_io, "w", zipfile.ZIP_DEFLATED) as z:
                xlsx_name = f"Coupang_Restock_Full_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"
                z.writestr(xlsx_name, excel_bytes)
                z.writestr(f"工单_采购单_找工厂_{pd.Timestamp.now().strftime('%Y%m%d')}.{ext}", buy_bytes)
                z.writestr(f"工单_调拨单_发橙火_{pd.Timestamp.now().strftime('%Y%m%d')}.{ext}", trans_bytes)
                z.writestr(f"工单_库龄预警_需重入库_{pd.Timestamp.now().strftime('%Y%m%d')}.{ext}", fee_bytes)

            if not REPORTLAB_OK:
                st.warning(
                    "⚠️ 当前环境未安装 reportlab，已自动改为生成 HTML 工单（可直接黑白打印）。\n\n"
                    "如需 PDF：请在仓库根目录 requirements.txt 加一行：reportlab，然后重启部署。"
                )
            else:
                st.success("✅ 已生成：Excel + 3个PDF工单（已打包）")

            st.download_button(
                "📦 下载：Excel + 3个工单（ZIP）",
                data=zip_io.getvalue(),
                file_name=f"Coupang_Report_Package_{pd.Timestamp.now().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                type="primary"
            )
else:
    st.info("👈 请在左侧上传文件")
