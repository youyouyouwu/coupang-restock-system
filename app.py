# ==========================================
# 3. 工具函数 (修复版)
# ==========================================
def clean_match_key(series):
    """清洗用于匹配的Key (SKU/编码/条码)"""
    return series.astype(str).str.replace(r'\.0$', '', regex=True).str.replace('"', '').str.strip().str.upper()

def clean_num(series):
    """清洗数值列"""
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)

def read_file(file):
    """
    通用读取函数 (增强版)
    自动尝试多种编码，解决 UnicodeDecodeError
    """
    # 1. 如果是 Excel 文件，直接用 openpyxl
    if file.name.endswith(('.xlsx', '.xls', '.xlsm')):
        try:
            file.seek(0)
            return pd.read_excel(file, dtype=str, engine='openpyxl')
        except Exception as e:
            st.error(f"Excel读取失败: {file.name}, 错误: {e}")
            return pd.DataFrame()

    # 2. 如果是 CSV，尝试多种编码轮询
    # Coupang 常用 cp949/euc-kr，Excel保存常用 utf-8-sig
    encodings_to_try = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'gbk', 'gb18030', 'latin1']
    
    for encoding in encodings_to_try:
        try:
            file.seek(0)  # !!! 关键：每次重试前必须把指针回到文件开头
            return pd.read_csv(file, dtype=str, encoding=encoding)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue  # 当前编码失败，尝试下一个
        except Exception as e:
            st.error(f"未知错误: {file.name}, {e}")
            return pd.DataFrame()
            
    # 3. 所有编码都失败
    st.error(f"❌ 无法读取文件: {file.name}。请尝试将文件另存为标准的 Excel (.xlsx) 格式再上传。")
    return pd.DataFrame()
