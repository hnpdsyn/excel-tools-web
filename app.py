"""
Excel 工具集 - Web版 (Streamlit)
手机/电脑浏览器直接使用，无需安装
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import io
import tempfile
import os
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Excel 工具集", page_icon="📊", layout="centered")

# ===================== 样式 =====================
st.markdown("""
<style>
    .stApp { max-width: 800px; margin: 0 auto; }
    .main-title { font-size: 1.5rem; font-weight: bold; margin-bottom: 0.5rem; }
    .section-title { font-size: 1.1rem; font-weight: bold; margin: 1rem 0 0.5rem; }
    .info-box { background: #f0f2f6; border-radius: 8px; padding: 12px; margin: 8px 0; }
    .stButton button { width: 100%; }
    @media (max-width: 600px) {
        .main-title { font-size: 1.2rem; }
        .stTab { font-size: 0.8rem; }
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 Excel 工具集</div>', unsafe_allow_html=True)
st.caption("合并 / 拆分 / 清理 / 格式 / 转换 — 手机电脑都能用")

# ===================== 工具函数 =====================

@st.cache_data
def load_df(file):
    """读取上传文件"""
    ext = Path(file.name).suffix.lower()
    if ext == '.csv':
        return pd.read_csv(file, dtype=str)
    else:
        return pd.read_excel(file, dtype=str)

def to_excel(df, sheet_name="数据"):
    """DataFrame转Excel二进制"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

def save_workbook(wb):
    """保存workbook到二进制"""
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ===================== 侧边栏 =====================

tab = st.sidebar.radio("选择功能", [
    "📂 合并", "✂️ 拆分", "🧹 清理", "🎨 格式", "🔄 转换"
])

# ===================== 1. 合并 =====================

if tab == "📂 合并":
    st.markdown('<div class="section-title">合并Excel</div>', unsafe_allow_html=True)

    merge_mode = st.radio("合并方式", ["多文件按行合并", "单文件按Sheet合并"], horizontal=True)

    if merge_mode == "多文件按行合并":
        files = st.file_uploader("上传多个Excel文件", type=['xlsx', 'xls', 'csv'],
                                 accept_multiple_files=True, key="merge_files")
        if files and len(files) >= 2:
            if st.button("🚀 开始合并", use_container_width=True):
                with st.spinner("合并中..."):
                    all_dfs = []
                    for f in files:
                        df = load_df(f)
                        df['来源文件'] = f.name
                        all_dfs.append(df)
                        st.info(f"✅ {f.name}：{len(df)}行")
                    result = pd.concat(all_dfs, ignore_index=True)
                    st.success(f"🎉 合并完成！共 {len(files)} 个文件，{len(result)} 行")
                    st.dataframe(result.head(10), use_container_width=True)
                    st.download_button("📥 下载合并结果", to_excel(result),
                                       "merged_result.xlsx", use_container_width=True)
        else:
            st.info("请上传至少2个文件")

    else:
        file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'], key="merge_sheet_file")
        if file:
            wb = openpyxl.load_workbook(file)
            sheet_names = wb.sheetnames
            selected = st.multiselect("选择要合并的Sheet", sheet_names, default=sheet_names[:2])
            if selected and st.button("🚀 开始合并", use_container_width=True):
                with st.spinner("合并中..."):
                    all_dfs = []
                    for name in selected:
                        df = pd.read_excel(file, sheet_name=name, dtype=str)
                        df['来源Sheet'] = name
                        all_dfs.append(df)
                        st.info(f"✅ {name}：{len(df)}行")
                    result = pd.concat(all_dfs, ignore_index=True)
                    st.success(f"🎉 合并完成！共 {len(selected)} 个Sheet，{len(result)} 行")
                    st.dataframe(result.head(10), use_container_width=True)
                    st.download_button("📥 下载合并结果", to_excel(result),
                                       "merged_sheets.xlsx", use_container_width=True)

# ===================== 2. 拆分 =====================

elif tab == "✂️ 拆分":
    st.markdown('<div class="section-title">拆分Excel</div>', unsafe_allow_html=True)

    split_mode = st.radio("拆分方式", ["按列值拆分", "按行数均分"], horizontal=True)

    file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'], key="split_file")
    if file:
        df = load_df(file)

        if split_mode == "按列值拆分":
            col = st.selectbox("选择拆分依据的列", df.columns)
            if st.button("🚀 开始拆分", use_container_width=True):
                with st.spinner("拆分中..."):
                    groups = df.groupby(df[col].fillna('空值'))
                    st.success(f"🎉 将拆分为 {len(groups)} 个文件")
                    for name, group in groups:
                        safe = str(name).replace('/', '_')[:30]
                        st.download_button(
                            f"📥 {safe}（{len(group)}行）",
                            to_excel(group), f"{safe}.xlsx",
                            use_container_width=True
                        )

        else:
            rows = st.number_input("每份行数", min_value=100, max_value=100000, value=1000, step=500)
            if st.button("🚀 开始拆分", use_container_width=True):
                with st.spinner("拆分中..."):
                    total = len(df)
                    parts = (total + rows - 1) // rows
                    st.success(f"🎉 将拆分为 {parts} 个文件")
                    for i in range(0, total, rows):
                        chunk = df.iloc[i:i + rows]
                        part_num = i // rows + 1
                        st.download_button(
                            f"📥 第{part_num}部分（{len(chunk)}行）",
                            to_excel(chunk), f"拆分_{part_num}.xlsx",
                            use_container_width=True
                        )

# ===================== 3. 清理 =====================

elif tab == "🧹 清理":
    st.markdown('<div class="section-title">一键清理</div>', unsafe_allow_html=True)

    file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'], key="clean_file")
    if file:
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        col1, col2 = st.columns(2)
        with col1:
            clean_trim = st.checkbox("去除前后空格", True)
            clean_empty_row = st.checkbox("删除空行", True)
            clean_empty_col = st.checkbox("删除空列", True)
        with col2:
            clean_dup = st.checkbox("删除重复行", True)
            clean_special = st.checkbox("清除特殊字符", True)
        has_header = st.checkbox("第一行为表头（保留表头）", True)

        if st.button("🚀 开始清理", use_container_width=True):
            with st.spinner("清理中..."):
                logs = []

                if clean_trim:
                    count = 0
                    for row in ws.iter_rows():
                        for cell in row:
                            if isinstance(cell.value, str):
                                cell.value = cell.value.strip()
                                count += 1
                    logs.append(f"✅ 已清理 {count} 个单元格空格")

                if clean_empty_row:
                    to_del = []
                    for r in range(ws.max_row, 0, -1):
                        if all(ws.cell(r, c).value is None for c in range(1, ws.max_column + 1)):
                            to_del.append(r)
                    for r in to_del:
                        ws.delete_rows(r)
                    logs.append(f"✅ 已删除 {len(to_del)} 行空行")

                if clean_empty_col:
                    to_del = []
                    for c in range(ws.max_column, 0, -1):
                        if all(ws.cell(r, c).value is None for r in range(1, ws.max_row + 1)):
                            to_del.append(c)
                    for c in to_del:
                        ws.delete_cols(c)
                    logs.append(f"✅ 已删除 {len(to_del)} 列空列")

                if clean_dup:
                    rows_data = []
                    for r in range(1, ws.max_row + 1):
                        rows_data.append(tuple(ws.cell(r, c).value for c in range(1, ws.max_column + 1)))
                    seen = set()
                    to_del = []
                    start = 1 if has_header else 0
                    for idx, vals in enumerate(rows_data[start:], start):
                        if vals in seen:
                            to_del.append(idx + 1)
                        else:
                            seen.add(vals)
                    for r in reversed(to_del):
                        ws.delete_rows(r)
                    logs.append(f"✅ 已删除 {len(to_del)} 行重复数据")

                if clean_special:
                    count = 0
                    for row in ws.iter_rows():
                        for cell in row:
                            if isinstance(cell.value, str):
                                cleaned = re.sub(r'[^\x20-\x7E\u4e00-\u9fff\uFF00-\uFFEF]', '', cell.value)
                                if cleaned != cell.value:
                                    cell.value = cleaned
                                    count += 1
                    logs.append(f"✅ 已清理 {count} 个单元格特殊字符")

                for log in logs:
                    st.info(log)

                # 预览
                data = ws.iter_rows(values_only=True)
                headers = [next(data)] if has_header else []
                rows_data = list(data)
                preview_df = pd.DataFrame(rows_data, columns=headers[0] if headers else None)
                st.success(f"🎉 清理完成！当前数据：{len(rows_data)}行 × {ws.max_column}列")
                st.dataframe(preview_df.head(10), use_container_width=True)

                st.download_button("📥 下载清理结果",
                                   save_workbook(wb),
                                   Path(file.name).stem + "_清理.xlsx",
                                   use_container_width=True)

# ===================== 4. 格式 =====================

elif tab == "🎨 格式":
    st.markdown('<div class="section-title">格式统一</div>', unsafe_allow_html=True)

    file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'], key="fmt_file")
    if file:
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        col1, col2 = st.columns(2)
        with col1:
            font_name = st.selectbox("字体", ['微软雅黑', '宋体', '黑体', 'Arial', 'Calibri'], index=0)
            font_size = st.slider("字号", 8, 20, 11)
        with col2:
            align_h = st.selectbox("水平对齐", ['center', 'left', 'right', 'general'], index=0)
            align_v = st.selectbox("垂直对齐", ['center', 'top', 'bottom'], index=0)

        bold_header = st.checkbox("表头加粗", True)
        auto_width = st.checkbox("列宽自适应", True)
        add_border = st.checkbox("添加边框", True)

        if st.button("🎨 开始格式化", use_container_width=True):
            with st.spinner("格式化中..."):
                font = Font(name=font_name, size=font_size)
                header_font = Font(name=font_name, size=font_size, bold=True)
                align = Alignment(horizontal=align_h, vertical=align_v)
                side = Side(style='thin')
                border = Border(left=side, right=side, top=side, bottom=side)

                for r_idx, row in enumerate(ws.iter_rows(), 1):
                    for cell in row:
                        cell.font = header_font if (r_idx == 1 and bold_header) else font
                        cell.alignment = align
                        if add_border:
                            cell.border = border

                if auto_width:
                    for c in range(1, ws.max_column + 1):
                        letter = get_column_letter(c)
                        max_len = 0
                        for r in range(1, ws.max_row + 1):
                            val = ws.cell(r, c).value
                            if val:
                                s = str(val)
                                length = sum(2 if '\u4e00' <= ch <= '\u9fff' else 1 for ch in s)
                                max_len = max(max_len, length)
                        ws.column_dimensions[letter].width = min(max_len + 2, 50)

                st.success(f"🎉 格式化完成！")
                st.download_button("📥 下载格式化结果",
                                   save_workbook(wb),
                                   Path(file.name).stem + "_格式化.xlsx",
                                   use_container_width=True)

# ===================== 5. 转换 =====================

elif tab == "🔄 转换":
    st.markdown('<div class="section-title">格式转换</div>', unsafe_allow_html=True)

    conv_mode = st.radio("转换方向", ["Excel → CSV", "CSV → Excel"], horizontal=True)

    if conv_mode == "Excel → CSV":
        file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'], key="conv_xlsx")
        encoding = st.selectbox("编码", ['utf-8-sig', 'utf-8', 'gbk'], index=0)
        if file and st.button("🔄 开始转换", use_container_width=True):
            with st.spinner("转换中..."):
                df = load_df(file)
                csv_data = df.to_csv(index=False, encoding=encoding)
                st.success(f"✅ 转换完成！{len(df)}行 × {len(df.columns)}列")
                st.dataframe(df.head(10), use_container_width=True)
                st.download_button("📥 下载CSV", csv_data,
                                   Path(file.name).stem + ".csv",
                                   use_container_width=True)

    else:
        file = st.file_uploader("上传CSV文件", type=['csv'], key="conv_csv")
        encoding = st.selectbox("编码", ['utf-8', 'utf-8-sig', 'gbk', 'gb2312'], index=0)
        if file and st.button("🔄 开始转换", use_container_width=True):
            with st.spinner("转换中..."):
                df = pd.read_csv(file, dtype=str, encoding=encoding)
                st.success(f"✅ 转换完成！{len(df)}行 × {len(df.columns)}列")
                st.dataframe(df.head(10), use_container_width=True)
                st.download_button("📥 下载Excel", to_excel(df),
                                   Path(file.name).stem + ".xlsx",
                                   use_container_width=True)

# ===================== 底部 =====================

st.markdown("---")
st.caption("💡 上传文件处理后可下载结果 | 大文件建议用桌面版EXE（合并拆分工具 / 格式整理工具）")