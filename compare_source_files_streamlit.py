# -*- coding: utf-8 -*-
"""
Streamlit应用，用于比较多个Excel文件并生成差异报告。
"""

import streamlit as st
import pandas as pd
import os
import glob
import json
import time
from http import HTTPStatus
import dashscope
from datetime import datetime
import logging
# from tkinter import Tk, filedialog # 移除 tkinter 导入
from itertools import combinations # 用于生成文件对
from io import BytesIO # 导入BytesIO

# --- Kimi API 相关函数 (从 compare_source_files.py 迁移) ---

def get_comparison_from_kimi(file1_content, file2_content, file1_name, file2_name, sheet_name, api_key, retries=3, delay=5):
    """
    使用Moonshot-Kimi模型来比较两个DataFrame的内容并生成总结。
    """
    model_name = "Moonshot-Kimi-K2-Instruct"
    prompt = f"""
# 角色
你是一位精通数据比对的数据分析专家。

# 背景
我需要比较两个Excel文件（`{file1_name}` 和 `{file2_name}`）中，名为 '{sheet_name}' 的工作表。你需要帮我精确地识别并总结这两个数据版本之间的所有差异。

# 任务
你的任务是深入、细致地比较以下两个JSON格式的数据内容，它们分别来自两个Excel文件的 '{sheet_name}' 工作表。然后，以一个清晰、结构化的Markdown表格形式，总结出所有的不同之处。

# 输入数据
## 文件1: `{file1_name}` (工作表: {sheet_name})
```json
{file1_content}
```

## 文件2: `{file2_name}` (工作表: {sheet_name})
```json
{file2_content}
```

# 输出要求
1.  **进行思考** (但不要在最终输出中显示思考过程):
    *   首先，通览两个数据集，理解其整体结构。
    *   逐项对比，找出所有差异。差异可能包括但不限于：
        *   **数值或文本不同**: 同一位置的单元格内容不一致。
        *   **存在性差异**: 某处在一个文件中有数据，在另一个文件中为空。
        *   **格式不同**: 内容相似但表达方式或格式有别（例如，“N/A” vs “-”, “1,000” vs “1000”）。
        *   **行或列的增删**: 一个文件可能比另一个文件多或少几行或几列数据。
        *   **逻辑差异**: 例如，一个文件标记为“不适用”，另一个文件却有具体数值。

2.  **格式化输出**:
    *   你 **必须** 以一个Markdown表格来呈现比较结果。
    *   表格的 **表头必须是**：`| 项目 | 文件1：{file1_name} | 文件2：{file2_name} | 差异说明 |`
    *   在“项目”列中，清晰地描述差异所在的行、列或字段。
    *   在“差异说明”列中，简要解释差异的类型（例如，“数值不同”、“格式不一致”、“行被移除”等）。
    *   **如果两个文件的工作表内容完全没有差异**，请返回一个仅包含表头的空Markdown表格。
    *   **不要输出任何** 表格之外的文字、解释、总结、标题或代码块标记。你的输出必须从 `| 项目 |` 开始。

# 示例输出格式
请严格遵循以下格式。

| 项目 | 文件1：Report_v2.xlsx | 文件2：Report_v1.xlsx | 差异说明 |
|---|---|---|---|
| **第3行, '销售额'列** | 15,000 | 12,500 | 数值不同 |
| **第5行** | (此行为新增) | (此行不存在) | 文件1新增了一行数据 |
| **'备注'列** | 所有备注均为大写 | 所有备注均为小写 | 文本格式不同 |
"""
    messages = [{'role': 'user', 'content': prompt}]

    for attempt in range(retries):
        try:
            response = dashscope.Generation.call(
                model=model_name,
                messages=messages,
                api_key=api_key,
                result_format='message'
            )

            if response.status_code == HTTPStatus.OK:
                content = response.output.choices[0].message.content
                logging.info(f"Kimi对工作表 '{sheet_name}' 分析成功 (尝试 {attempt + 1}/{retries})。")
                return content
            else:
                error_msg = (f"Kimi API 调用失败 (尝试 {attempt + 1}/{retries}) for sheet '{sheet_name}'. "
                             f"状态码: {response.status_code}, 错误码: {response.code}, 错误信息: {response.message}")
                logging.error(error_msg)

        except Exception as e:
            error_msg = f"调用Kimi API时发生异常 (尝试 {attempt + 1}/{retries}) for sheet '{sheet_name}': {str(e)}"
            logging.error(error_msg)

        if attempt < retries - 1:
            logging.warning(f"将在 {delay} 秒后重试...")
            time.sleep(delay)

    logging.error(f"所有重试均失败，无法获取工作表 '{sheet_name}' 的比较结果。")
    return None


def convert_df_to_json_string(df, orient='records', indent=4):
    """将DataFrame转换为格式化的JSON字符串用于Prompt。"""
    return df.to_json(orient=orient, indent=indent, force_ascii=False)

# --- Streamlit UI 配置 ---
st.set_page_config(page_title="Excel 文件对比工具", page_icon="📊", layout="wide")
st.title("📊 Excel 文件对比工具")

# --- 日志配置 ---
log_expander = st.expander("查看日志", expanded=False)
log_container = log_expander.container()

class StreamlitLogHandler(logging.Handler):
    """将日志记录发送到Streamlit UI容器的日志处理器。"""
    def __init__(self, container):
        super().__init__()
        self.container = container

    def emit(self, record):
        """格式化并显示日志记录。"""
        msg = self.format(record)
        level = record.levelno
        if level >= logging.ERROR:
            self.container.error(msg)
        elif level >= logging.WARNING:
            self.container.warning(msg)
        else:
            self.container.info(msg)

def setup_logging(container):
    """配置根日志记录器以将日志重定向到Streamlit UI。"""
    logger = logging.getLogger()
    if not any(isinstance(h, StreamlitLogHandler) for h in logger.handlers):
        logger.setLevel(logging.INFO)
        handler = StreamlitLogHandler(container)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%H:%M:%S')
        handler.setFormatter(formatter)
        logger.addHandler(handler)

# --- 文件上传组件 ---
def handle_file_upload():
    """处理用户上传的文件，并返回文件列表。"""
    uploaded_files = st.file_uploader("请上传要对比的 Excel 文件 (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    
    if uploaded_files:
        # 将上传的文件保存到临时目录，以便后续处理
        # 注意：在 Streamlit Cloud 中，文件上传是临时的，通常保存在内存或临时存储中
        # 这里我们直接处理上传的文件对象
        return uploaded_files
    return []

# --- 初始化会话状态 ---
if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []
if 'output_dir' not in st.session_state:
    st.session_state['output_dir'] = ""
if 'api_key' not in st.session_state:
    st.session_state['api_key'] = ""
if 'comparison_results' not in st.session_state:
    st.session_state['comparison_results'] = None
if 'final_excel_path' not in st.session_state:
    st.session_state['final_excel_path'] = None

# --- 侧边栏配置 ---
with st.sidebar:
    st.header("⚙️ 配置选项")

    # 1. 文件上传
    st.subheader("1. 上传文件")
    uploaded_files = handle_file_upload()
    st.session_state['uploaded_files'] = uploaded_files # 保存上传的文件列表

    st.divider()

    # 2. API密钥输入
    st.subheader("2. 输入密钥")
    st.text_input("Kimi API 密钥", type="password", key='api_key', placeholder="请输入您的DashScope API密钥", help="此工具需要调用Kimi模型进行AI分析。")

    st.divider()

    st.subheader("操作")
    process_button = st.button("开始对比分析", type="primary", use_container_width=True)

# --- 文件比较核心逻辑 ---
def perform_comparison(uploaded_files, api_key):
    """
    处理上传的Excel文件，进行两两比较，并将所有结果整合到一个Excel文件的内存对象中。
    返回一个包含Excel文件数据的BytesIO对象。
    """
    if len(uploaded_files) < 2:
        logging.error("请上传至少两个 Excel 文件进行比较。")
        return None

    file_data = [{'name': f.name, 'file_obj': f} for f in uploaded_files]
    file_pairs = list(combinations(file_data, 2))
    logging.info(f"发现 {len(file_data)} 个 Excel 文件，将进行 {len(file_pairs)} 对两两比较。")

    # 创建一个内存中的Excel写入器
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        overview_data = [] # 用于总览表的数据

        for i, (file1_info, file2_info) in enumerate(file_pairs):
            file1_name, file2_name = file1_info['name'], file2_info['name']
            file1_obj, file2_obj = file1_info['file_obj'], file2_info['file_obj']
            
            # 创建一个对用户友好的工作表名称
            pair_sheet_name_base = f"{file1_name[:10]}_vs_{file2_name[:10]}"
            
            logging.info(f"\n--- 开始比较对 {i+1}/{len(file_pairs)}: {file1_name} vs {file2_name} ---")

            try:
                # 重置文件对象的读取指针并使用 ExcelFile 优化内存
                file1_obj.seek(0)
                file2_obj.seek(0)
                xls1 = pd.ExcelFile(file1_obj)
                xls2 = pd.ExcelFile(file2_obj)
                sheets1, sheets2 = set(xls1.sheet_names), set(xls2.sheet_names)

            except Exception as e:
                logging.error(f"打开 Excel 文件 '{file1_name}' 或 '{file2_name}' 时出错: {e}")
                overview_data.append({'文件1': file1_name, '文件2': file2_name, '状态': '打开错误', '说明': str(e)})
                continue

            common_sheets = sorted(list(sheets1.intersection(sheets2)))
            
            # 为每个文件对创建一个概览工作表
            pair_overview_data = {
                '状态': ['共有工作表', '仅在文件1中', '仅在文件2中'],
                '工作表名称': [", ".join(common_sheets), ", ".join(sorted(list(sheets1 - sheets2))), ", ".join(sorted(list(sheets2 - sheets1)))]
            }
            pair_overview_df = pd.DataFrame(pair_overview_data)
            pair_overview_df.to_excel(writer, sheet_name=f"概览_{pair_sheet_name_base[:20]}", index=False)

            if not common_sheets:
                logging.warning(f"文件 '{file1_name}' 和 '{file2_name}' 没有共同的工作表可供比较。")
                overview_data.append({'文件1': file1_name, '文件2': file2_name, '状态': '无共同工作表', '说明': '无共同工作表，跳过。'})
                continue

            logging.info(f"将比较共同的工作表: {', '.join(common_sheets)}")

            for sheet_name in common_sheets:
                logging.info(f"--- 正在处理工作表: {sheet_name} ---")
                try:
                    logging.info(f"正在从 '{file1_name}' 读取工作表 '{sheet_name}'...")
                    current_df1 = pd.read_excel(xls1, sheet_name=sheet_name)
                    logging.info(f"正在从 '{file2_name}' 读取工作表 '{sheet_name}'...")
                    current_df2 = pd.read_excel(xls2, sheet_name=sheet_name)
                except Exception as e:
                    logging.error(f"读取工作表 '{sheet_name}' 时出错: {e}")
                    pd.DataFrame({'错误': [f"读取工作表 '{sheet_name}' 时出错: {e}"]}).to_excel(writer, sheet_name=f"错误_{pair_sheet_name_base[:20]}", index=False)
                    continue

                # 定义当前比较的详细工作表名称
                details_sheet_name = f"差异_{pair_sheet_name_base[:15]}_{sheet_name[:10]}"

                if current_df1.equals(current_df2):
                    logging.info(f"工作表 '{sheet_name}' 内容完全相同，跳过API分析。")
                    details_df = pd.DataFrame([{'状态': '内容完全相同', '说明': f"工作表 '{sheet_name}' 在两个文件中的内容完全相同。"}])
                    details_df.to_excel(writer, sheet_name=details_sheet_name, index=False)
                    continue
                
                logging.info(f"工作表 '{sheet_name}' 内容存在差异，准备调用 Kimi API 进行分析。")
                comparison_result = get_comparison_from_kimi(
                    convert_df_to_json_string(current_df1),
                    convert_df_to_json_string(current_df2),
                    file1_name, file2_name, sheet_name, api_key
                )

                if comparison_result:
                    try:
                        table_str = comparison_result.strip()
                        lines = table_str.strip().split('\n')
                        if len(lines) > 1 and '|' in lines[0] and '---' in lines[1]:
                            header = [h.strip() for h in lines[0].strip().strip('|').split('|')]
                            data_rows = [ [p.strip() for p in line.strip().strip('|').split('|')] for line in lines[2:] if '|' in line]
                            details_df = pd.DataFrame(data_rows, columns=header)
                            if details_df.empty:
                                details_df.loc[0] = ["无程序化差异"] * len(header)
                                details_df.iloc[0, -1] = "Kimi报告了一个空表格，可能意味着内容虽不同但无显著结构性差异。"
                        else:
                             details_df = pd.DataFrame([{'说明': f"Kimi报告在 '{sheet_name}' 中未发现结构化差异。", '原始输出': table_str}])
                        
                        details_df.to_excel(writer, sheet_name=details_sheet_name, index=False)
                        
                        # 自动调整列宽
                        worksheet = writer.sheets[details_sheet_name]
                        for idx, col in enumerate(details_df):
                            series = details_df[col]
                            max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 2
                            worksheet.set_column(idx, idx, min(max_len, 50))

                        logging.info(f"已将 '{sheet_name}' 的详细差异对比结果写入到总报告中。")
                    except Exception as e:
                        logging.error(f"解析Kimi为工作表 '{sheet_name}' 返回的Markdown表格并保存时出错: {e}")
                        pd.DataFrame({'原始返回内容': [comparison_result]}).to_excel(writer, sheet_name=f"错误_{pair_sheet_name_base[:20]}", index=False)
                else:
                    logging.warning(f"未能从Kimi获取工作表 '{sheet_name}' 的比较结果。")
                    pd.DataFrame({'错误': [f"未能从Kimi获取 '{sheet_name}' 的工作流比较结果。"]}).to_excel(writer, sheet_name=f"错误_{pair_sheet_name_base[:20]}", index=False)
            
            overview_data.append({'文件1': file1_name, '文件2': file2_name, '状态': '已完成', '说明': f"详细比较结果已生成在Excel报告中。"})
            logging.info(f"--- 比较对 {file1_name} vs {file2_name} 完成 ---")

        # 最后写入总览表
        overall_overview_df = pd.DataFrame(overview_data)
        overall_overview_df.to_excel(writer, sheet_name='总览-所有比较对', index=False)
        worksheet = writer.sheets['总览-所有比较对']
        for idx, col in enumerate(overall_overview_df):
            series = overall_overview_df[col]
            max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 2
            worksheet.set_column(idx, idx, min(max_len, 60))
            
        logging.info("已生成总的概览表。")

    logging.info("\n所有比较完成！准备提供下载。")
    output_buffer.seek(0)
    return output_buffer


# --- 主界面 ---
setup_logging(log_container) # 配置日志处理器

if process_button:
    log_container.empty()
    st.session_state['comparison_results'] = None
    st.session_state['final_excel_path'] = None

    uploaded_files = st.session_state.get('uploaded_files', [])
    api_key = st.session_state.get('api_key')

    if not uploaded_files or len(uploaded_files) < 2:
        st.error("❌ 请先上传至少两个 Excel 文件。")
    elif not api_key or "sk-" not in api_key:
        st.error("❌ 请输入有效的 Kimi API 密钥。")
    else:
        dashscope.api_key = api_key
        logging.info("API密钥已设置。开始执行比较...")

        with st.spinner("🤖 AI正在进行文件两两对比分析，请稍候..."):
            final_report_buffer = perform_comparison(uploaded_files, api_key)

        if final_report_buffer:
            st.success("✅ 对比分析完成！请点击下方按钮下载总报告。")
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            final_filename = f"Overall_Comparison_{timestamp}.xlsx"
            
            st.download_button(
                label="📥 下载总报告 (Excel)",
                data=final_report_buffer,
                file_name=final_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("⚠️ 文件对比分析过程中发生错误，请检查上方日志获取详细信息。")

else:
    st.info("👋 欢迎使用！请在左侧上传 Excel 文件，输入 API 密钥，然后点击“开始对比分析”。")
