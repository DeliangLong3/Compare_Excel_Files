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

# --- 文件选择函数 (修改为使用Streamlit组件) ---
# def select_folder(key, label): # 移除旧函数
#     """打开文件夹选择器并更新session_state中的路径。"""
#     root = Tk()
#     root.withdraw()  # 隐藏主窗口
#     root.attributes('-topmost', True)  # 将对话框置于顶层
#     folder_path = filedialog.askdirectory(title=label)
#     root.destroy()
#     if folder_path:
#         st.session_state[key] = folder_path.replace("/", "\\") # 统一路径分隔符

# --- 初始化会话状态 ---
if 'input_path' not in st.session_state:
    st.session_state['input_path'] = ""
if 'output_path' not in st.session_state:
    st.session_state['output_path'] = ""
if 'api_key' not in st.session_state:
    st.session_state['api_key'] = ""
if 'comparison_results' not in st.session_state:
    st.session_state['comparison_results'] = None
if 'final_excel_path' not in st.session_state:
    st.session_state['final_excel_path'] = None

# --- 侧边栏配置 ---
with st.sidebar:
    st.header("⚙️ 配置选项")

    # 使用 st.text_input 来让用户输入目录路径
    st.text_input("1. 输入源文件目录", key='input_path', placeholder="包含Excel文件的文件夹路径")
    # 移除 tkinter 相关的按钮

    st.text_input("2. 输入输出目录", key='output_path', placeholder="保存对比结果的文件夹路径")
    # 移除 tkinter 相关的按钮

    st.divider()

    st.text_input("3. Kimi API 密钥", type="password", key='api_key', placeholder="请输入您的DashScope API密钥")

    st.divider()

    st.subheader("操作")
    process_button = st.button("开始对比分析", type="primary", use_container_width=True)

# --- 文件比较核心逻辑 ---
def perform_comparison(input_dir, output_dir, api_key):
    """
    查找输入目录下的所有Excel文件，进行两两比较，并将结果保存到输出目录。
    """
    excel_files = [f for f in glob.glob(os.path.join(input_dir, '*.xlsx')) if not os.path.basename(f).startswith('~$')]

    if len(excel_files) < 2:
        logging.error(f"在目录 '{input_dir}' 中需要至少2个 .xlsx 文件进行比较，但只找到 {len(excel_files)} 个。")
        return None

    # 生成所有文件对的组合
    file_pairs = list(combinations(excel_files, 2))

    all_comparison_outputs = {} # 存储所有比较结果的字典

    logging.info(f"发现 {len(excel_files)} 个Excel文件，将进行 {len(file_pairs)} 对两两比较。")

    # 创建一个总的ExcelWriter来写入所有结果
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    overall_output_filename = os.path.join(output_dir, f'Overall_Comparison_{timestamp}.xlsx')

    try:
        with pd.ExcelWriter(overall_output_filename, engine='xlsxwriter') as writer:
            # 写入一个概览表，列出所有比较对
            overview_data = []
            for i, (file1_path, file2_path) in enumerate(file_pairs):
                file1_name, file2_name = os.path.basename(file1_path), os.path.basename(file2_path)
                logging.info(f"\n--- 开始比较对 {i+1}/{len(file_pairs)}: {file1_name} vs {file2_name} ---")

                try:
                    xls1 = pd.ExcelFile(file1_path)
                    xls2 = pd.ExcelFile(file2_path)
                    sheets1 = set(xls1.sheet_names)
                    sheets2 = set(xls2.sheet_names)
                except Exception as e:
                    logging.error(f"读取Excel文件 '{file1_name}' 或 '{file2_name}' 时出错: {e}")
                    overview_data.append({'文件1': file1_name, '文件2': file2_name, '状态': '读取错误', '说明': str(e)})
                    continue

                common_sheets = sorted(list(sheets1.intersection(sheets2)))

                if not common_sheets:
                    logging.warning(f"文件 '{file1_name}' 和 '{file2_name}' 没有共同的工作表可供比较。")
                    overview_data.append({'文件1': file1_name, '文件2': file2_name, '状态': '无共同工作表', '说明': '两个文件没有共同的工作表可供比较。'})
                    continue

                logging.info(f"正在比较共同工作表: {', '.join(common_sheets)}")

                # 为当前比较对创建一个临时的ExcelWriter，用于写入其详细结果
                # 注意：这里我们不直接写入总的writer，而是先处理完一对，再将结果整合
                # 或者，我们可以为每一对创建一个单独的sheet，但文件名需要处理

                comparison_pair_output_filename = os.path.join(output_dir, f'Comparison_{file1_name}_vs_{file2_name}.xlsx')

                try:
                    with pd.ExcelWriter(comparison_pair_output_filename, engine='xlsxwriter') as pair_writer:
                        # 写入概览到当前比较对的Excel文件
                        overview_pair_data = {
                            '状态': ['共有工作表', '仅在文件1中', '仅在文件2中'],
                            '工作表名称': [", ".join(common_sheets), ", ".join(sorted(list(sheets1 - sheets2))), ", ".join(sorted(list(sheets2 - sheets1)))]
                        }
                        overview_pair_df = pd.DataFrame(overview_pair_data)
                        overview_pair_df.to_excel(pair_writer, sheet_name='概览', index=False)
                        logging.info(f"已生成 '{file1_name}_vs_{file2_name}' 的“概览”工作表。")

                        # 比较共有的工作表
                        for sheet_name in common_sheets:
                            logging.info(f"--- 正在处理工作表: {sheet_name} ---")
                            df1 = xls1.parse(sheet_name)
                            df2 = xls2.parse(sheet_name)

                            if df1.equals(df2):
                                logging.info(f"工作表 '{sheet_name}' 内容完全相同，跳过API分析。")
                                summary_text = f"工作表 '{sheet_name}' 在两个文件中的内容完全相同。"
                                df_details = pd.DataFrame([{'状态': '相同', '说明': summary_text}])

                                summary_df = pd.DataFrame({'总结': [summary_text]})
                                summary_df.to_excel(pair_writer, sheet_name=f"{sheet_name[:25]}_总结", index=False)
                                df_details.to_excel(pair_writer, sheet_name=f"{sheet_name[:25]}_差异", index=False)
                                continue

                            df1_content_str = convert_df_to_json_string(df1)
                            df2_content_str = convert_df_to_json_string(df2)

                            comparison_result = get_comparison_from_kimi(
                                df1_content_str, df2_content_str, file1_name, file2_name, sheet_name, api_key
                            )

                            if comparison_result:
                                try:
                                    table_str = comparison_result.strip()
                                    lines = table_str.strip().split('\n')

                                    if len(lines) > 1 and '|' in lines[0] and '---' in lines[1]:
                                        from io import StringIO
                                        header = [h.strip() for h in lines[0].strip().strip('|').split('|')]
                                        data_rows = []
                                        for line in lines[2:]:
                                            parts = [p.strip() for p in line.strip().strip('|').split('|')]
                                            if len(parts) == len(header):
                                                data_rows.append(parts)
                                        details_df = pd.DataFrame(data_rows, columns=header)
                                    elif '|' in lines[0]: # 可能是只有表头的空表格
                                        header = [h.strip() for h in lines[0].strip().strip('|').split('|')]
                                        details_df = pd.DataFrame(columns=header)
                                        if details_df.empty:
                                            details_df.loc[0] = ['无差异'] * len(header)
                                            details_df['差异说明'] = "Kimi报告在此工作表中未发现显著差异。"
                                    else:
                                        details_df = pd.DataFrame([{'说明': f"Kimi报告在工作表 '{sheet_name}' 中未发现差异或返回格式不正确。", '原始输出': table_str}])

                                    details_df.to_excel(pair_writer, sheet_name=f"{sheet_name[:25]}_差异对比", index=False)

                                    # 自动调整列宽
                                    worksheet = pair_writer.sheets[f"{sheet_name[:25]}_差异对比"]
                                    for idx, col in enumerate(details_df):
                                        series = details_df[col]
                                        max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 2
                                        worksheet.set_column(idx, idx, min(max_len, 50))

                                    logging.info(f"已将 '{sheet_name}' 的详细差异对比结果写入到输出文件中。")

                                except Exception as e:
                                    logging.error(f"解析Kimi为工作表 '{sheet_name}' 返回的Markdown表格并保存时出错: {e}")
                                    error_df = pd.DataFrame({'原始返回内容': [comparison_result]})
                                    error_df.to_excel(pair_writer, sheet_name=f"{sheet_name[:25]}_原始返回", index=False)
                            else:
                                logging.warning(f"未能从Kimi获取工作表 '{sheet_name}' 的比较结果。")
                                error_df = pd.DataFrame({'错误': [f"未能从Kimi获取 '{sheet_name}' 的工作流比较结果。"]})
                                error_df.to_excel(pair_writer, sheet_name=f"{sheet_name[:25]}_错误", index=False)

                    # 将当前比较对的结果添加到总概览中
                    overview_data.append({'文件1': file1_name, '文件2': file2_name, '状态': '已完成', '说明': f"比较结果已保存至: {os.path.basename(comparison_pair_output_filename)}"})
                    logging.info(f"--- 比较对 {file1_name} vs {file2_name} 完成 ---")

                except Exception as e:
                    logging.error(f"处理比较对 '{file1_name}' vs '{file2_name}' 时发生严重错误: {e}")
                    overview_data.append({'文件1': file1_name, '文件2': file2_name, '状态': '处理错误', '说明': str(e)})

            # 将总概览写入总的Excel文件
            overall_overview_df = pd.DataFrame(overview_data)
            overall_overview_df.to_excel(writer, sheet_name='总览', index=False)
            logging.info("已生成总的概览表。")

        logging.info(f"\n所有比较完成！详细结果已保存至: {overall_output_filename}")
        st.session_state['final_excel_path'] = overall_output_filename
        return overall_output_filename

    except Exception as e:
        logging.critical(f"生成总的Excel文件时发生严重错误: {e}", exc_info=True)
        st.error(f"生成总的Excel文件时发生严重错误: {e}")
        return None


# --- 主界面 ---
if __name__ == "__main__":
    setup_logging(log_container) # 配置日志处理器

    if process_button:
        log_container.empty()
        st.session_state['comparison_results'] = None
        st.session_state['final_excel_path'] = None

        input_dir = st.session_state.get('input_path')
        output_dir = st.session_state.get('output_path')
        api_key = st.session_state.get('api_key')

        if not input_dir or not os.path.isdir(input_dir):
            st.error("请先输入一个有效的源文件目录路径。")
        elif not output_dir or not os.path.isdir(output_dir): # 增加对输出目录的检查
            st.error("请先输入一个有效的输出目录路径。")
        elif not api_key or "sk-" not in api_key:
            st.error("请输入有效的 Kimi API 密钥。")
        else:
            # os.makedirs(output_dir, exist_ok=True) # 确保输出目录存在
            dashscope.api_key = api_key
            logging.info(f"API密钥已设置。源目录: {input_dir}, 输出目录: {output_dir}")

            with st.spinner("🤖 AI正在进行文件两两对比分析，请稍候..."):
                final_report_path = perform_comparison(input_dir, output_dir, api_key)

            if final_report_path:
                # 显示结果和下载链接
                st.success(f"对比分析完成！总报告已保存至: `{final_report_path}`")
                try:
                    with open(final_report_path, "rb") as f:
                        st.download_button(
                            label="📥 下载总报告Excel",
                            data=f,
                            file_name=os.path.basename(final_report_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except FileNotFoundError:
                    st.error(f"错误: 找不到生成的总报告文件以提供下载: {final_report_path}")
            else:
                st.error("文件对比分析过程中发生错误，请检查日志获取详细信息。")

    else:
        log_container.info("请在左侧配置源文件目录、输出目录和API密钥，然后点击“开始对比分析”。")
