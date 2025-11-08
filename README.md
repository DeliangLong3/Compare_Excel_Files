# Excel 文件对比工具 (Streamlit + Kimi LLM)

这是一个使用 Streamlit 构建的图形用户界面 (GUI) 工具，用于比较两个或多个 Excel 文件中的工作表内容，并利用 Kimi LLM (Moonshot-Kimi-K2-Instruct 模型) 来智能分析和总结差异。

## 功能

*   **多文件对比**: 支持同时选择多个 Excel 文件进行两两对比。
*   **智能差异分析**: 利用 Kimi LLM 精确识别和总结不同工作表之间的差异，包括数值、文本、格式、增删等。
*   **可视化报告**: 将对比结果以清晰的 Markdown 表格形式呈现，并生成汇总的 Excel 报告。
*   **用户友好的界面**: 通过 Streamlit 提供直观的界面，方便用户选择文件、目录和输入 API 密钥。

## 安装指南

### 前提条件

*   Python 3.7 或更高版本
*   有效的 DashScope API 密钥 (用于调用 Kimi LLM)

### 安装步骤

1.  **克隆仓库**:
    ```bash
    git clone <您的GitHub仓库URL>
    cd <您的项目目录>
    ```

2.  **创建并激活虚拟环境 (推荐)**:
    ```bash
    python -m venv venv
    # Windows
    .\venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```

3.  **安装依赖**:
    在项目根目录下，运行以下命令安装所需的 Python 包。请确保您已将 `requirements.txt` 文件添加到项目根目录。
    ```bash
    pip install -r requirements.txt
    ```
    如果您还没有 `requirements.txt` 文件，可以手动创建并添加以下内容：
    ```txt
    streamlit==1.30.0
    pandas==2.1.4
    dashscope==1.14.0
    xlsxwriter==3.1.9
    ```

## 使用方法

1.  **运行 Streamlit 应用**:
    在项目根目录下，执行以下命令启动应用：
    ```bash
    streamlit run compare_source_files_streamlit.py
    ```

2.  **配置参数**:
    *   **源文件目录**: 选择包含您要对比的 Excel (`.xlsx`) 文件的文件夹。
    *   **输出目录**: 选择保存对比结果报告的文件夹。
    *   **Kimi API 密钥**: 输入您的 DashScope API 密钥。

3.  **开始对比**:
    点击“开始对比分析”按钮。脚本将查找指定目录下的所有 `.xlsx` 文件，进行两两对比，并将详细的差异报告保存在输出目录中，同时生成一个总览报告。

4.  **查看结果**:
    对比完成后，您可以在输出目录中找到每个文件对的详细对比 Excel 文件，以及一个包含所有对比结果摘要的总览 Excel 文件。您也可以直接在应用界面下载总览报告。

## Kimi LLM API 说明

*   本工具使用 [DashScope](https://help.aliyun.com/zh/dashscope/developer-reference/api-key) 提供的 Kimi LLM API 进行智能分析。
*   您需要一个有效的 DashScope API 密钥才能使用此功能。
*   API 调用会发送每个工作表的 JSON 内容给 Kimi 模型进行分析。请确保您的 API 密钥具有足够的额度。
*   模型名称: `Moonshot-Kimi-K2-Instruct`

## 注意事项

*   请确保输入的 Excel 文件是 `.xlsx` 格式。
*   脚本会自动忽略以 `~$` 开头的临时 Excel 文件。
*   如果两个文件的工作表内容完全相同，将不会调用 Kimi API，并会标记为“相同”。
*   API 调用可能需要一些时间，具体取决于文件大小和 Kimi 服务的响应速度。请耐心等待。
*   请妥善保管您的 Kimi API 密钥，不要将其直接提交到公共代码仓库。