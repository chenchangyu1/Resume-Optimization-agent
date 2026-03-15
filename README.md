# 智能简历优化 Agent（LangChain + GLM）

这是一个基于 LangChain 与 GLM API 的简历优化项目。

项目可以读取 Word 简历与岗位要求，先做结构化分析与匹配，再对简历文本进行定向优化，在尽量保持原始样式的前提下输出新的 DOCX 文件。

## 功能特性

- 读取并解析 Word 简历（.docx/.doc）
- 支持岗位要求来自文本文件（txt/md）或图片文件（png/jpg/jpeg/webp/bmp）
- 将 Word 内容按单元类型提取与存储（文本/图片/超链接）
- 更完整的 Word 解析：支持文本框、嵌套表格，以及可选页眉页脚
- 使用 docx2python 补充提取脚注、尾注、批注等信息，作为额外上下文
- 先将岗位需求结构化，再与简历证据点进行映射后重写
- 仅优化文本内容，尽量保留模板排版与样式
- 输出优化后的 Word 简历（.docx）

## 目录结构

```text
.
├─ main.py                 # 命令行入口
├─ web_app.py              # Web 页面入口
├─ agent_service.py        # 核心处理逻辑
├─ templates/              # 前端模板
├─ static/                 # 静态资源
├─ data/                   # 示例输入数据
├─ uploads/                # Web 上传目录（运行时）
└─ output/                 # 输出目录（运行时）
```

## 1. 环境准备

推荐 Python 版本：3.10 - 3.13（避免 3.14+）。

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
```

然后编辑 `.env`，配置 `GLM_API_KEY`。

如果账号容易触发限流（429/1302），建议增加以下参数：

```bash
GLM_WORKFLOW_MODE=lite
GLM_MIN_REQUEST_INTERVAL_SECONDS=3
GLM_RATE_LIMIT_BACKOFF_SECONDS=15
GLM_ENABLE_VERIFICATION=false
GLM_INCLUDE_HEADER_FOOTER=true
GLM_ENABLE_CACHE=true
GLM_CACHE_DIR=.cache/glm
```

参数说明：

- `GLM_WORKFLOW_MODE`：`lite` 调用次数更少，限流场景更稳；`full` 保留完整多阶段流程
- `GLM_MIN_REQUEST_INTERVAL_SECONDS`：单任务中两次 GLM 请求之间的最小间隔
- `GLM_RATE_LIMIT_BACKOFF_SECONDS`：遇到限流后重试的基础等待时间
- `GLM_ENABLE_VERIFICATION`：是否启用额外结果校验步骤，关闭可减少一次调用
- `GLM_INCLUDE_HEADER_FOOTER`：是否在提取与重写中包含页眉页脚
- `GLM_ENABLE_CACHE`：是否开启本地缓存，重复请求时可加速
- `GLM_CACHE_DIR`：缓存目录路径

## 2. 准备输入文件

- 简历文件：如 `data/resume.docx`（也支持 .doc，会尝试转换）
- 岗位要求（二选一）：
- 文本文件：如 `data/jd.txt`
- 图片文件：如 `data/jd.png`

也可直接通过命令行传入岗位文本。

## 3. 命令行运行

使用岗位文本文件：

```bash
python main.py --resume data/resume.docx --job data/jd.txt --out output/optimized_resume.docx
```

使用岗位图片：

```bash
python main.py --resume data/resume.docx --job data/jd.png --out output/optimized_resume.docx
```

直接传岗位文字：

```bash
python main.py --resume data/resume.docx --job-text "负责LLM应用开发，3年以上Python经验，熟悉RAG与LangChain" --out output/optimized_resume.docx
```

## 4. Web 页面运行

```bash
python web_app.py
```

浏览器访问：http://127.0.0.1:8000

Web 功能：

- 上传 Word 简历（.docx/.doc）
- 岗位要求支持文件上传（txt/md/图片）或直接粘贴文本
- 处理完成后下载优化后的 DOCX

## 5. 输出说明

- 优化后简历：由 `--out` 指定路径输出
- 单元快照：会额外保存 JSON（`word_units_*.json` 或 `*.units.json`），用于记录文本/图片/链接/样式信息

## 6. 注意事项

- 本项目通过 OpenAI 兼容方式调用 GLM 接口
- 不同账号可用模型名称可能不同，请按需调整 `.env` 中模型配置
- 默认流程为 `lite`，可降低请求次数与延迟
- 若出现 429 或 1302 限流，请调大请求间隔与回退时间
- Python 3.14+ 可能触发 LangChain 依赖链兼容问题；本项目已验证 Python 3.13 可稳定运行
