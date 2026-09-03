# 教学实训智能体平台（软件杯参赛作品）

一个面向教学场景的智能实训系统：教师用 AI 出题、批改、生成学习计划，学生练题、答疑、看学情报告。系统基于 **Flask + SQLite + DeepSeek API** 构建，答疑模块支持 **RAG 检索增强**（回答依据课程资料、可溯源）。

## 功能

- **AI 出题**：按题型（选择 / 填空 / 简答 / 编程）自动生成题目，各题型使用差异化 Prompt 与 temperature 策略
- **智能批改**：AI 对主观题作答评分（0–100）并定位错误、给出修正建议
- **学情分析**：基于学生答题数据自动生成班级掌握情况分析与教学建议
- **学习计划**：结合学生知识掌握情况，AI 生成个性化学习路径与实训任务
- **学生答疑**：知识库问答
  - 普通问答：把选中的课程知识拼进 Prompt 回答
  - **RAG 检索增强答疑（`/student/ask_rag`）**：先检索知识文件中与问题最相关的片段，再让模型依据资料回答，末尾自动标注引用来源
- **实时随堂练习**：AI 实时出题、即时批改
- **知识文件管理**：上传 docx / pdf / txt 并自动解析入库，按课程组织
- **三角色体系**：学生 / 教师 / 管理员（注册、登录、选课、课程与用户管理）

## 技术栈

| 层 | 技术 |
|---|---|
| 后端 | Python · Flask · SQLite |
| LLM | DeepSeek API（`deepseek-chat`，Key 通过环境变量注入） |
| 文档解析 | python-docx · PyMuPDF（fitz） |
| RAG | sentence-transformers（`BAAI/bge-small-zh-v1.5`，本地向量化，零 API 费用）· FAISS 检索 |
| 前端 | Flask 模板 + 原生 HTML/JS（templates/） |

## 快速开始

```bash
# 1. 安装依赖（建议使用 venv）
pip install -r requirements.txt

# 2. 设置 DeepSeek API Key（不要写进代码！）
# Windows PowerShell:
#   $env:DEEPSEEK_API_KEY = "sk-xxxx"

# 3. 启动
python app.py
# 浏览器打开 http://127.0.0.1:5000
```

### 体验 RAG 答疑（可选增强）

```powershell
# 首次运行需下载向量模型（约 100MB，国内走镜像）
$env:HF_ENDPOINT = "https://hf-mirror.com"
$env:DEEPSEEK_API_KEY = "sk-xxxx"

# 命令行试一条
python rag_demo.py "什么是TensorFlow.js？请依据资料回答"

# 网页端：登录学生账号后访问 /student/ask_rag
```

RAG 的向量化在本地完成（免费）；只有"生成回答"调用 DeepSeek，单次问答成本约几分钱。索引缓存在 `rag_cache/`（已加入 .gitignore）。

## 目录结构

```
A7_practical_training/
├── app.py                 # Flask 主程序（含 /student/ask_rag 路由）
├── llm_api.py             # DeepSeek API 封装（多场景 Prompt）
├── rag_utils.py           # RAG：文档切分 → 向量化 → FAISS 检索 → 生成
├── rag_demo.py            # RAG 命令行演示
├── config.py              # 配置
├── knowledge_files/       # 上传的知识文件（docx/pdf/txt）
├── templates/             # 页面模板（学生/教师/管理端）
├── database.db            # SQLite 数据库（本地数据，不入库）
└── .gitignore             # 忽略 .venv/database.db/rag_cache 等
```

## 安全说明

- DeepSeek API Key **只通过环境变量注入**（`llm_api.py` 中不出现任何明文密钥）
- `database.db`（含用户与题库数据）、`.venv/`、`rag_cache/` 均已在 `.gitignore` 中排除，不会随仓库公开
- 首次本地部署小模型效果不佳后，选型迭代切换至 DeepSeek 云端 API（生成质量与稳定性显著提升）

## 致谢 / 说明

本作品为 2025 年中国软件杯竞赛参赛项目（大二暑期开发），用于教学智能场景的技术验证与演示。
