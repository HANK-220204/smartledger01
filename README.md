# SmartLedger 智慧账本

## 启动

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

浏览器打开 http://127.0.0.1:5000

## 部署到 Render

项目已包含 `render.yaml` 和生产启动配置。将项目上传到 GitHub 后，在 Render 中选择该仓库并创建 Web Service；Render 会自动执行 `pip install -r requirements.txt` 和 `gunicorn app:app`。`SMARTLEDGER_SECRET_KEY` 会由 Render 自动生成。正式使用前请配置持久化数据库，当前 SQLite 适合演示和小规模测试。

## 功能
- 手动记录收入与支出
- SQLite 本地数据存储
- CSV 账本导入
- 收支统计与图表可视化
- 财务规划建议
- 本地 AI 财务规划（Ollama，无需 API Key）
- 演示 AI 助手
- DeepSeek / OpenAI 兼容接口配置
- CSV、Excel、Word、PDF 账本文档导入（自动识别日期、金额、分类和收支类型）

CSV 示例：
```csv
kind,amount,category,note,tx_date
expense,35,餐饮,午餐,2026-09-23
income,500,兼职,项目收入,2026-09-23
```

文档导入支持 `.csv`、`.xlsx`、`.xls`、`.docx` 和文本型 `.pdf`。Word 和 PDF 会提取文档中的表格或文字内容进行识别，并根据备注自动归类为工资、餐饮、购物、交通等类型；扫描图片型 PDF 需要先进行 OCR，当前不会自动识别图片中的文字。

## 本地 AI 规划（可选）

如果希望使用免费的本地大模型而不配置任何云端 API Key，可以安装 [Ollama](https://ollama.com/download/windows)，启动后执行：

```bash
ollama pull qwen2.5:3b
ollama serve
```

在“模型设置”中选择“本地 AI / Ollama”。账本摘要会发送到本机的 Ollama，不会发送到 DeepSeek、豆包或 Kimi。若 Ollama 未安装或未运行，系统会自动使用基于真实账目、月份趋势、分类占比和收入稳定性的本地动态分析，不会退回固定模板。
