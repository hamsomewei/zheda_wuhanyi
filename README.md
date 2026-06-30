# 单词造句卡片 Streamlit App

这个应用可以输入英语单词或短语，生成两种结果：

- 可下载的 Word 文件，格式接近 `flashcards_english.docx`
- 网页里的英文例句，目标词红色高亮，并可播放发音

## 本地运行

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## 输入格式

每行一个单词或短语：

```text
watch TV
read books
```

也可以手动给中文意思，适合没有 AI API key 时使用：

```text
watch TV | 看电视
read books | 看书
```

## 使用 DeepSeek

本应用默认选择 DeepSeek。你需要在 DeepSeek 平台创建一个 API key，并确保账户有可用余额。

本地临时使用时，在左侧填写：

```text
AI 引擎：DeepSeek
API key：sk-你的新 key
模型名：deepseek-chat
Base URL：https://api.deepseek.com
```

## 部署到 Streamlit Community Cloud

1. 把本文件夹上传到 GitHub 仓库。
2. 打开 Streamlit 的部署页面。
3. 选择仓库，主文件填 `app.py`。
4. 打开 App settings > Secrets，添加：

```toml
DEEPSEEK_API_KEY = "你的新 DeepSeek API key"
```

5. 保存后重新运行 app。

不要把真实 key 写进 `app.py`、`README.md`，也不要提交 `.streamlit/secrets.toml` 到 GitHub。

## 其他 AI 引擎

通义千问：

```toml
DASHSCOPE_API_KEY = "你的 API key"
```

Kimi：

```toml
MOONSHOT_API_KEY = "你的 API key"
```

也可以使用通用配置：

```toml
AI_API_KEY = "你的 API key"
AI_BASE_URL = "https://你的-openai-compatible-api/v1"
AI_MODEL = "你的模型名"
```

没有 API key 也能运行，但会使用简单模板句和少量内置中文意思。联网 AI 调用失败时，应用会自动退回离线模板。
