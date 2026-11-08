# Jira/Gira Scanner

## 📋 概述

**Jira/Gira Scanner** 是一个用于扫描 Jira/Gira 任务和匹配现有实现的系统，特别适用于 Data Warehouse 和 Reporting System 项目。

**功能:**
- 📂 Script 文件库存分析
- 🎯 Jira Task 自动扫描
- 🔗 Requirements-Implementation Mapping
- ⚠️ Gap Detection & Reporting
- 🌐 Web Dashboard & API

---

## 🚀 快速开始

### 1. 安装依赖
```bash
cd /home/user1/work/jirascanner
pip install -r requirements.txt
```

### 2. 配置环境变量
创建 `.env` 文件:
```env
# Jira/Gira Server URL
JIRA_SERVER_URL=https://your-jira-server.com
GIRA_SERVER_URL=https://your-gira-server.com

# API Token (使用 OAuth 2.0 或个人令牌)
JIRA_API_TOKEN=your_api_token_here
GIRA_API_TOKEN=your_api_token_here

# Project Configuration
JIRA_PROJECT_KEY=PROJ
JIRA_SPACE_KEY=YOUR_SPACE

# Scripts Path (可选)
SCRIPTS_PATH=/home/user1/work/jirascanner/scripts
```

### 3. 运行 Main Scanner
```bash
python3 main.py --server https://your-jira-server.com \
                --token your_api_token \
                --project PROJ \
                --scripts /path/to/scripts
```

### 4. 启动 Web Dashboard
```bash
python3 -m uvicorn api:app --host 0.0.0.0 --port 8000
```

访问：http://localhost:8000/dashboard/dashboard.html

---

## 📁 文件结构

```
/home/user1/work/jirascanner/
├── scanner/                 # 核心扫描模块
│   ├── __init__.py
│   ├── inventorizer.py      # Script 库存分析器
│   └── scanner.py          # Jira Task 扫描器
│
├── scripts/                 # Script 库存存放目录
│   ├── adhoc/              # Ad-hoc Reports Scripts
│   ├── reports/            # 常规 Report Scripts
│   └── ...
│
├── api/                     # API 端点
│   ├── __init__.py
│   ├── suggestion_api.py   # Suggestion 接口
│   └── main.py             # API 主入口
│
├── dashboard/               # Web Dashboard
│   └── dashboard.html      # 前端界面
│
├── inventory/               # 库存数据存放
│   ├── inventory.json      # Inventory 数据
│   └── suggestions.json    # Suggestions 数据
│
├── main.py                  # 脚本入口
├── requirements.txt         # Python 依赖
├── README.md               # 本文档
└── .env                    # 环境变量 (本地)
```

---

## 🔧 使用方法

### Scan 现有 Scripts

```bash
python3 scanner/inventorizer.py --path /home/user1/work/jirascanner/scripts/adhoc
```

### 查看 Jira Tasks

```bash
python3 scanner/scanner.py --server https://jira.example.com \
                           --token your_token \
                           --project PROJ
```

### 查找 Reference 的 Scripts

```bash
# 搜索包含特定 Jira 引用的 Script
grep -r "PROJ-345" /home/user1/work/jirascanner/scripts/ --include="*.py" --include="*.sh"
```

### 手动 Upload Requirement

```bash
curl -X POST "http://localhost:8000/api/suggestions" \
  -H "Content-Type: application/json" \
  -d '{
    "task_id": "PROJ-345",
    "title": "添加月度报告性能优化功能",
    "severity": "high",
    "description": "当前脚本响应时间为 450ms，需要优化到 100ms 以内"
  }'
```

---

## 🔌 API 端点

### GET endpoints

```
GET /api/health              # 健康检查
GET /api/tasks               # 获取所有 Tasks
GET /api/tasks/{key}         # 获取单个 Task
GET /api/scripts             # 获取所有 Scripts 库存
GET /api/gaps/{task_key}     # 获取特定 Task 的 Gap 分析
GET /api/match               # Check Requirements-Implementation Mapping
```

### POST endpoints

```
POST /api/tasks              # 创建新 Task
POST /api/suggestions        # 创建新 Suggestion
POST /api/scripts/analyze    # Request 分析 Script
POST /api/inventory/upload   # Upload 新 Inventory 记录
```

---

## 📊 输出格式

### Task Scan Output
```json
{
  "tasks": [
    {
      "key": "PROJ-345",
      "summary": "Monthly Sales Dashboard",
      "status": "in progress",
      "priority": "high",
      "requirements": [
        {
          "id": "REQ-001",
          "description": "Calculate regional revenue breakdown",
          "type": "functional"
        }
      ],
      "confluence_docs": ["PROJ-345-Technical"]
    }
  ],
  "total_count": "1"
}
```

### Script Inventory Output
```json
{
  "scripts": [
    {
      "path": "/scripts/report_monthly_sales.py",
      "name": "report_monthly_sales.py",
      "size": 4523,
      "lines": 120,
      "timestamp": "2024-01-15T10:30:00Z",
      "jira_refs": ["PROJ-345", "PROJ-346"],
      "kpi_tags": ["kpi", "dashboard", "revenue"],
      "database_connections": ["SALES_DB", "MARKETING_APP"],
      "complexity": 3
    }
  ]
}
```

---

## 🛠️ 自定义 & 扩展

### 添加新的 Script Parser

在 `scanner/inventorizer.py` 中添加新的解析器类:

```python
class CustomScriptParser:
    def parse(self, file_path: Path) -> Dict:
        # 解析自定义格式的 Script
        pass
```

### 添加新的 Requirement Type

在 `api/suggestion_api.py` 中添加新的分类:

```python
class Requirement(BaseModel):
    # 添加新的要求类型
    tech_debt: Optional[bool] = False  # 技术债务
    compliance: Optional[bool] = False  # 合规要求
```

### 创建 Custom Jira Query

在 `.env` 中添加自定义 JQL query:

```env
CUSTOM_JQL="project=PROJ AND status in (Open, In Progress) ORDER BY created DESC"
```

---

## 🐳 Docker 部署 (可选)

### 构建 Docker Image

```bash
docker build -t jira-scanner .
docker run -p 8000:8000 -v your_dir:/data jira-scanner
```

### Docker Compose

```yaml
version: '3.8'
services:
  scanner:
    image: jira-scanner
    ports:
      - "8000:8000"
    environment:
      - JIRA_API_TOKEN=${JIRA_TOKEN}
      - JIRA_SERVER_URL=${JIRA_URL}
    volumes:
      - ./scripts:/app/scripts
      - ./inventory:/app/inventory
```

---

## 🔐 安全注意事项

### API Token Security
```bash
# 不要将 token 提交到 git
# 使用 .gitignore
```

**添加 `.gitignore`**:
```gitignore
.env
.env.local
*.token
__pycache__/
*.pyc
inventory/inventory.json
inventory/suggestions.json
```

---

## 📚 相关工具

| 工具 | 用途 |
|------|------|
| [jira](https://pypi.org/project/jira/) | Jira API 客户端 |
| [FastAPI](https://fastapi.tiangolo.com/) | Web 应用框架 |
| [requests](https://pypi.org/project/requests/) | HTTP Client |
| [tqdm](https://pypi.org/project/tqdm/) | 进度条显示 |
| [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) | HTML 解析 (Confluence) |

---

## 📞 支持

如有问题，请联系支持团队。

MIT License - 可自由使用和修改

---

## 🔄 更新日志

### 1.0.0
- Initial Release
- Jira/Gira Task Scanner
- Script Inventory Analyzer
- Requirements Matching Engine
- Web Dashboard
- REST API