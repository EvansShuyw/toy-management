# 货物报价管理系统

## 项目结构

```
├── frontend/          # Vue.js前端项目
└── backend/           # Python后端项目
    ├── main.py        # 主应用入口
    ├── models.py      # 数据模型定义
    ├── routers.py     # API路由
    ├── requirements.txt # 依赖包列表
    ├── uploads/       # 图片上传目录
    └── exports/       # Excel导出目录
```

## 技术栈

### 前端
- Vue.js - 前端框架
- Element Plus - UI组件库
- Axios - HTTP客户端
- vue-router - 路由管理
- pinia - 状态管理

### 后端
- Python FastAPI - Web框架
- SQLAlchemy - ORM框架
- SQLite - 开发环境数据库（可替换为MySQL等生产环境数据库）
- uvicorn - ASGI服务器
- Pillow - 图像处理
- openpyxl - Excel文件处理

## 功能特性

1. 货物报价表管理
   - 支持增删改查操作
   - 支持图片上传和预览
   - 字段包含：图片、货号、品名、装箱量PCS、箱规、毛重、净重

2. 数据导出
   - 支持勾选多条记录
   - 支持导出为Excel格式（.xlsx）
   - 导出的Excel包含所有数据字段和图片

## 安装与运行

### 前端
```bash
# 进入前端目录
cd frontend

# 安装依赖
npm install

# 启动开发服务器
npm run dev
```

### 后端
```bash
# 进入后端目录
cd backend

# 创建并激活虚拟环境
python -m venv venv
venv\Scripts\activate

# 安装依赖
venv\Scripts\python -m pip install -r requirements.txt

# 启动后端服务
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

## API接口

- `GET /items/` - 获取所有货物项目，支持按货号和品名筛选
- `POST /items/` - 创建新的货物项目
- `PUT /items/{item_id}` - 更新指定ID的货物项目
- `DELETE /items/{item_id}` - 删除指定ID的货物项目
- `POST /items/export` - 导出选中的货物项目为Excel文件

## 注意事项

- 后端默认使用SQLite数据库，数据存储在`toy_management.db`文件中
- 上传的图片存储在`uploads`目录
- 导出的Excel文件临时存储在`exports`目录，下载后会自动删除