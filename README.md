# SKU Tools Site

本项目是一个基于 Flask 的网页工具站，当前包含两类工具：

- `父SKU清洗工具`
- `店铺链接分析工具`

支持上传 `.xlsx` / `.csv` 文件，并在网页中完成处理、预览和导出。

## 本地启动

```powershell
python app.py
```

默认访问地址：

```text
http://127.0.0.1:5055
```

局域网访问：

```text
http://你的局域网IP:5055
```

## 主要文件

- `app.py`：Flask 入口与路由
- `sku_splitter.py`：父SKU提取与 Excel 处理逻辑
- `analysis_engine.py`：链接分析逻辑
- `xlsx_toolkit.py`：Excel / CSV 读取与导出
- `templates/`：页面模板
- `static/`：前端脚本与样式

## Render 部署

项目已补齐 Render 部署所需文件：

- `requirements.txt`
- `render.yaml`

默认启动命令：

```text
gunicorn app:app
```

部署步骤：

1. 把当前项目上传到 GitHub 仓库。
2. 打开 Render。
3. 选择 `New +` -> `Blueprint`。
4. 连接你的 GitHub 仓库。
5. Render 会自动识别 `render.yaml` 并创建服务。
6. 部署完成后，Render 会生成一个公网网址。

## Notion 使用方式

Notion 不能直接运行 Flask / Python 网站，但可以把已经部署好的公网网址嵌入页面。

做法：

1. 先把本项目部署到 Render 或其他支持 Python 的平台。
2. 在 Notion 页面中输入 `/embed`。
3. 粘贴部署后的公网网址。

## 依赖

- Flask
- requests
- gunicorn
