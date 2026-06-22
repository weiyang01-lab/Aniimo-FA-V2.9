# 游戏项目损益测算台 V2.9

这是一个适合直接上传到 GitHub 并部署到 Render 的发布目录。

## 目录说明

- `index.html`
  - 线上默认入口，已指向当前 `V2.9`
- `app.js`
  - V2.9 前端逻辑
- `styles.css`
  - V2.9 样式文件
- `vendor.echarts.5.5.1.min.js`
  - 图表依赖
- `server.py`
  - Python 单服务入口，负责静态页面和 `/api/*` 接口
- `excel_bridge.py`
  - Excel 解析逻辑
- `template-layout.xlsx`
  - 模板布局配置
- `requirements.txt`
  - Python 依赖
- `render.yaml`
  - Render 部署配置
- `Procfile`
  - 兼容部分平台的启动配置

## 本地运行

```bash
pip install -r requirements.txt
python server.py
```

本地默认端口为 `8000`，打开 `http://127.0.0.1:8000/` 即可。

如果需要指定端口：

```bash
set PORT=4174
python server.py
```

## GitHub 使用方式

1. 新建一个 GitHub 仓库。
2. 将本目录全部文件上传到仓库根目录。
3. 确保仓库根目录保留 `render.yaml`、`server.py`、`requirements.txt`。

## Render 部署方式

1. 在 Render 新建 `Web Service`。
2. 连接 GitHub 仓库。
3. Root Directory 留空，直接指向仓库根目录。
4. Render 会自动识别 `render.yaml`，或手动使用以下配置：
   - Runtime: `Python`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python server.py`

## 说明

- 线上部署时，前端会优先访问同域 `/api`，适合 Render 单服务部署。
- 页面中的金额单位统一为 `USD`，缩写按美式记法展示，例如 `K / M / B`。
- 当前发布目录已去掉历史版本和本地临时文件，更适合作为对外仓库根目录。
