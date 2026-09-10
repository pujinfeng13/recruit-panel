# 招聘流程阶段看板 · Recruitment Pipeline Panel

一个纯静态的招聘流程看板（KPI、漏斗、阶段 Tab、候选人管理），数据保存在浏览器 `localStorage`，无需任何后端。可直接托管到 GitHub Pages，别人打开链接即可查看。

## 目录结构

```
recruit-panel-site/
├── index.html          # 看板主页面（即 recruit-panel.html）
├── resume-parser.js    # 简历解析
└── lib/
    ├── jszip.min.js
    ├── pdf.min.js
    └── pdf.worker.min.js
```

## 本地预览

直接用浏览器打开 `index.html` 即可（双击或拖进浏览器）。数据存在当前浏览器的 `localStorage`，刷新不丢。

## 部署到 GitHub Pages

### 方式一：手动（推荐先看一遍）

1. 在 GitHub 上新建一个仓库（例如 `recruit-panel`）。
   - 仓库名随意；若想用 `https://<用户名>.github.io/` 作为根域名，则仓库必须命名为 `<用户名>.github.io`。
2. 把本目录内容推送到该仓库的 `main` 分支：
   ```bash
   cd recruit-panel-site
   git remote add origin https://github.com/<用户名>/<仓库名>.git
   git branch -M main
   git push -u origin main
   ```
3. 仓库 → **Settings → Pages → Build and deployment**：
   - Source 选 **Deploy from a branch**
   - Branch 选 **main** / **root**（根目录）
   - 保存后等待 1~2 分钟。
4. 访问 `https://<用户名>.github.io/<仓库名>/` 即可。

> 注意：GitHub Pages 使用相对路径加载 `lib/`，本项目已全部用相对路径，无需额外配置。

### 方式二：GitHub Actions（如需自动部署）

在仓库添加 `.github/workflows/pages.yml`：

```yaml
name: Deploy to Pages
on:
  push:
    branches: [main]
permissions:
  contents: read
  pages: write
  id-token: write
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/upload-pages-artifact@v3
        with:
          path: .
      - uses: actions/deploy-pages@v4
```

仓库 Settings → Pages → Source 改为 **GitHub Actions** 即可。

## 数据说明

- 所有候选人 / 阶段数据仅保存在**访问者各自浏览器**的 `localStorage`，不会上传到任何服务器。
- 因此 GitHub Pages 版本适合"查看 / 演示 / 个人使用"；若需要**多人共享同一份数据**，需另接后端或把数据同步到仓库（非本静态版能力范围）。
- 导出 / 备份：看板内一般提供导出功能（JSON / Excel），可本地保存后再导入。
