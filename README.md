# 记工

记工日、加班和日薪，工资实时合计。纯前端、离线可用，数据只存在浏览器本地（localStorage）。

- 打开 `index.html` 即用（或部署到任意静态托管）
- 数据随浏览器保存，换设备/清缓存前请在 app 内「导出 → 导出数据」备份
- 这是通用版本，不包含任何个人记工数据

## 本地预览

```sh
python3 -m http.server 8000
# 打开 http://localhost:8000
```

## 部署

推到 GitHub 仓库后，在 Settings → Pages 里选「Deploy from a branch」→ `main` / `(root)` 即可。
