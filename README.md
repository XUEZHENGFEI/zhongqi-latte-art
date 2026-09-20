# 中秋限定拉花 · 作业看板

单文件静态页面 · 3 阶点击钻取（区域 → 门店 → 人员）。

## 数据口径

- **保留职级（4 个）**：门店副经理 / 咖啡师 / 门店经理 / 值班经理
- **通过状态**：已通过=pass；未通过 / 待重新提交=fail；其他=pending
- **通过率**：已通过 ÷ 四类职级总数（含待批/未提交）

## 区域分组

- **江苏+合肥**：NJ / SZ / WX / CZ / NT / YZ / KS / ZJG / HA / HF
- **上海**：SH / SHP / SHL
- **浙江**：HZ / NB / JH / SX / WZ / JX / TZ

## 文件说明

- `index.html` — 单文件 HTML，数据内嵌（877 条作业明细），可直接双击打开
- `build.py` — 数据生成脚本（读取 Excel → 输出 index.html）

## 本地重新生成

需要 `openpyxl`：

```bash
python build.py
```

源数据路径需修改脚本里的 `SRC` 常量。

## 在线访问

部署在 Cloudflare Pages：<https://zhongqi-latte-art.pages.dev>

（域名以 Cloudflare 实际分配为准）