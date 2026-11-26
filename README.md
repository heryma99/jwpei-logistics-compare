# 比价易 — 物流报价比价工具

单文件 SPA（GitHub Pages）：https://heryma99.github.io/jwpei-logistics-compare/

## 这是什么

出口物流报价比价工具。覆盖：小包专线 / 商业快递 / toB 头程 / 海外仓 / 尾程 / 环比 / 配置中心 / 全渠道比价等模块。
价格数据来自邮箱报价表和本地报价文件，通过烘焙脚本生成 `rates.json`，再同步到 `index.html` 的内嵌数据。

## 架构特点与风险

**所有功能和数据都在一个 `index.html` 里**（约 5.4MB）。新功能以"追加 script / 手术式修改"方式加入。

这带来一个反复出现过的问题：**自动重建内嵌 RATES 时，如果基底文件缺了某个追加功能，缺失会被反复固化，且回归测试查不出来**。历史上级联筛选、快速算价等功能都因此丢失过。

已建立两道防线：
- `quote_pull/guard_predeploy.py features` — 功能指纹守卫（查函数定义 + 调用点）
- `quote_pull/auto_sync.py:rebuild_inline_index()` — 从本地权威模板 `index_master.html` 重建，不从线上继承基底

## 开发铁律

1. 不改用户价格/渠道/架构数据；结构变化必须停并报。
2. 方案 → UI → 用户确认 → 才动手。
3. 部署三关（快照 / 白名单 / Puppeteer 回归）必过；成功判定看线上 blob sha。
4. 新增追加式功能时，**必须同步更新 `FEATURE_MARKERS` 和 `index_master.html`**。
5. **带缓存兜底的 baker 必须在回退旧缓存时写 stale 标记并禁止自动部署**——绝不允许过期价冒充新价上线。

## 完整文档

见项目根目录 `接手必读.md`（含路径、凭据、命令、数据来源、功能清单）。
本地项目根：`D:/WB文件/2026-07-31-11-24-33/`
