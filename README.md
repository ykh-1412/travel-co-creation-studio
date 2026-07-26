# 出行共创台

一个可以复用到不同目的地的团队旅行资料整理模板。朋友可以提交攻略、住宿、餐饮、密室、室内休闲或景点链接，也可以在没有链接时直接写旅行诉求。后台调用 DeepSeek（也可切换豆包、Ollama 或演示模式），把内容整理成可比较的分类字段并写入 Excel，网站再读取 Excel 首页生成最终行程。

当前仓库保留了一份扬州周末旅行作为完整示例；应用品牌、AI 行程背景、地图城市、Excel 名称和新项目数据已经与“扬州”解耦。

## 已实现

- 无账号，仅用团队共享密码访问
- 一次提交多个旅行链接，或直接写一段旅行想法
- 一篇合集攻略可拆成多个具体商户或地点，同时保留原始来源
- 团队诉求、攻略资料、真实地点分开；只有真实地点可投票和加入行程，诉求可单独标记采纳
- 保留原文、AI 摘要、已提取信息、缺失信息和失败原因
- 住宿、餐饮、密室、室内休闲、景点与攻略分类整理
- Excel 与网站双向同步，人工修改过的关键字段不会被重新分析覆盖
- `候选决策台` 是核实状态、地点入选 / 需求采纳、人工结论和备注的唯一 Excel 修改入口，避免多张表互相覆盖
- `行程首页`、`候选决策台`、`团队需求`、`投递汇总`、六张分类明细、`处理报告`、`预订清单`、`两日行程` 等 15 张工作表
- 朋友在公网团队页投递和投票；主电脑才能删除、重分析、修改候选和定稿
- 后台只读取公网 `http/https` 链接，并拦截 `localhost`、局域网地址和跳转到内网的链接
- 一个仓库可切换多个独立旅行项目，不覆盖原来的数据

## 把它用作新旅行模板

GitHub 仓库启用 Template repository 后，可以点击 **Use this template** 创建自己的新仓库。

首次安装：

```bash
git clone https://github.com/ykh-1412/travel-co-creation-studio.git
cd travel-co-creation-studio
cp .env.example .env
npm ci
```

创建一个新的旅行项目：

```bash
npm run trip:new -- --destination=苏州 --slug=suzhou --people=6 --days=2 --nights=2
```

这条命令会：

1. 新建 `data/trips/suzhou.json`
2. 新建独立的 `outputs/suzhou/` Excel 目录
3. 更新本机 `.env`，让网站切换到苏州项目
4. 保留原来的扬州项目，不删除、不覆盖

重新启动网站后，新项目生效。以后切回旧项目，只需把 `.env` 中的 `TRIP_DATA_FILE`、`TRIP_OUTPUT_DIR` 和 `WORKBOOK_FILE_NAME` 改回对应值。

当前扬州示例数据位于 `data/store.json`，当前 Excel 默认生成在：

```text
outputs/019f7eda-a998-7660-a73d-e43b3af67965/出行共创项目.xlsx
```

## 本地运行

需要 Node.js 22.13 或更新版本，推荐 Node.js 24。

```bash
cp .env.example .env
npm ci
npm run build
npm start
```

主电脑打开 `http://localhost:8787`。这是本机管理入口，可修改候选、重新分析、读写 Excel 和定稿。

`http://localhost:8788` 是公网团队入口的本机预览，会要求输入共享密码，且不提供管理权限。3000 只是内部网页端口。

开发模式：

```bash
npm run local
```

## 当前扬州固定公网入口

团队成员可打开：

```text
https://yangzhou-trip.tail84dc10.ts.net/
```

- 朋友无需安装 Tailscale 或注册账号，只需输入团队共享密码
- 公网地址由本机 Tailscale Funnel 转发到无管理权限的团队端口 `127.0.0.1:8788`
- 公网团队页可投递、查看和投票；即使伪造 `Host: localhost` 也无法获得主电脑管理权限
- 网站能否访问取决于这台 Mac、Clash、网页服务、后台服务和 Tailscale 转发均保持运行
- 密码读取本机 `.env` 的 `PUBLIC_ACCESS_PASSWORD`，不要把真实密码提交到 GitHub
- 原来的 `trycloudflare.com` 是临时地址，已停止使用

本机维护信息：

```text
网页日志：logs/web.log、logs/web.error.log
后台日志：logs/api.log、logs/api.error.log
公网转发日志：logs/tailscaled.log、logs/tailscaled.error.log
```

Tailscale 后台服务开机自动启动；刚开机时通常需要约 10 秒完成联网。Clash 当前混合代理端口为 `7897`，如果以后修改 Clash 端口，也要同步修改本机 Tailscale LaunchAgent 中的代理地址。

## 用 Docker 部署到服务器

```bash
git clone https://github.com/ykh-1412/travel-co-creation-studio.git
cd travel-co-creation-studio
cp .env.example .env
```

编辑 `.env`，至少填写：

```dotenv
AI_PROVIDER=deepseek
DEEPSEEK_API_KEY=你的新密钥
PUBLIC_ACCESS_PASSWORD=你的团队密码
```

然后启动：

```bash
docker compose up -d --build
```

团队入口为 `http://服务器IP:8788`。正式给朋友使用时，请在 8788 前配置 HTTPS 域名。

管理端口 8787 只绑定服务器本机。如需远程管理，请先建立 SSH 端口转发，再在自己电脑打开 `http://localhost:8787`，不要把 8787 直接暴露到公网。

## 项目切换配置

```dotenv
TRIP_DATA_FILE=data/store.json
TRIP_OUTPUT_DIR=outputs/current
WORKBOOK_FILE_NAME=出行共创项目.xlsx
```

- `TRIP_DATA_FILE`：当前旅行的 JSON 数据
- `TRIP_OUTPUT_DIR`：当前旅行的 Excel 目录
- `WORKBOOK_FILE_NAME`：团队下载和编辑的 Excel 文件名

每个旅行使用独立数据文件和 Excel 目录，因此本机运行与服务器部署也可以各自选择不同项目。

## 模型配置

- `AI_PROVIDER=deepseek`：DeepSeek API
- `AI_PROVIDER=doubao`：豆包 / 火山方舟
- `AI_PROVIDER=ollama`：服务器或本机上的 Ollama
- `AI_PROVIDER=demo`：不调用外部模型，仅用于流程演示

所有 AI 结论仍需在预订前人工核实，网站不会自动付款或替团队下单。

## 安全说明

- `.env`、日志、依赖、构建产物和本机隧道程序已排除在 Git 之外
- 仓库只包含 `.env.example`，不会包含真实 DeepSeek Key
- 本机管理端口和公网团队端口使用两个独立监听器，权限不依赖可伪造的网址头
- 所有投票、投递、网页编辑和 Excel 写入会串行保存，避免多人同时操作相互覆盖
- 切换旅行不会删除旧项目；覆盖同名项目必须显式追加 `--force`
- GitHub 仓库建议保持 Private；确认不含私人数据后再考虑公开
