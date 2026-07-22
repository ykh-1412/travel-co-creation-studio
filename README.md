# 下扬州 · 团队旅行共创台

一个可部署在个人电脑或服务器上的团队旅行资料整理网站。朋友可以提交攻略、住宿、餐饮、密室、室内休闲或景点链接，也可以在没有链接时直接写下旅行诉求。后台调用 DeepSeek（也可切换豆包、Ollama 或演示模式），先识别大类、子分类和特征标签，再把专属字段写入 Excel，最后由网站展示候选和周末行程。

## 已实现

- 无账号，仅用团队共享密码访问
- 一次提交多个旅行链接，或直接写一段旅行想法
- 文字诉求会保留原文，并明确标记为“团队偏好”，不会冒充真实商户资料
- 明确记录“成功读取 / 读取受限 / 已整理 / 未整理”
- 保存已提取信息、缺失信息和失败原因
- 6 人、周五晚至周日、住宿两晚的扬州行程框架
- 住宿地址、房间床位、两晚总价、烧烤条件和预订状态
- 两顿早餐、6 人密室以及汗蒸、桑拿、洗浴等室内休闲候选
- 餐饮细分烧烤、火锅、早茶早餐、炒菜正餐等；密室细分恐怖程度、难度、规模与六人价格
- Excel 与网站双向同步，人工修改过的关键字段会在重新分析时保留
- `行程首页`、`候选决策台`、`投递汇总`、六张分类明细、`处理报告`、`预订清单`、`两日行程` 等 14 张工作表

当前攻略数据保存在 `data/store.json`，当前 Excel 在：

```text
outputs/019f7eda-a998-7660-a73d-e43b3af67965/扬州团队旅行攻略.xlsx
```

真实 API 密钥不会提交到 GitHub。

## 本地运行

需要 Node.js 22.13 或更新版本，推荐 Node.js 24。

```bash
cp .env.example .env
npm ci
npm run build
npm start
```

打开 `http://localhost:8787`。8787 是统一入口，后台会把网页请求转发到内部的 3000 端口。

开发模式：

```bash
npm run local
```

## 用 Docker 部署到服务器（推荐）

服务器安装 Git 和 Docker 后：

```bash
git clone https://github.com/ykh-1412/yangzhou-trip-studio.git
cd yangzhou-trip-studio
cp .env.example .env
```

编辑 `.env`，至少填写：

```dotenv
AI_PROVIDER=deepseek
DEEPSEEK_API_KEY=你的新密钥
PUBLIC_ACCESS_PASSWORD=你的团队密码
PUBLIC_REQUIRE_PASSWORD=true
```

然后启动：

```bash
docker compose up -d --build
```

网站入口为 `http://服务器IP:8787`。正式给朋友使用时，请在 8787 前配置 Nginx、Caddy 或 Cloudflare 的 HTTPS 域名；密码 Cookie 使用 `Secure`，公网部署应使用 HTTPS。

`data/` 和 `outputs/` 已映射到服务器目录，容器重建后链接资料和 Excel 仍会保留。

## 以后更新服务器代码

```bash
git pull
docker compose up -d --build
```

如果服务器上的 `data/store.json` 或 Excel 已产生新内容，建议更新前先下载 Excel，或对 `data/`、`outputs/` 做一次备份。

## 不使用 Docker

```bash
git clone https://github.com/ykh-1412/yangzhou-trip-studio.git
cd yangzhou-trip-studio
cp .env.example .env
npm ci
npm run build
npm start
```

可再用 `systemd`、PM2 或其他进程管理工具保持 `npm start` 常驻，并用反向代理提供 HTTPS。

## 模型配置

- `AI_PROVIDER=deepseek`：DeepSeek API
- `AI_PROVIDER=doubao`：豆包 / 火山方舟
- `AI_PROVIDER=ollama`：服务器或本机上的 Ollama
- `AI_PROVIDER=demo`：不调用外部模型，仅用于流程演示

所有 AI 结论仍需在预订前人工核实，网站不会自动付款或替团队下单。

## 安全说明

- `.env`、日志、依赖、构建产物和本机 `cloudflared` 二进制已排除在 Git 之外。
- 仓库只包含 `.env.example`，不会包含真实 DeepSeek Key。
- 服务器请自行设置新密码和新 API Key，不要把 `.env` 提交到 GitHub。
- GitHub 仓库默认建议使用 Private；确认不含私人数据后再考虑公开。
