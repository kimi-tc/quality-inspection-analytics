# 飞书邮箱报表自动导入

这套方案用于在 Mac mini 上自动下载飞书邮箱报表附件，并导入看板。

推荐优先使用 **macOS 自带邮件.app 模式**：只要邮件.app 已经配置好飞书邮箱并能收信，脚本就从本机邮件里找附件，不需要在项目里保存邮箱专用密码。

也保留 **IMAP 模式** 作为备用方案。

## 0. 推荐：邮件.app 模式

先在 Mac mini 自带“邮件.app”里添加飞书邮箱，并确认能收到主题为：

```text
看板_基础数据
看板_人效
```

的报表邮件。

然后在项目根目录创建配置：

```bash
cd ~/quality-inspection-analytics
cp .mail-reports.env.example .mail-reports.env
```

邮件.app 模式只需要重点改这个路径：

```bash
MAIL_REPORT_DOWNLOAD_DIR=/Users/tangchao144522/Downloads/预质检每日数据
```

如果邮件.app 里邮箱账户或收件箱不是默认名称，再配置：

```bash
APPLE_MAILBOX_NAME=INBOX
APPLE_MAIL_ACCOUNT_NAME=
```

测试从邮件.app 下载附件：

```bash
MAIL_REPORT_DRY_RUN=1 npm run mailapp:fetch
```

下载并导入本地看板：

```bash
npm run mailapp:import
```

下载并同步本地 + ECS：

```bash
npm run mailapp:import:sync
```

首次运行时，macOS 可能会弹出权限提示，请允许终端/Node 控制“邮件.app”。

## 1. 本地配置

以下是 IMAP 模式配置；如果使用邮件.app 模式，可以不填写 `FEISHU_EMAIL_AUTH_CODE`。

在项目根目录创建本地配置文件：

```bash
cd ~/quality-inspection-analytics
cp .mail-reports.env.example .mail-reports.env
```

编辑 `.mail-reports.env`：

```bash
FEISHU_IMAP_HOST=imap.feishu.cn
FEISHU_IMAP_PORT=993
FEISHU_IMAP_SECURE=1
FEISHU_EMAIL=你的飞书邮箱
FEISHU_EMAIL_AUTH_CODE=你的第三方客户端专用密码
FEISHU_IMAP_MAILBOX=INBOX

MAIL_REPORT_QUALITY_SUBJECT=看板_基础数据
MAIL_REPORT_EFFICIENCY_SUBJECT=看板_人效
MAIL_REPORT_LOOKBACK_DAYS=3
MAIL_REPORT_DOWNLOAD_DIR=$HOME/Downloads/预质检每日数据
MAIL_REPORT_SYNC_TO_ECS=0
```

`.mail-reports.env` 已被 `.gitignore` 忽略，不会提交到 GitHub。

## 2. IMAP 模式：下载附件

只下载邮件附件，不导入：

```bash
npm run mail:fetch
```

附件会保存到：

```text
~/Downloads/预质检每日数据/quality
~/Downloads/预质检每日数据/efficiency
~/Downloads/预质检每日数据/inbox
```

脚本会跳过已下载附件，避免重复保存。

## 3. IMAP 模式：下载并导入

只导入 Mac mini 本地看板：

```bash
npm run mail:import
```

同时导入 Mac mini 和 ECS：

```bash
npm run mail:import:sync
```

## 4. 测试模式

不写入、不导入，只检查能否识别：

```bash
MAIL_REPORT_DRY_RUN=1 npm run mailapp:import
```

如果附件已经下载好，只想测试导入：

```bash
MAIL_REPORT_SKIP_FETCH=1 MAIL_REPORT_DRY_RUN=1 npm run mailapp:import
```

## 5. 支持的邮件主题

默认识别：

```text
看板_基础数据 -> 质量数据
看板_人效     -> 人效数据
```

如果后续主题变化，只需要修改 `.mail-reports.env` 中的：

```bash
MAIL_REPORT_QUALITY_SUBJECT=
MAIL_REPORT_EFFICIENCY_SUBJECT=
```
