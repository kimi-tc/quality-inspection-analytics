# 数仓网页半自动取数说明

这套脚本用于解决“每天从数仓网页导出 Excel，再导入看板”的重复操作。

第一版采用“本地生成 SQL + 数仓网页粘贴运行 + 自动导入”的半自动流程：

1. 本地脚本自动生成质量 SQL 和人效 SQL，并替换好日期。
2. 你在数仓网页中粘贴 SQL、运行并导出 Excel。
3. 脚本捕获或识别下载文件。
4. 下载后自动解析 Excel，转换文本数字，并导入看板数据服务。

这样避免脚本强依赖数仓页面按钮和编辑器 DOM，稳定性会更高。

## 生成 SQL

默认生成“昨天”的质量 SQL 和人效 SQL：

```bash
cd ~/quality-inspection-analytics
export PATH="/Users/a144522/.nvm/versions/node/v25.9.0/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"
npm run sql:datadev
```

生成文件默认放在：

```text
generated-sql/
```

指定日期：

```bash
npm run sql:datadev -- --date 2026-06-08
```

指定日期范围：

```bash
npm run sql:datadev -- --start 2026-06-01 --end 2026-06-08
```

只生成质量 SQL：

```bash
npm run sql:datadev -- --type quality --date 2026-06-08
```

只生成人效 SQL：

```bash
npm run sql:datadev -- --type efficiency --date 2026-06-08
```

生成并复制质量 SQL 到剪贴板：

```bash
npm run sql:datadev -- --type quality --date 2026-06-08 --copy quality
```

生成并复制人效 SQL 到剪贴板：

```bash
npm run sql:datadev -- --type efficiency --date 2026-06-08 --copy efficiency
```

## 每日使用

推荐每日流程：

```bash
cd ~/quality-inspection-analytics
export PATH="/Users/a144522/.nvm/versions/node/v25.9.0/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"

# 1. 生成昨天的质量 SQL 和人效 SQL
npm run sql:datadev

# 2. 打开数仓网页，手动粘贴 generated-sql 中的 SQL，运行并导出
npm run datadev:auto
```

如果只想打开网页并等待导出：

```bash
cd ~/quality-inspection-analytics
export PATH="/Users/a144522/.nvm/versions/node/v25.9.0/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"
npm run datadev:auto
```

脚本会打开：

```text
http://datadev.aihuishou.com/workbench
```

如果页面要求登录，请手动登录。登录完成后，在网页里粘贴已生成的 SQL、运行并导出 Excel 即可。

默认下载目录：

```text
~/Downloads/预质检每日数据
```

导出完成后，脚本会自动导入最新下载的 Excel。

## 只导入最新 Excel

如果你已经手动下载好了文件，只想导入最新文件：

```bash
cd ~/quality-inspection-analytics
export PATH="/Users/a144522/.nvm/versions/node/v25.9.0/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"
npm run import:quality
```

如果要导入指定文件：

```bash
QUALITY_IMPORT_FILE="/Users/a144522/Downloads/预质检每日数据/0608.xlsx" npm run import:quality
```

## 预检模式

只解析文件、不写入看板：

```bash
QUALITY_IMPORT_DRY_RUN=1 npm run import:quality
```

## 可配置项

```bash
# 数仓地址
export DATADEV_URL="http://datadev.aihuishou.com/workbench"

# 下载目录
export QUALITY_DOWNLOAD_DIR="$HOME/Downloads/预质检每日数据"

# 看板数据服务地址
export QUALITY_API_BASE_URL="http://127.0.0.1:3000"

# 保留登录态的浏览器 profile
export DATADEV_BROWSER_PROFILE="$HOME/.quality-inspection-datadev-profile"

# 使用本机 Chrome。若没有 Chrome，可改为空并安装 Playwright Chromium。
export DATADEV_BROWSER_CHANNEL="chrome"

# 等待导出超时时间，默认 20 分钟
export DATADEV_WAIT_TIMEOUT_MS=1200000
```

## 后续升级为更自动

如果确认数仓页面中 SQL 编辑器、执行按钮、导出按钮的 CSS 选择器稳定，可以配置：

```bash
export DATADEV_SQL_FILE="/path/to/daily.sql"
export DATADEV_SQL_EDITOR_SELECTOR="..."
export DATADEV_RUN_BUTTON_SELECTOR="..."
export DATADEV_EXPORT_BUTTON_SELECTOR="..."
npm run datadev:auto
```

配置后脚本会尝试自动填 SQL、点击执行和导出。

当前不建议自动保存账号密码，因为公司网页可能有验证码、二次验证或风控策略，手动登录更安全也更稳定。
