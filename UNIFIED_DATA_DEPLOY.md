# 三端统一数据服务配置

## 目标

将阿里云 ECS 作为唯一数据源。以后只需在任意一个看板地址导入一次数据，本地开发环境、Mac mini 内网地址和 ECS 外网地址刷新后都会读取同一份数据。

统一的数据包括：

- 质量底表
- 人效底表
- 属性项分类字典
- 数据导入记录

AI 模型密钥和 AI 服务仍由各环境自行配置，不会转发至 ECS。

## 一、ECS 配置

ECS 是主数据源，`.env` 中不要设置 `SHARED_DATA_API_BASE_URL`，或者将其留空：

```env
SHARED_DATA_API_BASE_URL=""
```

更新并重启：

```bash
cd /opt/quality-inspection-analytics
git pull --ff-only
npm install
npm run build
systemctl restart quality-inspection
```

检查 ECS 数据源模式：

```bash
curl -s http://127.0.0.1:3000/api/data-source
```

应返回 `"mode":"local"`。

## 二、Mac mini 配置

编辑 Mac mini 项目目录中的 `.env`：

```bash
cd ~/quality-inspection-analytics
nano .env
```

增加：

```env
SHARED_DATA_API_BASE_URL="http://39.107.221.251:3000"
SHARED_DATA_API_TIMEOUT_MS=120000
```

然后更新并重启：

```bash
./scripts/update-local-prod.sh
```

检查：

```bash
curl -s http://127.0.0.1:3000/api/data-source
```

应返回 `"mode":"remote"`，地址为 ECS。

## 三、本地开发环境配置

在本地项目目录创建或编辑 `.env`：

```bash
cd /Users/a144522/quality-inspection-analytics
nano .env
```

增加：

```env
SHARED_DATA_API_BASE_URL="http://39.107.221.251:3000"
SHARED_DATA_API_TIMEOUT_MS=120000
```

重启本地服务后生效。

## 四、使用方式

配置完成后，可以在三个地址中的任意一个进入“数据导入”模块并导入数据。导入请求最终都会写入 ECS：

- 本地：`http://127.0.0.1:3000`
- Mac mini：`http://10.181.20.208:3000`
- ECS：`http://39.107.221.251:3000`

其他页面刷新后即可读取最新数据，无需重复导入。

## 五、注意事项

- ECS 不可用时，本地和 Mac mini 将无法读取或导入共享数据。
- 不要在 ECS `.env` 中将 `SHARED_DATA_API_BASE_URL` 指向 ECS 自己，否则会形成循环转发。
- ECS 的 `/opt/quality-inspection-analytics/data` 需要定期备份。
