# 阿里云 VPS 公网访问部署方案：数据保留在 Mac mini

这套方案的目标是：让老板可以通过公网域名访问看板，同时周度导入数据仍然只存放在 Mac mini 本地。

阿里云 VPS 只承担公网入口、HTTPS、访问鉴权和反向代理，不保存你导入的 Excel 明细或聚合数据。

## 架构说明

```text
老板浏览器
  -> https://你的域名
  -> 阿里云 VPS：Nginx + 访问密码
  -> VPS 本机 127.0.0.1:13000
  -> Mac mini 主动建立的反向 SSH 隧道
  -> Mac mini 本机 127.0.0.1:3000
  -> Mac mini 本地 data/shared-dataset.json
```

为什么推荐这个方案：

- 周度数据继续保存在 Mac mini 本地。
- VPS 不存储导入后的看板数据。
- 老板可以使用一个正常的公网域名访问。
- 不需要在公司或家里的路由器上做端口映射。
- Mac mini 主动连 VPS，通常更容易穿透局域网限制。

## 0. 准备工作

你需要先准备：

- 一台阿里云 ECS/VPS，并知道公网 IP。
- 一个域名或子域名，并将它解析到 VPS 公网 IP。
- Mac mini 可以 SSH 登录到 VPS。
- Mac mini 本地看板已经能正常访问：`http://127.0.0.1:3000`。

建议域名形式：

```text
dashboard.your-domain.com -> <ALIYUN_VPS_PUBLIC_IP>
```

阿里云安全组需要放行：

- TCP `22`：用于 Mac mini SSH 连接 VPS。
- TCP `80`：用于 HTTP 和申请 HTTPS 证书。
- TCP `443`：用于 HTTPS 公网访问。

## 1. 配置 VPS

先登录 VPS：

```bash
ssh root@<ALIYUN_VPS_PUBLIC_IP>
```

安装 Nginx、基础密码工具和证书工具。

如果 VPS 是 Ubuntu/Debian：

```bash
apt update
apt install -y nginx apache2-utils certbot python3-certbot-nginx
systemctl enable --now nginx
```

如果 VPS 是 CentOS/Alibaba Cloud Linux：

```bash
yum install -y nginx httpd-tools certbot python3-certbot-nginx
systemctl enable --now nginx
```

为看板创建访问账号。下面示例用户名是 `dashboard`，执行后会要求你输入访问密码：

```bash
htpasswd -c /etc/nginx/.weekly-inspection.htpasswd dashboard
```

创建 Nginx 配置文件：

```bash
nano /etc/nginx/conf.d/weekly-inspection.conf
```

将仓库里的这个模板内容复制进去：

```text
deploy/aliyun-nginx-weekly-inspection.conf
```

然后把模板里的：

```text
dashboard.example.com
```

替换成你的真实域名，例如：

```text
dashboard.your-domain.com
```

检查配置并重载 Nginx：

```bash
nginx -t
systemctl reload nginx
```

## 2. 开启 HTTPS

确认域名已经解析到 VPS 后，在 VPS 上执行：

```bash
certbot --nginx -d dashboard.your-domain.com
```

执行过程中如果提示是否将 HTTP 自动跳转到 HTTPS，建议选择跳转。

完成后，你的公网访问地址就是：

```text
https://dashboard.your-domain.com
```

## 3. 配置 Mac mini 到 VPS 的 SSH 密钥

在 Mac mini 上生成专用 SSH 密钥：

```bash
ssh-keygen -t ed25519 -C "weekly-inspection-tunnel" -f ~/.ssh/weekly_inspection_tunnel
```

把公钥复制到 VPS：

```bash
ssh-copy-id -i ~/.ssh/weekly_inspection_tunnel.pub root@<ALIYUN_VPS_PUBLIC_IP>
```

如果你的 Mac mini 没有 `ssh-copy-id`，可以先查看公钥：

```bash
cat ~/.ssh/weekly_inspection_tunnel.pub
```

然后把输出内容复制到 VPS 的这个文件中：

```bash
~/.ssh/authorized_keys
```

测试免密登录：

```bash
ssh -i ~/.ssh/weekly_inspection_tunnel root@<ALIYUN_VPS_PUBLIC_IP> "echo ok"
```

如果返回 `ok`，说明 SSH 密钥配置成功。

## 4. 在 Mac mini 上配置反向隧道

先更新 Mac mini 上的项目代码：

```bash
cd ~/quality-inspection-analytics
git pull
chmod +x scripts/start-aliyun-tunnel.sh scripts/check-aliyun-tunnel.sh
```

创建本地环境配置文件：

```bash
nano ~/.weekly-inspection-tunnel.env
```

写入以下内容，并替换成你的真实信息：

```bash
VPS_USER=root
VPS_HOST=<ALIYUN_VPS_PUBLIC_IP>
VPS_PORT=22
REMOTE_PORT=13000
LOCAL_PORT=3000
PUBLIC_URL=https://dashboard.your-domain.com
```

建议再配置 SSH alias，方便脚本稳定使用专用密钥：

```bash
nano ~/.ssh/config
```

追加以下内容：

```sshconfig
Host aliyun-weekly-inspection
  HostName <ALIYUN_VPS_PUBLIC_IP>
  User root
  IdentityFile ~/.ssh/weekly_inspection_tunnel
  ServerAliveInterval 30
  ServerAliveCountMax 3
```

如果使用了上面的 SSH alias，可以把 `~/.weekly-inspection-tunnel.env` 改成：

```bash
VPS_USER=root
VPS_HOST=aliyun-weekly-inspection
VPS_PORT=22
REMOTE_PORT=13000
LOCAL_PORT=3000
PUBLIC_URL=https://dashboard.your-domain.com
```

手动启动隧道测试：

```bash
cd ~/quality-inspection-analytics
./scripts/start-aliyun-tunnel.sh
```

这个命令会占用当前终端窗口。保持它运行，再新开一个终端检查连通性：

```bash
cd ~/quality-inspection-analytics
./scripts/check-aliyun-tunnel.sh
```

如果检查通过，就可以打开公网域名测试：

```text
https://dashboard.your-domain.com
```

## 5. 设置 Mac mini 开机自动连接隧道

确认手动测试没问题后，把隧道注册成 Mac mini 的 launchd 服务：

```bash
cd ~/quality-inspection-analytics
cp deploy/com.kimi.weekly-inspection-tunnel.plist ~/Library/LaunchAgents/
launchctl unload ~/Library/LaunchAgents/com.kimi.weekly-inspection-tunnel.plist 2>/dev/null || true
launchctl load ~/Library/LaunchAgents/com.kimi.weekly-inspection-tunnel.plist
launchctl list | grep weekly-inspection-tunnel
```

查看隧道日志：

```bash
tail -f ~/Library/Logs/weekly-inspection-tunnel.log
tail -f ~/Library/Logs/weekly-inspection-tunnel-error.log
```

如果服务断开，脚本会自动重连。

## 6. 日常使用

更新 Mac mini 上的看板代码：

```bash
cd ~/quality-inspection-analytics
./scripts/update-local-prod.sh
```

检查本地看板服务：

```bash
cd ~/quality-inspection-analytics
./scripts/check-local-prod.sh
```

检查公网隧道：

```bash
cd ~/quality-inspection-analytics
./scripts/check-aliyun-tunnel.sh
```

每周导入数据仍然在看板页面里操作。导入后的数据会写入 Mac mini 本地：

```text
~/quality-inspection-analytics/data/shared-dataset.json
```

## 7. 安全建议

公网访问一定要保留 Nginx 访问密码。当前看板页面包含导入和清空数据能力，如果公网无密码开放，会有误操作和数据风险。

建议：

- 必须启用 HTTPS。
- Nginx Basic Auth 密码设置得复杂一些。
- 如果可以，阿里云安全组的 SSH 端口只允许你的办公网络或 Mac mini 所在网络访问。
- 定期备份 `data/shared-dataset.json`。
- 不要把 `~/.weekly-inspection-tunnel.env`、SSH 私钥或任何密码提交到 GitHub。

## 8. 常见问题排查

如果公网打开显示 `502 Bad Gateway`，先在 VPS 上检查反向隧道端口：

```bash
curl -I http://127.0.0.1:13000
```

如果这一步失败，说明 Mac mini 到 VPS 的反向隧道没有连上。

再回到 Mac mini 检查：

```bash
launchctl list | grep weekly-inspection-tunnel
tail -n 80 ~/Library/Logs/weekly-inspection-tunnel-error.log
```

如果公网页面一直要求用户名和密码，或者密码不对，在 VPS 上重置密码：

```bash
htpasswd /etc/nginx/.weekly-inspection.htpasswd dashboard
nginx -t
systemctl reload nginx
```

如果公网页面能打开，但页面提示数据服务不可用，检查 Mac mini 本地服务：

```bash
curl -I http://127.0.0.1:3000
curl http://127.0.0.1:3000/api/dataset
```

如果 Mac mini 本地服务没启动，执行：

```bash
cd ~/quality-inspection-analytics
./scripts/update-local-prod.sh
./scripts/check-local-prod.sh
```

## 9. 推荐执行顺序

首次部署建议按这个顺序来：

1. 确认 Mac mini 本地 `http://127.0.0.1:3000` 可访问。
2. 域名解析到阿里云 VPS。
3. VPS 安装 Nginx 并配置访问密码。
4. VPS 配置 Nginx 反代到 `127.0.0.1:13000`。
5. VPS 申请 HTTPS 证书。
6. Mac mini 配置 SSH 密钥免密登录 VPS。
7. Mac mini 手动运行 `scripts/start-aliyun-tunnel.sh` 测试。
8. 打开公网域名验证。
9. Mac mini 注册 launchd 自动启动隧道。
