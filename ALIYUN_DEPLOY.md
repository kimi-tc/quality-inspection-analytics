# Aliyun VPS Public Access With Local Data

This deployment keeps the dataset on the Mac mini and uses an Aliyun VPS only as a public HTTPS entry.

## Architecture

```text
Boss browser
  -> https://your-domain
  -> Aliyun VPS nginx + basic auth
  -> 127.0.0.1:13000 on VPS
  -> reverse SSH tunnel
  -> Mac mini 127.0.0.1:3000
  -> local data/shared-dataset.json
```

Why this fits the current requirement:

- Weekly data remains on the Mac mini.
- The VPS does not store the imported Excel data.
- You can give your boss a normal domain URL.
- No home/company router port forwarding is required.

## 0. Prerequisites

You need:

- Aliyun ECS/VPS public IP.
- A domain or subdomain pointed to the VPS public IP.
- SSH access from the Mac mini to the VPS.
- The Mac mini local dashboard already running at `http://127.0.0.1:3000`.

Suggested DNS:

```text
dashboard.your-domain.com -> <ALIYUN_VPS_PUBLIC_IP>
```

Aliyun security group should allow:

- TCP `22` from your Mac mini/network for SSH.
- TCP `80` and `443` from the internet for web access.

## 1. Configure the VPS

SSH into the VPS:

```bash
ssh root@<ALIYUN_VPS_PUBLIC_IP>
```

Install nginx and basic auth tooling.

Ubuntu/Debian:

```bash
apt update
apt install -y nginx apache2-utils certbot python3-certbot-nginx
systemctl enable --now nginx
```

CentOS/Alibaba Cloud Linux:

```bash
yum install -y nginx httpd-tools certbot python3-certbot-nginx
systemctl enable --now nginx
```

Create a login account for the dashboard:

```bash
htpasswd -c /etc/nginx/.weekly-inspection.htpasswd dashboard
```

Copy the nginx config template:

```bash
nano /etc/nginx/conf.d/weekly-inspection.conf
```

Paste `deploy/aliyun-nginx-weekly-inspection.conf`, then replace:

```text
dashboard.example.com
```

with your real domain.

Validate and reload nginx:

```bash
nginx -t
systemctl reload nginx
```

## 2. Enable HTTPS

After DNS has resolved to the VPS:

```bash
certbot --nginx -d dashboard.your-domain.com
```

Choose the redirect-to-HTTPS option when prompted.

## 3. Configure Mac mini SSH key

On the Mac mini:

```bash
ssh-keygen -t ed25519 -C "weekly-inspection-tunnel" -f ~/.ssh/weekly_inspection_tunnel
```

Copy the public key to the VPS:

```bash
ssh-copy-id -i ~/.ssh/weekly_inspection_tunnel.pub root@<ALIYUN_VPS_PUBLIC_IP>
```

If `ssh-copy-id` is unavailable:

```bash
cat ~/.ssh/weekly_inspection_tunnel.pub
```

Then paste the output into this file on the VPS:

```bash
~/.ssh/authorized_keys
```

Test passwordless SSH:

```bash
ssh -i ~/.ssh/weekly_inspection_tunnel root@<ALIYUN_VPS_PUBLIC_IP> "echo ok"
```

## 4. Configure the tunnel on Mac mini

Create a local env file on the Mac mini:

```bash
nano ~/.weekly-inspection-tunnel.env
```

Example:

```bash
VPS_USER=root
VPS_HOST=<ALIYUN_VPS_PUBLIC_IP>
VPS_PORT=22
REMOTE_PORT=13000
LOCAL_PORT=3000
PUBLIC_URL=https://dashboard.your-domain.com
```

Add SSH key config:

```bash
nano ~/.ssh/config
```

Example:

```sshconfig
Host aliyun-weekly-inspection
  HostName <ALIYUN_VPS_PUBLIC_IP>
  User root
  IdentityFile ~/.ssh/weekly_inspection_tunnel
  ServerAliveInterval 30
  ServerAliveCountMax 3
```

If you use the SSH alias above, set:

```bash
VPS_USER=root
VPS_HOST=aliyun-weekly-inspection
```

Start the tunnel manually:

```bash
cd ~/quality-inspection-analytics
chmod +x scripts/start-aliyun-tunnel.sh scripts/check-aliyun-tunnel.sh
./scripts/start-aliyun-tunnel.sh
```

In another terminal, check:

```bash
./scripts/check-aliyun-tunnel.sh
```

## 5. Enable Mac mini auto-start

Copy the launchd plist:

```bash
cd ~/quality-inspection-analytics
cp deploy/com.kimi.weekly-inspection-tunnel.plist ~/Library/LaunchAgents/
launchctl unload ~/Library/LaunchAgents/com.kimi.weekly-inspection-tunnel.plist 2>/dev/null || true
launchctl load ~/Library/LaunchAgents/com.kimi.weekly-inspection-tunnel.plist
launchctl list | grep weekly-inspection-tunnel
```

Logs:

```bash
tail -f ~/Library/Logs/weekly-inspection-tunnel.log
tail -f ~/Library/Logs/weekly-inspection-tunnel-error.log
```

## 6. Daily operation

To update the dashboard code on Mac mini:

```bash
cd ~/quality-inspection-analytics
./scripts/update-local-prod.sh
```

To check the local service and public tunnel:

```bash
cd ~/quality-inspection-analytics
./scripts/check-local-prod.sh
./scripts/check-aliyun-tunnel.sh
```

## 7. Security notes

Keep nginx basic auth enabled. The current app has import and clear-data actions in the UI, so a public URL should not be open without authentication.

Recommended:

- Use HTTPS.
- Use a strong basic-auth password.
- Limit SSH security-group source IP if possible.
- Back up `data/shared-dataset.json` regularly.
- Avoid committing `~/.weekly-inspection-tunnel.env` or SSH keys.

## 8. Troubleshooting

If the public URL returns `502 Bad Gateway`:

```bash
ssh root@<ALIYUN_VPS_PUBLIC_IP> "curl -I http://127.0.0.1:13000"
```

If that fails, the reverse tunnel is not connected.

On the Mac mini:

```bash
launchctl list | grep weekly-inspection-tunnel
tail -n 80 ~/Library/Logs/weekly-inspection-tunnel-error.log
```

If the public URL asks for username/password forever:

```bash
htpasswd /etc/nginx/.weekly-inspection.htpasswd dashboard
nginx -t
systemctl reload nginx
```

If the dashboard opens but data cannot load, check the Mac mini local service:

```bash
curl -I http://127.0.0.1:3000
curl http://127.0.0.1:3000/api/dataset
```
