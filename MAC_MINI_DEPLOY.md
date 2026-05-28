# Mac mini LAN Deployment

## Best-fit use case

Use this path when:
- you and your boss are usually on the same LAN
- the Mac mini stays powered on
- you want to avoid cloud hosting costs

## Recommended runtime shape

Run one production service on the Mac mini:
- frontend served from `dist/`
- backend API served from `server/index.ts`
- shared data stored in local `data/shared-dataset.json`

## 1. Put the project on the Mac mini

Recommended path:

```bash
cd ~
git clone https://github.com/kimi-tc/quality-inspection-analytics.git
cd quality-inspection-analytics
```

## 2. Start it manually first

```bash
cd ~/quality-inspection-analytics
chmod +x scripts/start-local-prod.sh
PORT=3000 PROJECT_DIR=~/quality-inspection-analytics DATA_DIR=~/quality-inspection-analytics/data ./scripts/start-local-prod.sh
```

Then open on the Mac mini:

```text
http://127.0.0.1:3000
```

## 3. Let others in the LAN visit it

Find the Mac mini LAN IP:

```bash
ipconfig getifaddr en0
```

If Wi-Fi is used and `en0` is empty, try:

```bash
ipconfig getifaddr en1
```

Then others in the same LAN can open:

```text
http://<MAC_MINI_IP>:3000
```

Example:

```text
http://192.168.1.23:3000
```

## 4. Make the IP stable

In your router admin page, bind the Mac mini to a fixed DHCP IP.

Recommended:
- reserve one fixed IP for the Mac mini
- keep port `3000`

## 5. Enable auto-start at boot

1. Copy `deploy/com.kimi.weekly-inspection-analytics.plist`
2. Replace every `REPLACE_ME` with the Mac mini account name
3. Copy it to:

```bash
~/Library/LaunchAgents/com.kimi.weekly-inspection-analytics.plist
```

4. Load it:

```bash
launchctl unload ~/Library/LaunchAgents/com.kimi.weekly-inspection-analytics.plist 2>/dev/null || true
launchctl load ~/Library/LaunchAgents/com.kimi.weekly-inspection-analytics.plist
```

Check status:

```bash
launchctl list | grep weekly-inspection
```

## 6. Update the app later

```bash
cd ~/quality-inspection-analytics
chmod +x scripts/update-local-prod.sh
./scripts/update-local-prod.sh
```

## 7. Logs

After launchd is enabled:

```bash
tail -f ~/Library/Logs/weekly-inspection-analytics.log
tail -f ~/Library/Logs/weekly-inspection-analytics-error.log
```

## 8. Practical recommendation

For your current situation, this is the most cost-effective path.

You can keep this LAN deployment first, and only later add:
- Tailscale
- reverse proxy
- cloud database

if remote access becomes necessary.
