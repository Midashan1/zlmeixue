#!/usr/bin/env bash
# 在腾讯云轻量应用服务器上部署「质量背锅侠&美学文章库」
#
# 用法（root 执行）：
#   bash deploy.sh [安装包目录] [端口]
#   安装包目录默认 /tmp/zlmeixue-deploy，端口默认 80
#
# 目录约定：
#   <包>/site/index.html      ← 唯一对公网开放的目录
#   <包>/site/articles.json
#   <包>/app/refresh.py       ← 抓取脚本（不对外）
#   <包>/app/album_fetch.py
#   <包>/app/accounts.json
#   <包>/data/all_articles.xlsx  ← 原始数据（不对外）
set -euo pipefail

APP=/opt/zlmeixue
SRC=${1:-/tmp/zlmeixue-deploy}
PORT=${2:-${PORT:-80}}
DOMAIN=${DOMAIN:-_}

# 兼容两种包结构：site/ app/ 分子目录，或全部平铺
if [ -f "$SRC/site/index.html" ]; then SITE_SRC="$SRC/site"; else SITE_SRC="$SRC"; fi
if [ -f "$SRC/app/refresh.py" ]; then APP_SRC="$SRC/app"; else APP_SRC="$SRC"; fi

echo "==> 1/7 安装依赖"
if command -v apt-get >/dev/null 2>&1; then
  export DEBIAN_FRONTEND=noninteractive
  apt-get update -qq
  apt-get install -y -qq nginx python3 python3-venv python3-pip curl unzip >/dev/null
elif command -v dnf >/dev/null 2>&1; then
  dnf install -y -q nginx python3 python3-pip curl unzip >/dev/null
elif command -v yum >/dev/null 2>&1; then
  yum install -y -q nginx python3 python3-pip curl unzip >/dev/null
else
  echo "!! 没识别到包管理器，请手动安装 nginx / python3-venv"; exit 1
fi

echo "==> 2/7 复制站点（只有这个目录对公网开放）"
mkdir -p "$APP/site" "$APP/app" "$APP/data" "$APP/logs"
test -f "$SITE_SRC/index.html"    || { echo "!! 缺少 index.html"; exit 1; }
test -f "$SITE_SRC/articles.json" || { echo "!! 缺少 articles.json"; exit 1; }
cp -f "$SITE_SRC/index.html" "$SITE_SRC/articles.json" "$APP/site/"

echo "==> 3/7 复制脚本与数据（不对公网开放）"
for f in refresh.py album_fetch.py accounts.json; do
  if [ -f "$APP_SRC/$f" ]; then cp -f "$APP_SRC/$f" "$APP/app/"; fi
done
if [ -f "$SRC/data/all_articles.xlsx" ]; then
  cp -f "$SRC/data/all_articles.xlsx" "$APP/data/"
fi
chmod 750 "$APP/app" "$APP/data"

echo "==> 4/7 准备 Python 环境"
if [ -f "$APP/app/refresh.py" ]; then
  [ -d "$APP/venv" ] || python3 -m venv "$APP/venv"
  "$APP/venv/bin/pip" install -q --upgrade pip
  "$APP/venv/bin/pip" install -q pandas openpyxl
fi

echo "==> 5/7 写 Nginx 配置（端口 $PORT，只服务 $APP/site）"
cat > /etc/nginx/conf.d/zlmeixue.conf <<CONF
server {
    listen ${PORT} default_server;
    listen [::]:${PORT} default_server;
    server_name ${DOMAIN};
    root ${APP}/site;
    index index.html;
    charset utf-8;
    server_tokens off;

    # 数据文件不缓存，保证更新后立刻生效
    location = /articles.json {
        add_header Cache-Control "no-store";
    }
    location / {
        try_files \$uri \$uri/ /index.html;
    }
    # 明确拒绝任何想摸到脚本/数据的路径
    location ~* \.(py|xlsx|json\.bak)$ { deny all; }

    access_log ${APP}/logs/access.log;
}
CONF
rm -f /etc/nginx/sites-enabled/default 2>/dev/null || true
nginx -t
systemctl enable nginx >/dev/null 2>&1 || true
systemctl restart nginx

echo "==> 6/7 每日刷新 + 每日数据快照（9:10 / 19:10）"
if [ -f "$APP/app/refresh.py" ]; then
  cat > /etc/cron.d/zlmeixue <<CRON
10 9,19 * * * root cd ${APP}/app && ${APP}/venv/bin/python refresh.py --json ${APP}/site/articles.json --quiet >> ${APP}/logs/refresh.log 2>&1 && cp -f ${APP}/site/articles.json ${APP}/data/articles-\$(date +\\%Y\\%m\\%d).json
CRON
  chmod 644 /etc/cron.d/zlmeixue
  (systemctl enable --now crond 2>/dev/null || systemctl enable --now cron 2>/dev/null || true)
fi

# 6.5) 套用正式站点配置（含 /zlmeixue/ 路径），避免重部署把路径冲掉
CANON=https://cdn.jsdelivr.net/gh/Midashan1/zlmeixue@main/server/nginx-zlmeixue.conf
if curl -fsSL "$CANON" -o /tmp/zlm-canon.conf 2>/dev/null; then
  sed -e "s/listen 8090/listen ${PORT}/g" -e "s/listen \[::\]:8090/listen [::]:${PORT}/g"       /tmp/zlm-canon.conf > /etc/nginx/conf.d/zlmeixue.conf
  if nginx -t >/dev/null 2>&1; then
    systemctl reload nginx
    echo "  已套用正式站点配置（访问路径 /zlmeixue/）"
  else
    echo "  ! 正式配置校验失败，保留端口直连配置"
  fi
fi

echo "==> 7/7 自检"
IP=$(curl -s --max-time 5 https://api.ipify.org || echo "<服务器IP>")
curl -s -o /dev/null -w "  首页        HTTP %{http_code}\n" "http://127.0.0.1:${PORT}/" || true
curl -s -o /dev/null -w "  articles    HTTP %{http_code}\n" "http://127.0.0.1:${PORT}/articles.json" || true
curl -s -o /dev/null -w "  脚本保护    HTTP %{http_code}（应为 403）\n" "http://127.0.0.1:${PORT}/refresh.py" || true

echo
echo "部署完成：http://${IP}:${PORT}/"
echo "打不开就去轻量服务器控制台「防火墙」放行 ${PORT} 端口。"
echo "每日数据快照在 ${APP}/data/articles-YYYYMMDD.json"
