#!/usr/bin/env bash
# 一键引导：从 CDN 拉取站点与脚本，然后调用 deploy.sh 完成安装
# 用法： curl -fsSL <本脚本URL> | sudo bash -s 8080
set -euo pipefail

PORT=${1:-8080}
BASE=https://cdn.jsdelivr.net/gh/Midashan1/zlmeixue@main
D=/tmp/zlmeixue-deploy

echo "==> 准备目录 $D"
rm -rf "$D"
mkdir -p "$D/site" "$D/app" "$D/data"

echo "==> 下载站点文件"
curl -fsSL "$BASE/v2/index.html"    -o "$D/site/index.html"
curl -fsSL "$BASE/v2/articles.json" -o "$D/site/articles.json"

echo "==> 下载抓取脚本"
curl -fsSL "$BASE/server/refresh.py"     -o "$D/app/refresh.py"
curl -fsSL "$BASE/server/album_fetch.py" -o "$D/app/album_fetch.py"
curl -fsSL "$BASE/server/accounts.json"  -o "$D/app/accounts.json"

echo "==> 下载部署脚本"
curl -fsSL "$BASE/server/deploy.sh" -o "$D/deploy.sh"

ls -lh "$D/site" "$D/app"
echo "==> 开始安装（端口 $PORT）"
bash "$D/deploy.sh" "$D" "$PORT"
