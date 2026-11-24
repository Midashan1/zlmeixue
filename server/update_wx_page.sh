#!/bin/bash
# 在腾讯云轻量服务器上执行：把文章库网页同步为「含批量导出 Excel」的新版
set -e
SITE=/opt/zlmeixue/site
SRC=https://midashan1.github.io/zlmeixue/index.html

echo "==> 下载新版页面"
curl -fsSL --max-time 60 "$SRC" -o /tmp/wx_index.html
if ! grep -q "buildArticleXlsx" /tmp/wx_index.html; then
  echo "× 下载内容不是新版（缺少导出功能标记），已中止"; exit 1
fi
sudo cp /opt/zlmeixue/site/index.html /opt/zlmeixue/site/index.html.bak-$(date +%Y%m%d)
sudo cp /tmp/wx_index.html $SITE/index.html
echo "==> 已更新：$SITE/index.html"
wc -c $SITE/index.html

echo "==> 加入每日自动同步（以后改网页不用再登录服务器）"
sudo tee /etc/cron.d/zlmeixue-web >/dev/null <<'CRON'
5 9 * * * root curl -fsSL --max-time 60 https://midashan1.github.io/zlmeixue/index.html -o /tmp/wx_index.html && grep -q buildArticleXlsx /tmp/wx_index.html && cp -f /tmp/wx_index.html /opt/zlmeixue/site/index.html
CRON
sudo chmod 644 /etc/cron.d/zlmeixue-web
echo "==> 完成"
