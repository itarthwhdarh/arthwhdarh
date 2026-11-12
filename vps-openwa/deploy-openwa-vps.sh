#!/usr/bin/env bash
# ══════════════════════════════════════════════════════════════
# سكريبت تثبيت وتشغيل OpenWA تلقائياً على خادم سحابي Ubuntu/Debian
# جمعية إرث وحضارة بالقريات
# ══════════════════════════════════════════════════════════════

set -e

echo "=================================================="
echo "  بدء تثبيت خدمة OpenWA السحابية لجمعية إرث وحضارة"
echo "=================================================="

# 1. تحديث الحزم
sudo apt-get update -y
sudo apt-get install -y ca-certificates curl gnupg lsb-release ufw

# 2. تثبيت Docker إن لم يكن مثبتاً
if ! command -v docker &> /dev/null; then
    echo "[1/4] جاري تثبيت Docker..."
    sudo install -m 0755 -d /etc/apt/keyrings
    curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
    sudo chmod a+r /etc/apt/keyrings/docker.gpg
    echo \
      "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu \
      $(. /etc/os-release && echo "$VERSION_CODENAME") stable" | \
      sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
    sudo apt-get update -y
    sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
fi

# 3. إعداد المجلدات والصلاحيات
echo "[2/4] إعداد مجلدات التخزين المستمر..."
mkdir -p openwa-data openwa-sessions caddy-data caddy-config
chmod 777 openwa-data openwa-sessions

# 4. إعداد جدار الحماية (UFW)
echo "[3/4] ضبط جدار الحماية..."
sudo ufw allow ssh
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
sudo ufw allow 2785/tcp
sudo ufw --force enable || true

# 5. تشغيل الحاويات
echo "[4/4] تشغيل خدمة OpenWA عبر Docker Compose..."
sudo docker compose up -d

echo "=================================================="
echo "  تم تشغيل OpenWA بنجاح وهو يعمل الآن على مدار 24 ساعة!"
echo "  لوحة التحكم: http://$(curl -s ifconfig.me):2785"
echo "  التوثيق: http://$(curl -s ifconfig.me):2785/api/docs"
echo "=================================================="
