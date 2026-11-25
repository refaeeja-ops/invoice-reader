#!/bin/sh
# رفع رقم الإصدار: يحدّث ?v= في الصفحات و version.json (يُستخدم للتحديث التلقائي)
v=$(date +%Y%m%d%H%M)
sed -i -E "s/\.js\?v=[0-9]+/.js?v=$v/g; s/var V = '[0-9]+'/var V = '$v'/" index.html full.html
printf '{"v":"%s"}\n' "$v" > version.json
echo $v
