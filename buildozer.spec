[app]
title = UstLugaApp
package.name = ustlugaapp
package.domain = org.ustluga
version = 1.0
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

# 1. Явно указываем стабильную версию Python
# 2. Добавляем jnius, так как он часто требуется для работы с Android API
# 3. Фиксируем версию Cython, совместимую с Kivy
requirements = python3==3.10.11,hostpython3==3.10.11,kivy==2.3.1,openpyxl,et_xmlfile,jnius>=1.4.0,Cython==0.29.34

android.permissions = android.permission.INTERNET, android.permission.ACCESS_NETWORK_STATE

# 4. Указываем более новые версии API и NDK для совместимости с Kivy 2.3+
android.api = 33
android.minapi = 21
android.ndk = 25b
android.ndk_api = 21
android.archs = arm64-v8a

android.accept_sdk_license = True

[buildozer]
log_level = 2
# 5. Используем ветку develop для p4a, которая лучше совместима с Kivy 2.3+
p4a.branch = develop
