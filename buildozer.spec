[app]
title = UstLugaApp
package.name = ustlugaapp
package.domain = org.ustluga
version = 1.0

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

# ИЗМЕНЕНО: убрана жёсткая привязка к 3.10.11, убран jnius из требований
# ИЗМЕНЕНО: kivy==2.3.0 — проверенный релиз
requirements = python3,hostpython3,kivy==2.3.0,openpyxl,et_xmlfile,Cython==0.29.34

# Разрешения для доступа в интернет
android.permissions = android.permission.INTERNET, android.permission.ACCESS_NETWORK_STATE

# ИЗМЕНЕНО: NDK 23b — стабильнее для Kivy 2.x, чем 25b
android.api = 33
android.minapi = 21
android.ndk = 23b
android.ndk_api = 21
android.archs = arm64-v8a

android.accept_sdk_license = True

# ДОБАВЛЕНО: корректная работа с кириллицей в логах и файлах
android.utf8 = True

[buildozer]
log_level = 2

# ИЗМЕНЕНО: master вместо develop — стабильная сборка
p4a.branch = master
