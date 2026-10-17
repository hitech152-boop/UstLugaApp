[app]
title = UstLugaApp
package.name = ustlugaapp
package.domain = org.ustluga
version = 1.0
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

# Пробуем pandas вместо openpyxl
requirements = python3,kivy==2.3.0,openpyxl,et_xmlfile
android.permissions = android.permission.INTERNET, android.permission.ACCESS_NETWORK_STATE
android.api = 33
android.minapi = 24
android.ndk_api = 24
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_license = True

[buildozer]
log_level = 2
# Убираем ветку develop, используем стабильный master
p4a.branch = develop
cython = 0.29.34
