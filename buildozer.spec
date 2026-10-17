[app]
title = UstLugaApp
package.name = ustlugaapp
package.domain = org.ustluga
version = 1.0

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

# Оставляем только самое необходимое
requirements = python3,kivy==2.3.0,openpyxl

android.permissions = android.permission.INTERNET, android.permission.ACCESS_NETWORK_STATE

android.api = 31
android.minapi = 24
android.ndk_api = 24
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_license = True

[buildozer]
log_level = 2
# Принудительно используем ветку develop
p4a.branch = develop
# Фиксируем версию Cython
cython = 0.29.34