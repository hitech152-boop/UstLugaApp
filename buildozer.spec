[app]
title = UstLugaApp
package.name = ustlugaapp
package.domain = org.ustluga
version = 1.0
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

# Явно указываем Python 3.10 и фиксируем Kivy
requirements = python3==3.10.11,hostpython3==3.10.11,kivy==2.3.1,openpyxl,et_xmlfile

android.permissions = android.permission.INTERNET, android.permission.ACCESS_NETWORK_STATE
android.api = 31
android.minapi = 24
android.ndk_api = 24
android.archs = arm64-v8a
android.accept_sdk_license = True

[buildozer]
log_level = 2
