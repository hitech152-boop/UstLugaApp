[app]
title = UstLugaApp
package.name = ustlugaapp
package.domain = org.ustluga
version = 1.0
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

# Добавляем jnius и фиксируем Cython
requirements = python3,kivy==2.3.0,openpyxl,et_xmlfile,jnius>=1.4.0,Cython==0.29.34

android.permissions = android.permission.INTERNET, android.permission.ACCESS_NETWORK_STATE
android.api = 33
android.minapi = 21
android.ndk = 25b
android.archs = arm64-v8a
android.accept_sdk_license = True

[buildozer]
log_level = 2
