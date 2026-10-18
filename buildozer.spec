[app]
title = UstLugaApp
package.name = ustlugaapp
package.domain = org.ustluga
version = 1.0
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

# ВРЕМЕННО ОСТАВЛЯЕМ ТОЛЬКО KIVY ДЛЯ ПРОВЕРКИ
requirements = python3,kivy==2.3.0

android.permissions = android.permission.INTERNET, android.permission.ACCESS_NETWORK_STATE
android.api = 30
android.minapi = 21
android.ndk_api = 21
# ОСТАВЛЯЕМ ТОЛЬКО ОДНУ АРХИТЕКТУРУ
android.archs = arm64-v8a 
android.accept_sdk_license = True

[buildozer]
log_level = 2
