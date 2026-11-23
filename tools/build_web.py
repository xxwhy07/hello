# -*- coding: utf-8 -*-
"""Build the self-contained interactive lab HTML."""
import base64
import json
import mimetypes
import os

ROOT = r"E:\deepseek harness\学习\实训六_蜂鸣器"
TEMPLATE = os.path.join(ROOT, "web", "index.template.html")
CODE = os.path.join(ROOT, "code", "beep", "pwm_example.c")
ASSETS = os.path.join(ROOT, "assets")
OUT_SITE = os.path.join(ROOT, "site", "index.html")
OUT_DELIV = os.path.join(ROOT, "蜂鸣器实训交互网页.html")

CMAKE_TEMPLATE = """# peripheral/CMakeLists.txt（节选）
# 说明：真实文件里前面已经有很多 add_subdirectory_if_exist(...)
# 你只需要在 set(SOURCES "${SOURCES}" PARENT_SCOPE) 之前注册 beep。

# TODO: 在这里写 BEEP 注册块（if / add_subdirectory_if_exist / endif）

set(SOURCES "${SOURCES}" PARENT_SCOPE)
"""

KCONFIG_TEMPLATE = """# peripheral/Kconfig（文件末尾节选）
# 说明：文件前面已有 LED、BUTTON 等配置。
# 请在下面追加 BEEP 的配置项。

# TODO: 追加 config SAMPLE_SUPPORT_BEEP ...

menu "BUTTON Sample Configuration"
source "application/samples/peripheral/buttondemo/Kconfig"
endmenu
"""

IMAGES = {
    "IMG_HARDWARE": "hardware.jpg",
    "IMG_NEW_FOLDER": "new_folder.png",
    "IMG_COPY_SAMPLE": "copy_sample.png",
    "IMG_CMAKE": "cmake.png",
    "IMG_KCONFIG": "kconfig.png",
    "IMG_MENUCONFIG": "menuconfig.png",
    "IMG_SUPPORT_BEEP": "support_beep.png",
    "IMG_BUILD_SUCCESS": "build_success.jpg",
    "IMG_PROJECT_CONFIG": "project_config.jpg",
    "IMG_FLASH_START": "flash_start.png",
    "IMG_FLASH_CONNECTING": "flash_connecting.png",
    "IMG_BOARD_USB": "board_usb.jpg",
}


def data_uri(path):
    mime = mimetypes.guess_type(path)[0] or "application/octet-stream"
    with open(path, "rb") as f:
        b = base64.b64encode(f.read()).decode("ascii")
    return f"data:{mime};base64,{b}"


def main():
    with open(TEMPLATE, "r", encoding="utf-8") as f:
        html = f.read()
    with open(CODE, "r", encoding="utf-8") as f:
        code = f.read()

    html = html.replace("{{C_CODE_JSON}}", json.dumps(code, ensure_ascii=False))
    html = html.replace("{{CMAKE_TEMPLATE_JSON}}", json.dumps(CMAKE_TEMPLATE, ensure_ascii=False))
    html = html.replace("{{KCONFIG_TEMPLATE_JSON}}", json.dumps(KCONFIG_TEMPLATE, ensure_ascii=False))
    for key, name in IMAGES.items():
        path = os.path.join(ASSETS, name)
        if not os.path.exists(path):
            raise FileNotFoundError(path)
        html = html.replace("{{" + key + "}}", data_uri(path))

    os.makedirs(os.path.dirname(OUT_SITE), exist_ok=True)
    for path in (OUT_SITE, OUT_DELIV):
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(html)
    print("written", OUT_SITE, f"{os.path.getsize(OUT_SITE)/1024:.1f} KB")
    print("written", OUT_DELIV, f"{os.path.getsize(OUT_DELIV)/1024:.1f} KB")


if __name__ == "__main__":
    main()
