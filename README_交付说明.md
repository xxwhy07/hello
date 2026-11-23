# 实训六 蜂鸣器案例 · 交付说明

## 一、交付物

| 交付物 | 路径 |
|---|---|
| 实训报告（Word，21 页） | `E:\deepseek harness\学习\实训六_蜂鸣器\实训六_蜂鸣器实训报告_向栩言.docx` |
| 交互式网页（单文件） | `E:\deepseek harness\学习\实训六_蜂鸣器\蜂鸣器实训交互网页.html` |
| 发布用副本 | `E:\deepseek harness\学习\实训六_蜂鸣器\site\index.html` |
| 公网链接 | **https://dxuy69.gicp.fun** |
| 链接与资源信息 | `发布链接.txt` |
| 可复制代码 | `code\beep\pwm_example.c`、`code\beep\CMakeLists.txt`、`code\CMakeLists_新增片段.txt`、`code\Kconfig_新增片段.txt` |
| 编译/烧录/自测记录 | `验证记录\build2.log`、`flash3.log`、`selftest-dom.html`、`preview*.png` |
| 报告素材 | `assets\` |

## 二、实物结果（已完成）

- 开发板：HiHope_NearLink_DK_WS63E_V03 + 交通灯板 + 底板。
- 串口：**COM5（USB-SERIAL CH340K）**。注意 COM3 是蓝牙虚拟串口，不能用。
- 编译：`python build.py -c ws63-liteos-app`，**SUCCESS**，耗时 316.185354 秒。
- 固件：`src\output\ws63\fwpkg\ws63-liteos-app\ws63-liteos-app_all.fwpkg`，1,463,144 字节。
- 烧录：`BurnToolCmd.exe --burn -n ws63 -m serial COM5 --baudRate 115200`，**Burned successfully**。
- 现象：复位后交通灯板蜂鸣器**鸣响约 1 秒**，本人已确认听到。
- 频率/占空比：约 **1066.7 Hz / 33.3%**（low=20000，high=10000，32 MHz 计数）。

## 三、你还需要做的 3 件事

1. **补 3 张实物照片**：打开 Word 报告，在“七、实训结果 → 7.3 实物照片与记录”里，把三个虚线框替换为：
   - 开发板 + 交通灯板连接、Type-C 接 PC 的照片；
   - HiSpark Studio 编译成功（SUCCESS）的截图；
   - 烧录完成 / 蜂鸣器鸣响的照片或视频截帧。
2. **核对封面**：课程名暂写“星闪技术应用”；如果你们课表叫“嵌入式系统开发”等，改封面和页眉两处即可。姓名/学号/班级/指导教师/日期已按你提供的信息填写。
3. **提交前检查**：用 Word/WPS 打开报告，确认 21 页、图片和表格没有跨页错位；如需要，另存 PDF 后提交。

> 报告里没有伪造照片：SDK 界面图均标注“SDK 参考界面”，三处实物照片是留给你补拍的虚线框。

## 四、SDK 改动与回滚

本次实际修改了本机 SDK（已备份）：

| 文件 | 改动 | 备份 |
|---|---|---|
| `src\application\samples\peripheral\beep\` | 新增样例（从 `vendor\...\demo\beep` 复制） | 若原来有，已改名为 `beep.bak_20261116` |
| `peripheral\CMakeLists.txt` | 在 `set(SOURCES...)` 前加 BEEP if 块 | `CMakeLists.txt.bak_20261116` |
| `peripheral\Kconfig` | 末尾加 `config SAMPLE_SUPPORT_BEEP` | `Kconfig.bak_20261116` |
| `ws63_liteos_app.config` | `CONFIG_SAMPLE_SUPPORT_BEEP=y`，暂时 `# CONFIG_SAMPLE_SUPPORT_LED is not set` | `ws63_liteos_app.config.bak_20261116` |
| `buttondemo\Kconfig` | 新建一个只含注释的 Kconfig，修复首次编译报错 | 新增文件，可删除 |

回滚方法：把上述 `.bak_20261116` 恢复原名，删除 `beep` 文件夹和 `buttondemo\Kconfig`（如果确认之前没有）。如果还要做 LED 实验，把 config 里的 LED 改回 `y`、BEEP 改 `n`，重新编译即可。

## 五、交互式网页

- 双击 `蜂鸣器实训交互网页.html` 即可离线打开；公网地址：**https://dxuy69.gicp.fun**。
- 8 个步骤：新建 beep → 复制样例 → CMake → Kconfig → menuconfig → 提交/编译 → 烧录 → 虚拟蜂鸣器。
- 只有代码提交并通过网页编译、再完成烧录模拟，才会解锁“鸣响 1 秒 / 持续鸣响”。
- 网页里的编译是**教学模拟器**（规则校验 + 拟真日志），真实编译结果以报告第六章和 `验证记录\build2.log` 为准。
- 虚拟蜂鸣器用 **Web Audio** 合成同频脉冲波，从电脑扬声器/耳机发声；浏览器不能直接驱动主板 PC 蜂鸣器。
- 自测：`https://dxuy69.gicp.fun/?selftest=1` 显示 `SELFTEST PASS`；预览第 8 步：`?preview=8`。
- 进度保存在浏览器 localStorage，换浏览器/清缓存会丢；页面上有“重置”。

## 六、重新烧录（如果需要）

1. Type-C 数据线连接板子和电脑，设备管理器确认 `USB-SERIAL CH340K (COM5)`。
2. 执行：
   ```bat
   cd /d E:\HiSparkStudioToolCHain\tools\BurnToolCmd
   BurnToolCmd.exe --burn -n ws63 -m serial COM5 --baudRate 115200 ^
     -f E:\HiSparkStudioSDK\fbb_ws63\src\output\ws63\fwpkg\ws63-liteos-app\ws63-liteos-app_all.fwpkg
   ```
3. 出现 `Please reset the device` 后 30 秒内按开发板 RST（或重新上电）。
4. 看到 `Burned successfully` 后复位，蜂鸣器应鸣响约 1 秒。
5. 如果 COM5 变了，以设备管理器里的 CH340K 端口为准。

## 七、验证记录

- `验证记录\build2.log`：真实编译全量日志，含 `pwm_example.c.obj`、`success`、316.185354 秒。
- `验证记录\flash3.log`：COM5 烧录日志，含 `Burned successfully`。
- `验证记录\selftest-dom.html`：网页自测结果，`SELFTEST PASS`（18 项）。
- `验证记录\preview0.png / preview6.png / preview8.png`：网页三张界面截图（报告中使用裁剪版）。
- `验证记录\publish2.log`：花生壳发布与认领结果，`claimed:true`、`public_url:https://dxuy69.gicp.fun`。
- 公网链接已用 `Invoke-WebRequest` 验证 200，并用 Edge 无头自测在线验证 `SELFTEST PASS`。
