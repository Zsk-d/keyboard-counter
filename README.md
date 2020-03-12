# Keyboard-counter
键盘按键次数记录

自行安装pyhook3依赖
简单实现功能，未深度优化，仅供娱乐

输出结果
![Image](https://raw.githubusercontent.com/Zsk-d/Keyboard-counter/master/img/result.png)

## keyPressRecoder.py
 
- init() : 初始化，启动时需调用

- exportDataFile() : 导出，调用后导出记录
- saveKeyPressRecordThread() : 调整sleep时间来修改自动保存间隔 
