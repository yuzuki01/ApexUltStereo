# ApexUltStereo
基于图像识别和比例控制的Apex瓦鸡车载音响

# [INFO]
 - 音频文件存放于 assets 文件夹下
 - 支持.wav和.mp3格式
 - 长度控制在 15 秒 效果最佳
 - 在自己的电脑上测试的结果：
 -  - 内存占用： 约 50 MB
 -  - 完成一次循环用时： 约 0.04 秒 (为了防止卡顿，每次循环有 0.5 秒的休眠)
 -  - 测试的电脑配置：
 -  -  - CPU：AMD Ryzen 9 5900HX with Radeon Graphics    3.30 GHz
 -  -  - RAM：32 GB
 -  -  - GPU：NVIDIA GeForce RTX 3060 Laptop GPU

# [NOTICE]
 - 运行前请先设置 screen_dpi.json 保住分辨率和游戏内分辨率一致
 - 请使用英文字母及阿拉伯数字命名音频文件，汉字及其他字符可能会导致程序崩溃
 - 目前仅支持：
 -  - 瓦尔基里-使用大招
 - 使用的库（有版本要求的请严格安装版本要求安装）：
 -  - PyQt5
 -  - PIL
 -  - win32gui
 -  - playsound (version==1.2.2)

# [FAQ]
 - Q：为什么用图像识别？
 - A：一开始打算仿制外挂，直接获取游戏内数据，但是有被反作弊检测到的风险，同时技术力不够也不知道怎么获取想获取的数据。
