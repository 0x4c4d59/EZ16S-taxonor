# EZ16S-taxonor
![Static Badge](https://img.shields.io/badge/Language-Python-green)
![Static Badge](https://img.shields.io/badge/Browser-Chrome-green)
![Static Badge](https://img.shields.io/badge/License-GPL--3.0-blue)

---
故事的起因是[Ezbiocloud](https://www.ezbiocloud.net/identify)无法批量根据16S序列检索物种信息。

本脚本可以实现批量比对功能。

>ps: 由于个人代码水平较差，目前仅依托Chrome平台下的chromedriver进行爬虫的编写。

### 使用指南

1. 根据Chrome版本下载其对应的[chromedriver](https://chromedriver.storage.googleapis.com/index.html)，将压缩包解压在本地PC的任意文件夹中;
2. 下载release中打包好的exe文件，填写相关参数进行批量比对。

### 参数解释

- FASTA文件路径：待比对的`.fasta`格式文件
- Chrome Driver路径：本地PC上的chromedriver路径
- 用户名 密码：[Ezbiocloud](https://www.ezbiocloud.net/identify)上对应注册的帐号密码
    - 记住密码：在exe文件所在根目录生成`config.ini`于储存信息，**注意信息安全！**
- 输出路径：输出`.xlsx`结果文件，内容包括[Ezbiocloud](https://www.ezbiocloud.net/identify)所有比对结果，以及输入的序列信息
- 比对后等待时间（秒）：每条序列提交后等待系统检索的时间，若网络良好可适当降低，反之升高。默认3s
    - 参数过小导致输出结果不完整
    - 参数过大导致整体运行时间偏长

### 注意！
- 运行脚本前建议先手动清除网站上先前比对的记录，否则会导致本次输出结果内掺杂先前的比对结果
- 由于暂时没有开发`headless`界面，因此运行软件时请不要最小化窗口，也尽量不要使用鼠标操作弹出窗口
  - ~~`headless`界面什么的可能后续版本会加上吧，会有的吧（挖坑）~~