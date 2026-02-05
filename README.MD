这是一个适配obsidian Hiwords插件设计canvas背单词,阅读markdown文件中单词的工具
支持windows系统执行文件,或py下载所需库后运行.设计思路使用人脑,但代码完全使用AI生成和调试.

1. 系统要求（Windows 10/11）
2. Python版本要求（3.8+）

作用是将xlsx单词表转为支持上述插件的canvas文件
obsidian不做介绍,教程很多.推荐熟悉markdown格式和愿意投入时间学习的朋友

单词配置项目必选单词和含义,可选音标,派生词,音标.
## 如何生成canvas
- 下载exe文件工具
- 加载表格
- 选中单词和含义所在列,
- 生成canvas

## 在obsidian中使用
**想使用预设词汇,可以直接下载canvas文件**
*但`考研单词5400.canvas`文件较为卡顿,不太推荐.日语`日语单词.canvas`标日初级单词表2200左右,也有点卡顿,酌情使用.*
将canvas文件放置在你的obsidian库中
(需要网络环境)
第三方插件(关闭安全模式)->浏览社区插件 搜索->Hiwords ->并启用
在下栏第三方插件
选择Hiwords 添加单词本 选中对应canvas文件
(使用单词不宜过多,影响性能)

可以加载非扫描版的pdf和文章md,
不一定所有含有的词的都能加载出来.文件较大也可能卡顿
Ctrl+p 键输入Hiwords(也可配置快捷键),选择
>Hiwords

Hiwords: 显示Hiwords 侧边栏
## 自行生成单词
如果想使用程序自行生成单词列表
可以下载exe文件(无需配置,但文件较大68mb)或者py文件

下面展示py文件的注意事项
自行打包
```powershell
git clone https://github.com/SantaChains/xlsx2canvas.git
cd xlsx2canvas
```
使用python
## 默认
```powershell
pip install pyinstaller
python xlsx_to_canvas.py
```
## 打包
```powershell
pip install pyinstaller
pyinstaller --name="XLSX转Canvas工具" --onefile --windowed --icon=xlsx_to_canvas.ico --clean --noconfirm --hidden-import=pandas --hidden-import=openpyxl --hidden-import=PyQt5.sip xlsx_to_canvas.py
```
如果有使用upx压缩程序大小则
```powershell
pyinstaller --name="XLSX转Canvas工具" --onefile --windowed --icon=xlsx_to_canvas.ico --clean --noconfirm --upx-dir="%UPX_DIR%" --hidden-import=pandas --hidden-import=openpyxl --hidden-import=PyQt5.sip xlsx_to_canvas.py
```
## 可能有用的资源(加github前缀)
[obsidian-skills(含json-canvas)](https://github.com/kepano/obsidian-skills) kepano/obsidian-skills
[Hiwords地址](https://github.com/CatMuse/HiWords)  CatMuse/HiWords
## 声明
表格资源来源网络搜集,如有侵权,请联系删除
感谢所有赏脸阅读的朋友,感谢实现这个功能付出过的同志!