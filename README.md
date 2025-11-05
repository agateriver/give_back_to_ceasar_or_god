M$ or WP$ ?
=========================================

项目初衷：工作中经常要收到很多人发来的文档，有人用Word编辑，有人用的WPS编辑。虽
然  WPS Office 宣称兼容 MS Office，但并不完美，有时候WPS编辑的文档用WORD打开会出
现一些莫名其妙的排版错误。这时候就会想如果点击文档名能自动用原编辑程序打开该多好。
另外一个场景是会议,当演讲者在会议电脑中点击自己的文档时，他们希望用自己平常使用的
WPS office 或是 Microsoft Office 自动打开。于是开发了这样一个工具，让 WPS 和 MS 各得
其所吧……

此工具为DOC/DOCX/XLS/XLSX/PPPT/PPTX 创建新的文件关联，使得双击文档时，自动调用
原始编辑程序 （目前只考虑 WPS 或 MS Office）打开。如果文档是用WPS Office 创建的，
则调用WPS Office 相应组件打开，否则用 Microsoft Office 相应组件或系统默认程序打开。

用法：

    第一步：下载release版本，将压缩包内的顶级目录拖放到安装位置（比如: C:\Program Files\）中。


    第二步：在安装目录中双击register.cmd或者打开终端并运行下列命令以注册文件关联：
        give_back_to_ceasar_or_god.exe -r 

    然后在资源管理器中双击test目录中的各测试文档看是否用正确的方式打开就行了。如果出现
    选择“打开方式”，   选择 “自动调用 Microsoft 或 WPS Office 组件打开文档”，点“始终”。


当执行第二步时，它为 DOC/DOCX/XLS/XLSX/PPPT/PPTX 各创建一个新的文档类
型：Schrödinger's DOC/DOCX/XLS/XLSX/PPPT/PPTX，这些新的文档类型默认用
give_back_to_ceasar_or_god.exe 打开。当用户在资源管理器中双击相应的文档时，系统
将文档路径传递给give_back_to_ceasar_or_god.exe，give_back_to_ceasar_or_god 只是
一个中介，它判断原文档的元数据中的“程序名称"，如果是WPS则调用WPS打开，否则调用相
应的 Microsoft Office 应用程序打开（系统中需要同时安装有 Microsoft Office 和 WPS
office， 若缺则用系统默认程序打开）。

如果双击文档时没有按预期打开文档，右键点击文档，选择“打开方式”，再选择“其它应用”， 
继续选择 “自动调用 Microsoft 或 WPS Office 组件打开文档”或“在电脑上选择应用”并浏览
安装目录找到 give_back_to_ceasar_or_god.exe，点“始终”。

如果注册文件关联后文档图标不对, 执行第二步后重启电脑或资源管理器；
