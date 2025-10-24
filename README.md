M$ or WP$ ?
=========================================
为DOC/DOCX/XLS/XLSX/PPPT/PPTX 创建新的文件关联，使得双击文档时，自动调用原始编辑
程序（目前只考虑 WPS 或 MS Office）打开。如果文档是用WPS创建的，则调用WPS，否则
用Word打开。

项目初衷：工作中经常要收到很多人发来的文档，有人用Word编辑，有人用的WPS编辑。虽然 WPS Office 宣称兼容 MS Office，但并不完美，有时候WPS编辑的文档用WORD打开会出现一些莫名其妙的
排版错误。于是开发了这样一个工具，让 WPS 和 MS 各得其所吧……

项目用Rust语言开发，熟悉Rust生态的请自行编译。

用法：

    第一步：创建安装目录（比如: C:\Program Files\MsOrWPS）并将 GiveBackToCeasarOrGod.exe 和 config.toml 复制到目标目录中。

    第二步：编辑 config.toml，指定各个程序的路径。

    第三步：在安装目录中打开终端并运行下列命令以注册文件关联：
        GiveBackToCeasarOrGod.exe -r 

    然后在资源管理器中双击相应的文档看是否用正确的方式打开就行了。如果出现选择“打开方式”，选择 “M$ or WP$ ?”，点“始终”。


当执行第二步时，它为 DOC/DOCX/XLS/XLSX/PPPT/PPTX 各创建一个新的文件关联，当用户在资源管理器中双击相应的文档时，系统将文档路径传递给 GiveBackToCeasarOrGod.exe 。GiveBackToCeasarOrGod 只是一个中介，它判断原文档的元数据中的“程序名称"，如果是WPS则调用WPS打开，否则调用相应的 Microsoft Office 应用程序打开（系统中需要同时安装有 Microsoft Office 和 WPS office）。

如果双击文档时没有按预期打开文档，右键点击文档，选择“打开方式”，再选择“其它应用”，继续选择 “M$ or WP$ ?”，点“始终”。

如果注册文件关联后文档图标不对, 执行第二步并重启电脑或资源管理器；
