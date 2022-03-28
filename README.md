【需求】由于本地计算机性能原因，在将文档转换为PDF时比较慢，因此想开发一个简易的工具来转换文档，同时有时候想要切割和合并多个文件等，所以开发文件助手。


0327：

**【实现效果】----当前项目已基本完善，但长期维护，有问题联系作者：3538182550@qq.com**

**最新安装方式：直接下载“音符文档助手.exe”安装包，选择路径安装即可**

1、功能：
  
  a. word转pdf--需要电脑已安装office套件（wps也行）
  b. pdf转word--部分文件可能会出现乱码--功能有待完善
  c. 图片转pdf--将图片整合为一个pdf文件
  d. pdf转图片--将pdf文件按页转为图片
  e. 切割pdf--按照用户输入切割目标页数为单独的pdf文件
  f. 合并pdf--将选中的pdf合并为一个pdf文件
  g. 选择文件--有拖动选择和点击选择两种方式（可多选）
  h. 更换背景图--用户可自己更换背景图
  
效果：

![image](https://user-images.githubusercontent.com/81294772/160270761-687a39e1-4729-4953-884c-45d4c4c77f9f.png)
![image](https://user-images.githubusercontent.com/81294772/160270773-d72c830b-7038-45e6-b3b2-7e28b1332b34.png)
![image](https://user-images.githubusercontent.com/81294772/160270919-9fd09b03-6e19-4c0f-ae46-d33524752a4e.png)

**安装**

1、直接将./dist文件夹放置到你希望防止的地方如D:/temp

2、打开目录./dist/音符文档助手/，双击“音符文档助手.exe”文件即可启动

3、您也可右键单击“音符文档助手.exe”文件，选择“发送到”->“桌面快捷方式”，即可创建快捷方式，之后能够便捷启动软件


【当前状况】起步，刚实现文档转PDF，文档拆分和切割正在整合，PDF转word的效果不好或者说极差，所以现在主要是word转pdf好用。实现效果如下：

![image](https://user-images.githubusercontent.com/81294772/147771456-0545371e-1a87-40c7-b500-3b37f557acb0.png)


【想法】纯当练手，希望加入一些日常需求，加快办公速度

【开发】1、基于github编程，代码大多是借鉴或者说全盘拿来的，希望实现自主开发；
2、增加代码复用，避免“造轮子”。

【开发】1、基于python的第三方库如wxpython(GUI),pdfminer等等

3、本开发基于python脚本语言

4、希望做强做大走向富裕

【资源】1、资源已上传gitee代码仓库（源代码）：

https://gitee.com/TangGarlic/fileSystem

2、资源已上传github代码仓库（源代码）：

https://github.com/TonyTang-dev/fileSystem

3、均已开源，可自行下载同步开发
4、目前主要维护的是国内开源仓库gitee.

Git下载：
1、安装远程版本控制工具git

2、配置环境变量，学习简单实用方法

3、在终端执行命令：git clone https://gitee.com/TangGarlic/file-handing 下载源代码并可同步开发


【项目结构】

![image](https://user-images.githubusercontent.com/81294772/147771415-90874fe8-0994-4036-a8ef-68bea444d59c.png)
