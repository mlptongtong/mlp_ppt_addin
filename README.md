# mlp_ppt_addin
安装过程：

1. 首先解压压缩包，先关闭所有PPT程序，再双击setup.exe，如果电脑有杀毒软件禁止安装exe，可以双击MLP_PPT.vsto进行安装：

![image text](http://www.allmlp.com/#/blogs/blogdetail?id=255/data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAsIA…AAiRAAiRAAiRAApkk8H/ytK/zRTyZwAAAAABJRU5ErkJggg== "DBSCAN Performance Comparison")


2. 点击Install/安装：



3.点击Close完成安装：



4. 如果安装失败，请双击Update.bat，然后再次安装MLP_PPT.vsto:



如果报错部署清单签名的问题，如下：



请分别双击下面两个文件reg1.reg和reg2.reg，再次安装即可：



5. 打开PPT后，可以看到MLP导航栏：



卸载软件：

再次安装前请首先卸载软件，卸载步骤：

两种卸载方式：

一、通过自带的卸载uninstall.exe进行卸载：



双击后，依次进行卸载，注意卸载前会强行关闭所有PPT，确保卸载前保存您的PPT内容。





二、通过手动方式，操作步骤如下。（强烈建议通过第一种方法进行自动卸载，如无法自动卸载，再通过手动卸载）

1. 首先，打开PPT，点击开发者(Developer)如果没有，请到选项（Option）中添加此ribbon，然后再点击MLP_PPT，点击remove（删除）。



2. 其次，关闭所有PPT，请确保在任务管理器中是不存在的，否则安装的版本依旧是老版本。





3. 在控制面板找到程序路径，点击卸载：





主要功能：

1. FreezeSlide:

可以锁定当前Slide(幻灯片)，让你在编辑幻灯片时，不会跳转到其它幻灯片页。



2. FolderOpen:

可以快速打开当前PPT所在文件夹路径。



3. CopyFile:

可以快速复制当前PPT，然后通过Ctrl+v粘贴到具体位置。



4. Ex.Tool:

可以快速打开常用的windows工具。





5. Pic.：

可以打开常用图片网站，进行图片的下载。



6. Icon：

可以打开iconfont网站，可以下载海量ICON。



7. Vectorgraph：

可以打开PPT素材综合网站，可以下载海量素材。



8. AddShapes:

可以快速的创建指定数量和指定方向的shapes。先鼠标左键选中指定的Shape，窗体中的Width和Height指的是Shape的宽和高，X，Y可以指定shapes的创建方向，QTY是指沿着指定方向创建的shapes数量，Element-Dis是指同组内的shapes彼此之间的间距，Group-Dis是指组间的间距。（同一行的shapes称为一组）



亦可以，通过点选Example里预设定的方案来创建shapes，再通过上面的参数进行微调，通过点击“Confirm”生成最终的shapes。



9. AjustFont:

可以快速的将PPT中所有slides中的字体进行统一修改。控件会列举出所有slides中存在的字体，并且通过下拉框查看每个字体对应的控件所在位置，通过选择要修改的字体复选框，再选择要修改成的字体，点击Confirm按钮即可完成统一修改。





10. FindTemplate：

线上提供经典的PPT模板/素材，并且支持上传你专属的PPT模板/素材。

该功能需要首先完成注册，登陆后才可使用。

Public中是通用PPT模板和素材，是所有人可以使用的公共资源，可以通过搜索框查找相应的模板。



Private是专属模板和素材，是关联到具体账户的资源，只有本人可以使用，同时提供删除功能，以及分享到公共资源中的功能。



Publish提供上传资源的接口，通过“Synchronize to public”选框，可以设置上传的资源是否同步到公共资源，description是上传资源的描述，主要给搜索功能使用，方便快速定位到具体资源，refresh为刷新功能，当编辑好Slide(幻灯片)后，点击refresh，图片框会显示最新slide，然后可以通过Submit上传资源。

