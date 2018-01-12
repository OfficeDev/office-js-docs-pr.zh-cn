# <a name="content-office-add-ins"></a>内容 Office 外接程序

内容外接程序这种图面可被直接嵌入 Word、Excel 或 PowerPoint 文档中。内容外接程序让用户访问运行代码以修改文档或显示数据源中数据的界面控件。在你要将功能直接嵌入文档时，请使用内容加载项。  

**示例：内容外接程序**

![显示内容外接程序的典型布局的示例图像。](../../images/overview_withApp_content.png)

## <a name="best-practices"></a>最佳做法

- 在外接程序顶部包括某些导航或命令元素，如命令栏或透视。
- 包括位于外接程序底部的品牌元素，如品牌栏（仅适用于 Word、Excel 和 PowerPoint 外接程序）。

## <a name="variants"></a>变量

Office 2016 桌面和 Office 365 中的 Word、Excel 和 PowerPoint 的内容外接程序大小由用户指定。

## <a name="personality-menu"></a>“个性”菜单

“个性”菜单可能会妨碍靠近外接程序右上角的导航和命令元素。以下是 Windows 和 Mac 上的“个性”菜单的当前尺寸。

**Windows 上的“个性”菜单** 

对于 Windows，“个性”菜单尺寸为 12 x 32 像素，如下所示。

![显示 Windows 桌面上的“个性”菜单的图像](../../images/personalityMenu_Win.png)

**Mac 上的“个性”菜单**

对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将占用空间增加至 34x32 像素，如下所示。

![显示 Mac 桌面上的“个性”菜单的图像](../../images/personalityMenu_Mac.png)

## <a name="implementation"></a>实现

有关实现内容加载项的示例，请参阅 GitHub 中的 [ Excel 内容加载项 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)。

## <a name="additional-resources"></a>其他资源

- [Office 外接程序中的 Office UI Fabric](office-ui-fabric.md) 
- [适用于 Office 加载项的 UX 设计模式](ux-design-patterns.md)