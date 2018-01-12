# <a name="task-panes-in-office-add-ins"></a>Office 外接程序中的任务窗格
 
任务窗格作为界面图面，通常出现在 Word、PowerPoint、Excel 和 Outlook 内的窗口右侧。任务窗格允许用户访问界面控件，此类控件运行代码以修改文档或电子邮件，或显示数据源中的数据。无需将功能直接嵌入文档时，使用任务窗格。

**示例：任务窗格**

![显示典型任务窗格布局的图像](../../images/overview_withApp_taskPane.png)

## <a name="best-practices"></a>最佳做法

|**允许事项**|**禁止事项**|
|:-----|:--------|
|<ul><li>在标题中包括外接程序的名称。</li></ul>|<ul><li>请勿在标题中追加公司名称。</li></ul>|
|<ul><li>在标题中使用简短的描述性名称。</li></ul>|<ul><li>不要在外接程序标题中追加“Add-in”、“For Word”或“for Office”等字符串。</li></ul>|
|<ul><li>在外接程序顶部包括某些导航或命令元素，如命令栏或透视。</li></ul>||
|<ul><li>在外接程序底部包括品牌元素，如品牌栏，除非要在 Outlook 内使用外接程序。</li></ul>||


## <a name="variants"></a>变量

下图显示分辨率为 1366x768 时 Office 功能区的各种任务窗格大小。对于 Excel，需要额外的垂直空间来容纳编辑栏。  

**Office 2016 桌面任务窗格大小**

![显示大小为 1366x768 的桌面任务窗格的图像](../../images/addinTaskpaneSizes_desktop.png)

- Excel - 320 x 455
- PowerPoint - 320 x 531
- Word - 320 x 531
- Outlook - 348 x 535

**Office 365 任务窗格大小**

![显示大小为 1366x768 的桌面任务窗格的图像](../../images/addinTaskpaneSizes_online.png)

- Excel - 350 x 378
- PowerPoint - 348x391
- Word - 329 x 445
- Outlook Web App - 320x570

## <a name="personality-menu"></a>“个性”菜单

“个性”菜单可能会妨碍靠近外接程序右上角的导航和命令元素。以下是 Windows 和 Mac 上的“个性”菜单的当前尺寸。

**Windows 上的“个性”菜单**

对于 Windows，“个性”菜单尺寸为 12 x 32 像素，如下所示。

![显示 Windows 桌面上的“个性”菜单的图像](../../images/personalityMenu_Win.png)

**Windows 上的“个性”菜单**

对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将空间增加至 34x32 像素，如下所示。

![显示 Mac 桌面上的“个性”菜单的图像](../../images/personalityMenu_Mac.png)

## <a name="implementation"></a>实现

有关实现任务窗格的示例，请参阅 GitHub 上的 [Excel 外接程序 JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)。 


## <a name="additional-resources"></a>其他资源

- [Office 外接程序中的 Office UI Fabric](office-ui-fabric.md) 
- [适用于 Office 外接程序的 UX 设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)


