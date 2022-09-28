---
title: Office 加载项中的任务窗格
description: 任务窗格允许用户访问界面控件，此类控件运行代码以修改文档或电子邮件，或显示数据源中的数据。
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: d911101a7df1f1ad8aa01b8e0006bd93d994a193
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092915"
---
# <a name="task-panes-in-office-add-ins"></a>Office 加载项中的任务窗格

Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*图 1：典型任务窗格布局*

![显示典型任务窗格布局的插图，顶部有分区选项卡，左下角显示公司徽标和公司名称，右下角有设置图标。](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>最佳做法

|允许事项|禁止事项|
|:-----|:--------|
|在标题中包括外接程序的名称。|请勿在标题中追加公司名称。|
|在标题中使用简短的描述性名称。|不要将“加载项”、“For Word”或“for Office”等字符串追加到加载项的标题中。|
|在加载项顶部包括某些导航或命令元素，如命令栏或透视。|*没有。*|
|在外接程序底部包括品牌元素，如品牌栏，除非要在 Outlook 内使用外接程序。|*没有。*|

## <a name="variants"></a>变量

下图显示了各种任务窗格大小，其中 Office 应用功能区分辨率为 1366x768。 对于 Excel，需要额外的垂直空间来容纳编辑栏。  

*图 2：Office 2016 桌面任务窗格尺寸*

![显示桌面任务窗格大小为 1366x768 分辨率的图表。](../images/office-2016-taskpane-sizes.png)

- Excel - 320x455 像素
- PowerPoint - 320x531 像素
- Word - 320x531 像素
- Outlook - 348x535 像素

<br/>

*图 3.Office 任务窗格大小*

![显示 1366x768 分辨率的任务窗格大小的图示。](../images/office-365-taskpane-sizes.png)

- Excel - 350x378 像素
- PowerPoint - 348x391 像素
- Word - 329x445 像素
- Web) 上的 Outlook (- 320x570 像素

## <a name="personality-menu"></a>“个性”菜单

“个性”菜单可能会妨碍靠近外接程序右上角的导航和命令元素。 以下是 Windows 和 Mac 上的“个性”菜单的当前尺寸。  (Outlook.) 不支持个性菜单

对于 Windows，个性菜单尺寸为 12x32 像素，如下所示。

*图 4：Windows 上的个性菜单*

![显示 Windows 桌面上的个性菜单的图示。](../images/personality-menu-win.png)

对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将空间增加至 34x32 像素，如下所示。

*图 5：Mac 上的个性菜单*

![显示 Mac 桌面上的个性菜单的图示。](../images/personality-menu-mac.png)

## <a name="implementation"></a>实现

有关实现任务窗格的示例，请参阅 GitHub 上的 [Excel 加载项 JS WoodGrove 支出趋势](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)。

## <a name="see-also"></a>另请参阅

- [Office 加载项中的 Fabric Core](fabric-core.md)
- [适用于 Office 加载项的 UX 设计模式](../design/ux-design-pattern-templates.md)
