---
title: Office 加载项中的任务窗格
description: 任务窗格允许用户访问界面控件，此类控件运行代码以修改文档或电子邮件，或显示数据源中的数据。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39a96f4d5aa63d55f4dcb30d9aeb9e680357aa09
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093754"
---
# <a name="task-panes-in-office-add-ins"></a>Office 加载项中的任务窗格
 
Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*图 1：典型任务窗格布局*

![显示典型任务窗格布局的图像](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>最佳做法

|**允许事项**|**禁止事项**|
|:-----|:--------|
|<ul><li>在标题中包括外接程序的名称。</li></ul>|<ul><li>请勿在标题中追加公司名称。</li></ul>|
|<ul><li>在标题中使用简短的描述性名称。</li></ul>|<ul><li>不要将字符串（例如 "外接程序"、"for Word" 或 "for Office"）追加到外接程序的标题。</li></ul>|
|<ul><li>在加载项顶部包括某些导航或命令元素，如命令栏或透视。</li></ul>||
|<ul><li>在外接程序底部包括品牌元素，如品牌栏，除非要在 Outlook 内使用外接程序。</li></ul>||


## <a name="variants"></a>变量

The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.  

*图 2：Office 2016 桌面任务窗格尺寸*

![显示尺寸为 1366x768 的桌面任务窗格的图像](../images/office-2016-taskpane-sizes.png)

- Excel - 320 x 455
- PowerPoint - 320 x 531
- Word - 320 x 531
- Outlook - 348x535

<br/>

*图3。Office 任务窗格大小*

![显示尺寸为 1366x768 的桌面任务窗格的图像](../images/office-365-taskpane-sizes.png)

- Excel - 350 x 378
- PowerPoint - 348x391
- Word - 329 x 445
- Outlook（网页版）- 320x570

## <a name="personality-menu"></a>“个性”菜单

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

对于 Windows，个性菜单尺寸为 12x32 像素，如下所示。

*图 4：Windows 上的个性菜单*

![显示 Windows 桌面上个性菜单的图像](../images/personality-menu-win.png)

对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将空间增加至 34x32 像素，如下所示。

*图 5：Mac 上的个性菜单*

![显示 Mac 桌面上个性菜单的图像](../images/personality-menu-mac.png)

## <a name="implementation"></a>实现

有关实现任务窗格的示例，请参阅 GitHub 上的 [Excel 加载项 JS WoodGrove 支出趋势](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)。 


## <a name="see-also"></a>另请参阅

- [Office 加载项中的 Office UI Fabric](office-ui-fabric.md) 
- [适用于 Office 外接程序的 UX 设计模式](../design/ux-design-pattern-templates.md)

