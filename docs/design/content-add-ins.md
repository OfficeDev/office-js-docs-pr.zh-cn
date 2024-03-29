---
title: 内容 Office 加载项
description: 内容加载项是指可以直接嵌入 Excel 或 PowerPoint 文档的图面，用户可以通过它访问界面控件，运行代码以修改文档或显示数据源中的数据。
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: c10893d60f64d875d92aec979a5700630b2cf96c
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810244"
---
# <a name="content-office-add-ins"></a>内容 Office 加载项

内容加载项是指可以直接嵌入 Excel 或 PowerPoint 文档的图面。 用户可以通过内容加载项访问界面控件，运行代码以修改文档或显示数据源中的数据。 在你要将功能直接嵌入文档时，请使用内容加载项。  

*图 1. 内容加载项的典型布局*

![Office 应用程序中内容加载项的典型布局。](../images/overview-with-app-content.png)

## <a name="best-practices"></a>最佳做法

- 在加载项顶部包括某些导航或命令元素，如命令栏或透视。
- 包括位于加载项底部的品牌元素，如品牌栏（仅适用于 Excel 和 PowerPoint 加载项）。

## <a name="variants"></a>变量

Office 桌面版和 Web 浏览器中 Excel 和 PowerPoint 的内容加载项大小由用户指定。

## <a name="personality-menu"></a>“个性”菜单

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

对于 Windows，个性菜单尺寸为 12x32 像素，如下所示。

*图 2：Windows 上的个性菜单*

![Windows 桌面上的 12x32 像素个性菜单。](../images/personality-menu-win.png)

对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将占用空间增加至 34x32 像素，如下所示。

*图 3：Mac 上的个性菜单*

![Mac 桌面上的 34x32 像素个性菜单。](../images/personality-menu-mac.png)

## <a name="implementation"></a>实现

有关实现内容加载项的示例，请参阅 GitHub 上的 [Excel 内容加载项 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)。

## <a name="support-considerations"></a>支持注意事项

- 检查 Office 外接程序是否适用于 [特定的 Office 应用程序或平台](/javascript/api/requirement-sets)。
- 一些内容加载项可能会要求用户“信任”加载项对 Excel 或 PowerPoint 执行读取和写入操作。 可以在加载项清单中声明要拥有的[权限级别](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。  
- Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later. If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.

## <a name="see-also"></a>另请参阅

- [Office 客户端应用程序和平台的 Office 加载项可用性](/javascript/api/requirement-sets)
- [Office 加载项中的 Fabric Core](fabric-core.md)
- [适用于 Office 加载项的 UX 设计模式](../design/ux-design-pattern-templates.md)
- [在加载项中请求获取 API 使用权限](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
