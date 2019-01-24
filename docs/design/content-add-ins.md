---
title: 内容 Office 加载项
description: 内容加载项是指可以直接嵌入 Excel 或 PowerPoint 文档的图面，用户可以通过它访问界面控件，运行代码以修改文档或显示数据源中的数据。
ms.date: 12/13/2018
localization_priority: Priority
ms.openlocfilehash: 8114fe1d95739ba14fe54c379e44148f451010b2
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389394"
---
# <a name="content-office-add-ins"></a>内容 Office 加载项

内容加载项是指可以直接嵌入 Excel 或 PowerPoint 文档的图面。 用户可以通过内容加载项访问界面控件，运行代码以修改文档或显示数据源中的数据。 在你要将功能直接嵌入文档时，请使用内容加载项。  

*图 1. 内容加载项的典型布局*

![显示内容加载项的典型布局的示例图像。](../images/overview-with-app-content.png)

## <a name="best-practices"></a>最佳做法

- 在加载项顶部包括某些导航或命令元素，如命令栏或透视。
- 包括位于加载项底部的品牌元素，如品牌栏（仅适用于 Excel 和 PowerPoint 加载项）。

## <a name="variants"></a>变量

Office 桌面版和 Office 365 中的 Excel 和 PowerPoint 的内容加载项大小由用户指定。

## <a name="personality-menu"></a>“个性”菜单

“个性”菜单可能会妨碍靠近外接程序右上角的导航和命令元素。以下是 Windows 和 Mac 上的“个性”菜单的当前尺寸。

对于 Windows，个性菜单尺寸为 12x32 像素，如下所示。

*图 2：Windows 上的个性菜单* 

![显示 Windows 桌面上个性菜单的图像](../images/personality-menu-win.png)


对于 Mac，“个性”菜单尺寸为 26x26 像素，但是从右侧浮动 8 个像素，再从顶部浮动 6 个像素，能将占用空间增加至 34x32 像素，如下所示。

*图 3：Mac 上的个性菜单*

![显示 Mac 桌面上个性菜单的图像](../images/personality-menu-mac.png)

## <a name="implementation"></a>实现

有关实现内容加载项的示例，请参阅 GitHub 上的 [Excel 内容加载项 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)。

## <a name="support-considerations"></a>支持注意事项
- 检查 Office 加载项是否适用于[特定 Office 主机平台](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)。 
- 一些内容加载项可能会要求用户“信任”加载项对 Excel 或 PowerPoint 执行读取和写入操作。 可以在加载项清单中声明要拥有的[权限级别](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)。  
- Office 2013 版本及更高版本中的 Excel 和 PowerPoint 支持内容加载项。 如果在不支持 Office Web 加载项的 Office 版本中打开加载项，加载项会显示为图像。

## <a name="see-also"></a>另请参阅
- [Office 加载项主机和平台可用性](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)
- [Office 加载项中的 Office UI Fabric](https://docs.microsoft.com/office/dev/add-ins/design/office-ui-fabric) 
- [Office 加载项的用户体验设计模式](https://docs.microsoft.com/office/dev/add-ins/design/ux-design-pattern-templates)
- [在内容加载项和任务窗格加载项中请求获取 API 使用权限](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
