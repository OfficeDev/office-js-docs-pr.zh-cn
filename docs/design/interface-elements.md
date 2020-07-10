---
title: Office 加载项的 Office UI 元素
description: 获取 Office 外接程序中不同种类的 UI 元素的概述。
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 5b9907924c674ed9db2294621123c394419d0c12
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093761"
---
# <a name="office-ui-elements-for-office-add-ins"></a>Office 加载项的 Office UI 元素

You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.

下图显示了可以创建的 Office UI 元素的类型。

![在 Office 文档的功能区、任务窗格和对话框上显示外接程序命令的图像](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a>加载项命令

使用[外接程序命令](add-in-commands.md)将入口点添加到你的外接程序中的 Office 应用功能区。 命令通过运行 JavaScript 代码，或启动 HTML 容器开始在外接程序中操作。 可以创建以下两种类型的外接程序命令。

|**命令类型**|**说明**|
|:---------------|:--------------|
|功能区按钮、菜单和选项卡|Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.|
|上下文菜单| Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.| 

## <a name="html-containers"></a>HTML 容器

Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.

|**HTML 容器**|**说明**|
|:-----------------|:--------------|
|[任务窗格](task-pane-add-ins.md)|Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.|
|[内容加载项](content-add-ins.md)|Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources. |
|[对话框](dialog-boxes.md)|Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.|

## <a name="see-also"></a>另请参阅

- [Excel、Word 和 PowerPoint 加载项命令](add-in-commands.md)
- [任务窗格](task-pane-add-ins.md)
- [内容外接程序](content-add-ins.md)
- [对话框](dialog-boxes.md)
