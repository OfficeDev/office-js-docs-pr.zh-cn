---
title: Office 加载项的 Office UI 元素
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 444aca7b75e35ef502075876a7d1324fcdca0603
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388099"
---
# <a name="office-ui-elements-for-office-add-ins"></a>Office 加载项的 Office UI 元素

可以使用几种类型的 UI 元素来扩展 Office UI，包括外接程序命令和 HTML 容器。这些 UI 元素类似于 Office 的自然扩展，并且跨平台工作。可以将基于 Web 的自定义代码插入上述任一元素。

下图显示了可以创建的 Office UI 元素的类型。

![在 Office 文档的功能区、任务窗格和对话框上显示外接程序命令的图像](../images/overview-with-app-interface-elements.png)

## <a name="add-in-commands"></a>加载项命令

使用[外接程序命令](add-in-commands.md)将外接程序的入口点添加到 Office 功能区。命令通过运行 JavaScript 代码，或启动 HTML 容器开始在外接程序中操作。可以创建以下两种类型的外接程序命令。

|**命令类型**|**说明**|
|:---------------|:--------------|
|功能区按钮、菜单和选项卡|用于在 Office 的默认功能区中添加自定义按钮、菜单（下拉列表）或选项卡。使用 Office 中的按钮和菜单触发某一操作。使用选项卡对按钮和菜单进行分组和整理。|
|上下文菜单| 用于扩展默认上下文菜单。当用户用鼠标右键单击 Office 文档中的文本或 Excel 中的表时，将显示上下文菜单。| 

## <a name="html-containers"></a>HTML 容器

使用 HTML 容器在 Office 客户端中嵌入基于 HTML 的 UI 代码。然后，这些网页可以引用 Office JavaScript API 以与文档中的内容进行交互。可以创建三种类型的 HTML 容器。

|**HTML 容器**|**说明**|
|:-----------------|:--------------|
|[任务窗格](task-pane-add-ins.md)|在 Office 文档右侧窗格中显示自定义 UI。使用任务窗格以便用户与 Office 文档并行的外接程序进行交互。|
|[内容加载项](content-add-ins.md)|显示 Office 文档内嵌入的自定义 UI。使用内容外接程序以便用户直接与 Office 文档中的外接程序进行交互。例如，你可能想要显示外部内容，如其他来源的视频或数据可视化。 |
|[对话框](dialog-boxes.md)|在覆盖 Office 文档的对话框中显示自定义 UI。对需要焦点和更多空间的交互，但不需要与文档进行并行交互的交互使用对话框。|

## <a name="see-also"></a>另请参阅

- [Excel、Word 和 PowerPoint 加载项命令](add-in-commands.md)
- [任务窗格](task-pane-add-ins.md)
- [内容外接程序](content-add-ins.md)
- [对话框](dialog-boxes.md)
