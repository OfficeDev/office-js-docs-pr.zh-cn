---
title: 加载项命令要求集
description: Office 加载项命令要求集概述。
ms.date: 11/01/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 08fcb5df0e614e9b9f3ec9479fc958cc79adf320
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087957"
---
# <a name="add-in-commands-requirement-sets"></a>加载项命令要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

外接程序命令是 UI 元素，可扩展 Office UI，并在外接程序中启动操作。可以使用加载项命令在功能区上添加按钮，也可以向上下文菜单添加项。有关详细信息，请参阅 [Excel、Word 和 PowerPoint 的加载项命令](../../design/add-in-commands.md)和 [Outlook 的加载项命令](../../outlook/add-in-commands-for-outlook.md)。

外接程序命令的初始发布没有相应的要求集 (也就是说，没有 Addincommand 1.0 要求集) 。 下表列出了支持初始发布版本的 Office 客户端应用程序，以及这些应用程序的内部版本或编号。  

| 发布   |  Windows 版 Office 2013<br>（一次性购买） | Windows 版 Office 2016<br>（一次性购买） | Windows 版 Office 2019<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅）   |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| 加载项命令（初始版本，无要求集） | 不适用 | *仅 Outlook 支持* 16.0.4678.1000 | 版本 1809（内部版本 10827.20150）或更高版本 |版本 1603（内部版本 6769.0000）或更高版本 | 不适用 | 15.33 或更高版本| 2016 年 1 月 |

加载项命令 **1.1** 要求集引入了 [autoopen 任务窗格与文档](../../develop/automatically-open-a-task-pane-with-a-document.md)的功能。

外接程序命令 **1.3** 要求集引入了清单标记，该标记使外接程序能够自定义 Office 功能区上的自定义选项卡的位置，并将内置的 Office 功能区控件插入到自定义控件组中。

下表列出了加载项命令要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序的内部版本号或版本号。

|  要求集  |  Windows 版 Office 2013<br>（一次性购买） | Windows 版 Office 2016<br>（一次性购买） | Windows 版 Office 2019<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅）   |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Addincommand 1。3  | 不适用 | 不适用  | 不适用 | 即将推出 | 无 | 即将推出 | 2020 年 11 月 |
| AddinCommands 1.1  | 不适用 | *仅 Outlook 支持* 16.0.4678.1000  | 版本 1809（内部版本 10827.20150）或更高版本 | 版本 1705（内部版本 8121.1000）或更高版本 | 不适用 | 15.34 或更高版本\*| 2017 年 5 月 |

>\*针对版本 16.9 &ndash; 16.14（含），[Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) 方法将错误地返回 `false`，但这些版本 *支持* 需求集。

> [!IMPORTANT]
> Addincommand 1.3 处于预览阶段， *仅适用于 web 上的 PowerPoint*。 我们建议您仅在测试和开发环境中尝试标记。 请勿在生产环境中或在业务关键型文档中使用预览标记。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
