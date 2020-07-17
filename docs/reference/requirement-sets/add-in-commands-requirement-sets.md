---
title: 加载项命令要求集
description: Office 加载项命令要求集概述。
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 14952f97e3f03b83c15d0a4d51183fac3ec80af3
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159085"
---
# <a name="add-in-commands-requirement-sets"></a>加载项命令要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

外接程序命令是 UI 元素，可扩展 Office UI，并在外接程序中启动操作。可以使用加载项命令在功能区上添加按钮，也可以向上下文菜单添加项。有关详细信息，请参阅 [Excel、Word 和 PowerPoint 的加载项命令](../../design/add-in-commands.md)和 [Outlook 的加载项命令](../../outlook/add-in-commands-for-outlook.md)。

外接程序命令的初始版本没有相应的要求集（即，没有 AddinCommands 1.0 要求集）。下表列出了支持初始版本的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。  

| 发布   |  Windows 版 Office 2013<br>（一次性购买） | Windows 版 Office 2016<br>（一次性购买） | Windows 版 Office 2019<br>（一次性购买） | Windows 版 Office<br>（连接到 Microsoft 365 订阅）   |  iPad 版 Office<br>（连接到 Microsoft 365 订阅）  |  Mac 版 Office<br>（连接到 Microsoft 365 订阅）  | Office 网页版  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| 加载项命令（初始版本，无要求集） | 不适用 | *仅 Outlook 支持* 16.0.4678.1000 | 版本 1809（内部版本 10827.20150）或更高版本 |版本 1603（内部版本 6769.0000）或更高版本 | 不适用 | 15.33 或更高版本| 2016 年 1 月 |

外接程序命令 1.1 要求集介绍了[随文档自动打开任务窗格](../../develop/automatically-open-a-task-pane-with-a-document.md)的功能。

下表列出了外接程序命令 1.1 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。

|  要求集  |  Windows 版 Office 2013<br>（一次性购买） | Windows 版 Office 2016<br>（一次性购买） | Windows 版 Office 2019<br>（一次性购买） | Windows 版 Office<br>（连接到 Microsoft 365 订阅）   |  iPad 版 Office<br>（连接到 Microsoft 365 订阅）  |  Mac 版 Office<br>（连接到 Microsoft 365 订阅）  | Office 网页版  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.1  | 不适用 | *仅 Outlook 支持* 16.0.4678.1000  | 版本 1809（内部版本 10827.20150）或更高版本 | 版本 1705（内部版本 8121.1000）或更高版本 | 不适用 | 15.34 或更高版本\*| 2017 年 5 月 |

>\*针对版本 16.9 &ndash; 16.14（含），[Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) 方法将错误地返回 `false`，但这些版本*支持*需求集。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 主机和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)
