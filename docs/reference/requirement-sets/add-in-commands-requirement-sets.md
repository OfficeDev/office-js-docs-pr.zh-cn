---
title: 加载项命令要求集
description: Office 外接加载项命令要求集概述
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d1904d092988a445be3e481123eecbad39097764
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094412"
---
# <a name="add-in-commands-requirement-sets"></a>加载项命令要求集

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](../../design/add-in-commands.md) and [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md).

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office host applications that support the initial release version, and the build versions or number for those applications.  

| 发布   |  Windows 版 Office 2013<br>（一次性购买） | Windows 版 Office 2016<br>（一次性购买） | Windows 版 Office 2019<br>（一次性购买） | Windows 版 Office<br> (连接到 Microsoft 365 订阅)    |  iPad 版 Office<br> (连接到 Microsoft 365 订阅)   |  Mac 版 Office<br> (连接到 Microsoft 365 订阅)   | Office 网页版  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| 加载项命令（初始版本，无要求集） | 不适用 | *仅 Outlook 支持* 16.0.4678.1000 | 版本 1809（内部版本 10827.20150）或更高版本 |版本 1603（内部版本 6769.0000）或更高版本 | 不适用 | 15.33 或更高版本| 2016 年 1 月 |

外接程序命令 1.1 要求集介绍了[随文档自动打开任务窗格](../../develop/automatically-open-a-task-pane-with-a-document.md)的功能。

下表列出了外接程序命令 1.1 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。

|  要求集  |  Windows 版 Office 2013<br>（一次性购买） | Windows 版 Office 2016<br>（一次性购买） | Windows 版 Office 2019<br>（一次性购买） | Windows 版 Office<br> (连接到 Microsoft 365 订阅)    |  iPad 版 Office<br> (连接到 Microsoft 365 订阅)   |  Mac 版 Office<br> (连接到 Microsoft 365 订阅)   | Office 网页版  |  
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
