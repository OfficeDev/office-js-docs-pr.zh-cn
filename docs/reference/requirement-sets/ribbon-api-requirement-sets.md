---
title: 功能区 API 要求集
description: 指定哪些 Office 平台和生成支持动态功能区 Api。
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 878670367b253fa7700434681244b43b9cfa36a7
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996513"
---
# <a name="ribbon-api-requirement-sets"></a>功能区 API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

功能区 API 集支持编程控制何时自定义外接程序命令 (即启用和禁用自定义功能区按钮和菜单项) 。

Office 外接程序在多个 Office 版本中运行。 下表列出了功能区 API 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序的内部版本号或版本号。

|  要求集  | Windows 版 Office 2013<br>（一次性购买） | Windows 上的 Office 2016 或更高版本<br>（一次性购买）   | Windows 版 Office\*<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office\*<br>（关联至 Microsoft 365 订阅）  | Office 网页版\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1。1  | 不适用 | 不适用 | 请参阅支持<br>部分 | 不适用 | 16.38 | 2020年11月 | 不适用|

> **&#42;** 仅在 Excel 中支持功能区 API，并且它需要 Microsoft 365 订阅。

## <a name="office-on-windows-subscription-support"></a>Office on Windows (订阅) 支持

使用者通道版本 2006 (版本13001.20498 或更高版本中支持的要求集) 。 对于 Windows 上的 Office，Semi-Annual 通道中也支持该功能，并且每月14月14日（2020或更高版本）中都支持该功能。 每个频道支持的最低版本如下所示：  

|频道 | 版本 | 内部版本|
|:-----|:-----|:-----|
|当前频道 | 2006或更高版本 | 20266.20266 或更高版本|
|月度企业版频道 | 2005或更高版本 | 12827.20538 或更高版本|
|每月企业频道 | 2004 | 12730.20602 或更高版本|
|半年企业频道 | 2002或更高版本 | 12527.20880 或更高版本|

## <a name="more-information"></a>更多信息

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的 Microsoft 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [我使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Microsoft 365 客户端应用程序的版本和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> **RibbonApi 1.1** 要求集在清单中尚不受支持，因此不能在清单的部分中指定它 `<Requirements>` 。


## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="ribbon-api-11"></a>功能区 API 1。1

功能区 API 1.1 是 API 的第一个版本。 有关 API 的详细信息，请参阅 " [Office. 功能区 ](/javascript/api/office/office.ribbon) 参考" 主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 应用程序和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 加载项 XML 清单](/office/dev/add-ins/develop/add-in-manifests)
