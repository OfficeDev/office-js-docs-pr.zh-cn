---
title: 功能区 API 要求集
description: 指定支持动态功能区 API 的 Office 平台和内部版本。
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 91c909755779d122fba8d77dc246784f6a0dd1a3
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839983"
---
# <a name="ribbon-api-requirement-sets"></a>功能区 API 要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

功能区 API 集支持以编程方式控制自定义外接程序命令 (，即自定义功能区按钮和菜单项) 和禁用。

Office 外接程序在多个 Office 版本中运行。 下表列出了功能区 API 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序内部版本或版本号。

|  要求集  | Windows 版 Office 2013<br>（一次性购买） | Windows 版 Office 2016 或更高版本<br>（一次性购买）   | Windows 版 Office\*<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office\*<br>（关联至 Microsoft 365 订阅）  | Office 网页版\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | 不适用 | 不适用 | 请参阅支持<br>部分如下 | 无 | 16.38 | 2020 年 11 月 | 无|

> **&#42;** 功能区 API 仅在 Excel 上受支持，并且需要 Microsoft 365 订阅。

## <a name="office-on-windows-subscription-support"></a>Windows 版 Office (订阅) 支持

要求集在消费者频道版本 2006 (版本 13001.20498 或) 。 对于 Windows 版 Office，2020 Semi-Annual 2020 年 7 月 14 日版和每月企业频道版本也支持此功能。 每个频道支持的最低版本如下：  

|频道 | 版本 | 内部版本|
|:-----|:-----|:-----|
|当前频道 | 2006 或更大 | 20266.20266 或更大|
|每月企业频道 | 2005 或更大 | 12827.20538 或更大|
|每月企业频道 | 2004 | 12730.20602 或更大|
|半年企业频道 | 2002 或更大 | 12527.20880 或更大|

## <a name="more-information"></a>更多信息

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [Microsoft 365 客户端更新频道版本的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Microsoft 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> **由于清单中尚不支持 RibbonApi 1.1** 要求集，因此无法在清单的部分中指定 `<Requirements>` 它。


## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="ribbon-api-11"></a>功能区 API 1.1

功能区 API 1.1 是 API 的第一个版本。 有关 API 的详细信息，请参阅 [Office.ribbon ](/javascript/api/office/office.ribbon) 参考主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)