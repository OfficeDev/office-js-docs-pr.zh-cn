---
title: 功能区 API 要求集
description: 指定哪些Office和内部版本支持动态功能区 API。
ms.date: 05/12/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 18b081da28991eb15875c6999ddfe06ccea13eade8dd8a4e1f2a4cc09ac5e922
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084919"
---
# <a name="ribbon-api-requirement-sets"></a>功能区 API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

功能区 API 集支持以编程方式控制自定义外接程序命令 (，即自定义功能区按钮和菜单项) 和禁用。

Office 外接程序在多个 Office 版本中运行。 下表列出了功能区 API 要求集、Office要求集的客户端应用程序，以及功能区 API 要求集Office版本号。

|  要求集  | Windows 版 Office 2013<br>（一次性购买） | Office 2016 或更高版本Windows<br>（一次性购买）   | Windows 版 Office\*<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office\*<br>（关联至 Microsoft 365 订阅）  | Office 网页版\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | 不适用 | 不适用 | 请参阅支持<br>部分如下 | 不适用 | 16.38 | 2020 年 11 月 | 不适用|
| RibbonApi 1.2  | 不适用 | 不适用 | 2102 (内部版本 13801.20294)  | 不适用 | 即将推出 | 2021 年 5 月 | 不适用|

> **&#42;** 功能区 API 仅在 Excel并且需要一个Microsoft 365订阅。

## <a name="support-for-version-11-on-office-on-windows-subscription"></a>支持版本 1.1（Office订阅Windows (版本) 

1.1 版本的 RibbonApi 要求集在消费者频道版本 2006 (版本 13001.20498 或) 。 For Office on Windows the feature is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available july 14， 2020 or later. 每个频道支持的最低版本如下所示：  

|频道 | 版本 | 内部版本|
|:-----|:-----|:-----|
|当前频道 | 2006 或更大 | 20266.20266 或更大|
|每月企业频道 | 2005 或更大 | 12827.20538 或更大|
|每月企业频道 | 2004 | 12730.20602 或更大|
|半年企业频道 | 2002 或更大 | 12527.20880 或更大|

## <a name="more-information"></a>更多信息

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [适用于客户端的更新频道版本的版本号Microsoft 365版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到客户端应用程序的版本号Microsoft 365版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="ribbon-api-11"></a>功能区 API 1.1

功能区 API 1.1 是首版 API。 有关 API 的详细信息，请参阅[Office.ribbon](/javascript/api/office/office.ribbon)参考主题。

## <a name="ribbon-api-12"></a>功能区 API 1.2

功能区 API 1.2 增加了对上下文选项卡的支持。 更多信息，请参见[在Office插件中创建自定义上下文选项卡](../../design/contextual-tabs.md)。

> [!NOTE]
> **由于清单中尚不支持 RibbonApi 1.2** 要求集，因此不应在清单的部分中指定 `<Requirements>` 它。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
