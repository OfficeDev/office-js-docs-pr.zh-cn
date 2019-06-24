---
title: Dialog API 要求集
description: ''
ms.date: 06/20/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 3135c65120248194603b91510450519f106e0ad1
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127119"
---
# <a name="dialog-api-requirement-sets"></a>Dialog API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Office 外接程序在多个 Office 版本中运行。下表列出了 Dialog API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。

|  要求集  | Windows 版 Office 2013<br>（一次性购买） | Windows 上的 Office 2016 或更高版本<br>（一次性购买）   | Windows 版 Office<br>(连接到 Office 365 订阅) |  IPad 上的 Office<br>(连接到 Office 365 订阅)  |  Mac 上的 Office<br>(连接到 Office 365 订阅)  | 网上的 Office  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | 生成号 15.0.4855.1000 或更高版本 | 生成号 16.0.4390.1000 或更高版本 | 版本 1602（生成号 6741.0000）或更高版本 | 1.22 或更高版本 | 15.20 或更高版本| 2017 年 1 月 | 版本 1608（生成号 7601.6800）或更高版本|

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="dialog-api-11"></a>Dialog API 1.1

Dialog API 1.1 是首版 API。 有关 API 的详细信息，请参阅 [Dialog API](/javascript/api/office/office.ui) 参考主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](/office/dev/add-ins/develop/add-in-manifests)
