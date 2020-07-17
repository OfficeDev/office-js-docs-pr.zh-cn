---
title: Dialog API 要求集
description: 了解有关对话框 API 要求集的详细信息。
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: f53bd5c62c434c361d435eb51035e45079f8e429
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159099"
---
# <a name="dialog-api-requirement-sets"></a>Dialog API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

Office 外接程序在多个 Office 版本中运行。下表列出了 Dialog API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。

|  要求集  | Windows 版 Office 2013\*<br>（一次性购买） | Windows 上的 Office 2016 或更高版本\*<br>（一次性购买）   | Windows 版 Office<br>（连接到 Microsoft 365 订阅） |  iPad 版 Office<br>（连接到 Microsoft 365 订阅）  |  Mac 版 Office<br>（连接到 Microsoft 365 订阅）  | Office 网页版  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | 生成号 15.0.4855.1000 或更高版本 | 生成号 16.0.4390.1000 或更高版本 | 版本 1602（生成号 6741.0000）或更高版本 | 1.22 或更高版本 | 15.20 或更高版本| 2017 年 1 月 | 版本 1608（内部版本 7601.6800）或更高版本|

>\*一次性购买 Office 的用户可能未接受所有修补和更新。 如果是这样，即使在用户的计算机上未安装支持 DialogApi 所需的更新的 Dll，Office 用来在 UI 中报告其版本的 DLL 可能也会大于此处列出的版本。 若要确保安装了所需的修补程序，用户必须转到 Office 更新列表（[office 2013 列表](/officeupdates/msp-files-office-2013)或[office 2016 列表](/officeupdates/msp-files-office-2016)），搜索**osfclient-x**，并安装列出的修补程序。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="dialog-api-11"></a>Dialog API 1.1

Dialog API 1.1 是首版 API。 有关 API 的详细信息，请参阅[对话框 API](/javascript/api/office/office.ui)参考主题。

## <a name="see-also"></a>另请参阅

- [在 Office 加载项中使用 Office 对话框 API](../../develop/dialog-api-in-office-add-ins.md)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 主机和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)
