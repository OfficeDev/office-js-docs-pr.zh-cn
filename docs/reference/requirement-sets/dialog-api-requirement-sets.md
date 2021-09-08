---
title: Dialog API 要求集
description: 了解有关对话框 API 要求集的详细信息。
ms.date: 07/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: a5e0855cf6e7b7c9c3157fe24bed3610a26992ed
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937749"
---
# <a name="dialog-api-requirement-sets"></a>Dialog API 要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

Office 外接程序在多个 Office 版本中运行。 下表列出了对话框 API 要求集、Office要求集的客户端应用程序，以及该对话框应用程序Office版本号。

|  要求集  | Windows 版 Office 2013\*<br>（一次性购买） | Office 2016 或更高版本Windows\*<br>（一次性购买）   | Windows 版 Office<br> (订阅)  |  iPad 版 Office<br> (订阅)   |  Mac 版 Office<br> (订阅)   | Office 网页版  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2  | 不适用 | 不适用 | 请参阅支持<br>部分如下 | 2.37 或更高版本 | 16.37 或更高版本 | 2020 年 6 月 | 不适用 |
| DialogApi 1.1  | 生成号 15.0.4855.1000 或更高版本 | 生成号 16.0.4390.1000 或更高版本 | 版本 1602（生成号 6741.0000）或更高版本 | 1.22 或更高版本 | 15.20 或更高版本 | 2017 年 1 月 | 版本 1608（内部版本 7601.6800）或更高版本|

>\*一次购买的用户可能Office所有修补程序和更新。 如果是这样，Office在 UI 中报告其版本的 DLL 可能大于此处列出的版本，即使用户计算机上未安装支持 DialogApi 所需的更新 DLL。 若要确保已安装所需的修补程序，用户必须转到 Office 更新列表 ([Office 2013](/officeupdates/msp-files-office-2013)列表或 [Office 2016](/officeupdates/msp-files-office-2016)列表) ，搜索 **osfclient-x-none，** 然后安装列出的修补程序。

## <a name="office-on-windows-subscription-support"></a>Office订阅Windows (支持) 订阅

DialogApi 1.2 要求集在消费者频道版本 2005 (内部版本 12827.20268 或) 。 For Office on Windows， the feature is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available June 9， 2020 or later. 每个频道支持的最低版本如下所示：  

|频道 | 版本 | 内部版本|
|:-----|:-----|:-----|
|当前频道 | 2005 或更大 | 12827.20160 或更大|
|每月企业频道 | 2004 或更大 | 12730.20430 或更大|
|半年企业频道 | 2002 或更大 | 12527.20720 或更大|

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="dialog-api-11-and-12"></a>对话框 API 1.1 和 1.2

Dialog API 1.1 是首版 API。 要求集 1.2 增加了对使用[Office.dialog.messageChild](/javascript/api/office/office.dialog#messageChild_message_)方法将数据从父页面发送到对话框的支持。 有关这些 API 的详细信息，请参阅 [对话框 API](/javascript/api/office/office.ui) 参考主题。

## <a name="see-also"></a>另请参阅

- [在 Office 加载项中使用 Office 对话框 API](../../develop/dialog-api-in-office-add-ins.md)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
