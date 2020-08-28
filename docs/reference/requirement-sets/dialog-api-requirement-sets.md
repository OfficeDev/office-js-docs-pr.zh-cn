---
title: Dialog API 要求集
description: 了解有关对话框 API 要求集的详细信息
ms.date: 08/20/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 2056d2e55ad868d03b3dc0af0e6d30cd6207994c
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293553"
---
# <a name="dialog-api-requirement-sets"></a>Dialog API 要求集

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集或使用运行时检查来确定 Office 应用程序是否支持加载项所需的 Api。 有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

Office 外接程序在多个 Office 版本中运行。 下表列出了对话框 API 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序的内部版本号或版本号。

|  要求集  | Windows 版 Office 2013\*<br>（一次性购买） | Windows 上的 Office 2016 或更高版本\*<br>（一次性购买）   | Windows 版 Office<br> (订阅)  |  iPad 版 Office<br> (订阅)   |  Mac 版 Office<br> (订阅)   | Office 网页版  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | 生成号 15.0.4855.1000 或更高版本 | 生成号 16.0.4390.1000 或更高版本 | 版本 1602（生成号 6741.0000）或更高版本 | 1.22 或更高版本 | 15.20 或更高版本 | 2017 年 1 月 | 版本 1608（内部版本 7601.6800）或更高版本|
| DialogApi 1。2  | 不适用 | 不适用 | 请参阅支持<br>部分 | 2.67 或更高版本 | 16.37 或更高版本 | 2020 年 6 月 | 无 |

>\* 一次性购买 Office 的用户可能未接受所有修补和更新。 如果是这样，即使在用户的计算机上未安装支持 DialogApi 所需的更新的 Dll，Office 用来在 UI 中报告其版本的 DLL 可能也会大于此处列出的版本。 若要确保安装了所需的修补程序，用户必须转到 Office 更新列表 ([office 2013 列表](/officeupdates/msp-files-office-2013) 或 [office 2016 列表](/officeupdates/msp-files-office-2016) ") ，搜索 **osfclient-x**，并安装列出的修补程序。

## <a name="office-on-windows-subscription-support"></a>Office on Windows (订阅) 支持

DialogApi 1.2 要求集在消费者频道版本2005中受支持， (版本、12827.20268 或更高) 。 对于 Windows 上的 Office，在半年频道中也支持此功能，每月9日、2020年6月或更高版本的企业版频道内部版本都支持该功能。 每个频道支持的最低版本如下所示：  

|频道 | 版本 | 内部版本|
|:-----|:-----|:-----|
|当前频道 | 2005或更高版本 | 12827.20160 或更高版本|
|月度企业版频道 | 2004或更高版本 | 12730.20430 或更高版本|
|半年企业频道 | 2002或更高版本 | 12527.20720 或更高版本|

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="dialog-api-11-and-12"></a>对话框 API 1.1 和1。2

Dialog API 1.1 是首版 API。 版本1.2 添加了对使用方法将父页面中的数据发送到对话框的支持 `Office.ui.messageChild` 。 有关这些 Api 的详细信息，请参阅 [对话框 API](/javascript/api/office/office.ui) 参考主题。

## <a name="see-also"></a>另请参阅

- [在 Office 加载项中使用 Office 对话框 API](../../develop/dialog-api-in-office-add-ins.md)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)
