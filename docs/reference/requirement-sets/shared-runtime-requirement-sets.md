---
title: 共享运行时要求集
description: 指定支持 SharedRuntime Api 的平台和 Office 应用程序。
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 872277488dd8d26241d9b445200f429aa102e26e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293462"
---
# <a name="shared-runtime-requirement-sets"></a>共享运行时要求集

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集或使用运行时检查来确定 Office 应用程序是否支持加载项所需的 Api。 有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

运行 JavaScript 代码（例如任务窗格、从外接程序命令启动的函数文件和 Excel 自定义函数）的 Office 外接程序的各个部分可以共享单个 JavaScript 运行时。 这使所有部分都可以共享一组全局变量，共享一组已加载库，并且可以相互通信，而无需通过持久化存储传递邮件。

下表列出了 SharedRuntime 1.1 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序的内部版本号或版本号。

|  要求集  |  Windows 上的 Office 2013 (或更高版本) <br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅）   |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1。1  | 无 | 版本 2002 (内部版本 12527.20092) 或更高版本 | 无 | 16.35 或更高版本 | 2020 年 2 月 | 无 |

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)
