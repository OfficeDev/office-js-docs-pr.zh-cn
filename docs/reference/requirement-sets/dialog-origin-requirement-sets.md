---
title: 对话框源要求集
description: 了解有关 Dialog Origin 要求集的详细信息。
ms.date: 07/22/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 20cd797cfcc0c8b673f8531fbe71769196d0d274ea58a9b76171c1427c125aba
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086623"
---
# <a name="dialog-origin-requirement-sets"></a>对话框源要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

Office 外接程序在多个 Office 版本中运行。 下表列出了 Dialog Origin 要求集、Office要求集的客户端应用程序，以及该对话框应用程序Office版本号。

|  要求集  | Windows 版 Office 2013<br>（一次性购买） | Windows 版 Office 2016<br>（一次性购买） | Office 2019 年 10 月或Windows<br>（一次性购买） | Windows 版 Office<br> (订阅)  |  iPad 版 Office<br> (订阅)   |  Mac 版 Office<br> (订阅)   | Office 网页版  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | 内部版本<br>15.0.5371.1000<br>或更高版本 | 内部版本<br>16.0.5200.1000<br>或更高版本 | 内部版本<br>待定<br>或更高版本 | 待定 | 2.52 或更高版本 | 16.52 或更高版本 | 2021 年 7 月 | 版本 2108<br> (内部版本 10377.1000) <br>或更高版本 |

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="dialog-origin-11"></a>Dialog Origin 1.1

Dialog Origin 1.1 是 API 的第一个版本。 它为对话框及其父页面之间的跨域消息提供支持。 有关这些 API 的详细信息，请参阅[Office.ui](/javascript/api/office/office.ui)参考主题。

## <a name="see-also"></a>另请参阅

- [在 Office 加载项中使用 Office 对话框 API](../../develop/dialog-api-in-office-add-ins.md)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
