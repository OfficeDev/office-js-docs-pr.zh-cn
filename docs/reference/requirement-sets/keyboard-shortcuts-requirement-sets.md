---
title: 键盘快捷方式要求集
description: 键盘快捷方式要求集信息Office外接程序。
ms.date: 02/07/2022
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 074460510e054cdcfbeca4676883c4180bb2202d
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467841"
---
# <a name="keyboard-shortcuts-requirement-sets"></a>键盘快捷方式要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

Office 外接程序在多个 Office 版本中运行。 下表列出了键盘快捷方式要求集、Office要求集的客户端应用程序，以及该键盘快捷方式要求集Office版本号。

|  要求集  | Windows 上的 Office 2013 或更高版本<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| KeyboardShortcuts 1.1  | 不适用 | 版本：2111 (内部版本 14701.10000)  | 不适用 | 16.55 | 2021 年 9 月 |

> [!NOTE]
> **KeyboardShortcuts 1.1** 要求集仅在 Excel。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="keyboardshortcuts-11"></a>KeyboardShortcuts 1.1

有关此要求集内 API 的详细信息，请参阅 [Office.actions](/javascript/api/office/office.actions)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
