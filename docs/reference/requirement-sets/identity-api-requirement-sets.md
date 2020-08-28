---
title: Identity API 要求集
description: Office 外接程序的标识 API 要求集信息。
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c2c6ea449cef08248a9ba79051b7c0c5f9baa600
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293539"
---
# <a name="identity-api-requirement-sets"></a>Identity API 要求集

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集或使用运行时检查来确定 Office 应用程序是否支持加载项所需的 Api。 有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

Office 外接程序在多个 Office 版本中运行。 下表列出了标识 API 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序的内部版本号或版本号。

|  要求集  | Windows 上的 Office 2013 或更高版本<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1。3  | 无 | 2008 (内部版本 13127.20000) 或更高版本 | 即将推出 | 16.40 或更高版本 | 2020年8月 * |

> \* 最初，只有从 SharePoint Online 和 OneDrive.com 打开的文档才会在 web 上的 Office 中支持要求集。 对其他文档的支持稍后将在2020中向 Office 提供。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="identityapi-preview"></a>IdentityAPI 预览

有关此 API 的详细信息，请参阅在 [tokenhelper.getaccesstoken 以便](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) 中使用承诺的版本或在 [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)中使用回调的版本。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)
