---
title: Identity API 要求集
description: 加载项的标识 API 要求Office信息。
ms.date: 01/26/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 40f9154b06b0e373ac367b0dcc24cbfc05be4856
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671427"
---
# <a name="identity-api-requirement-sets"></a>Identity API 要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

Office 外接程序在多个 Office 版本中运行。 下表列出了 Identity API 要求集、Office要求集的客户端应用程序，以及该标识 API 要求集Office版本号。

|  要求集  | Windows 上的 Office 2013 或更高版本<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | 不适用 | 2008 (版本 13127.20000) 或更高版本 | 即将推出 | 16.40 或更高版本 | Microsoft Office SharePoint Online 和 OneDrive\* |

\*目前，要求集仅在Office web 版和文档打开的文档Microsoft Office SharePoint Online OneDrive。

> [!NOTE]
> Outlook：若要要求在加载项代码中将 Identity API 设置为 1.3，请通过调用 检查是否受支持 `isSetSupported('IdentityAPI', '1.3')` 。 不支持在Outlook清单中声明它。 还可通过检查其不是 `undefined` 来确定该 API 是否受到支持。 有关详细信息，请参阅 [从后续要求集中使用 API](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="identityapi-preview"></a>IdentityAPI 预览

有关此 API 的详细信息，请参阅 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) 处使用 Promises 的版本，或者使用 [getAccessTokenAsync](/javascript/api/office/office.auth#getAccessTokenAsync_options__callback_)的回调的版本。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求集](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
