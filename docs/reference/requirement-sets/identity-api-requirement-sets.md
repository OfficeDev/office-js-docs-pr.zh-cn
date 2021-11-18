---
title: Identity API 要求集
description: 加载项的标识 API 要求Office信息。
ms.date: 11/16/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: d953e3ca2d135b96ab8b3219d9fe0f52fbda9d99
ms.sourcegitcommit: 6e6c4803fdc0a3cc2c1bcd275288485a987551ff
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2021
ms.locfileid: "61066714"
---
# <a name="identity-api-requirement-sets"></a>Identity API 要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

Office 外接程序在多个 Office 版本中运行。 下表列出了 Identity API 要求集、Office要求集的客户端应用程序，以及该标识 API 要求集Office版本号。

|  要求集  | Office 2021 或更高版本Windows<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | 内部版本 16.0.14326.20454 或更高版本 | 版本 2008 (内部版本 13127.20000) 或更高版本 | 不支持 | 16.40 或更高版本 | Microsoft Office SharePoint Online 和 OneDrive\* |

\*目前，要求集仅在 Office web 版 中支持，仅适用于从 Microsoft Office SharePoint Online 和 OneDrive。

## <a name="outlook-and-identity-api-requirement-sets"></a>Outlook和标识 API 要求集

若要要求在加载项代码中将 Identity API 设置为 1.3 Outlook，请通过调用 检查它是否受支持 `isSetSupported('IdentityAPI', '1.3')` 。 不支持在Outlook清单中声明它。 还可通过检查其不是 `undefined` 来确定该 API 是否受到支持。 有关详细信息，请参阅 [从后续要求集中使用 API](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)。

> [!NOTE]
> 在使用Outlook激活的 Outlook 外接程序中，Windows 版本 2108 (版本 14326.20258) 或更高版本上的 Office 支持[OfficeRuntime.Auth](/javascript/api/office-runtime/officeruntime.auth)接口。 [Office。版本](/javascript/api/office/office.auth)2109 和内部版本 14425.10000 (版本 14425.10000) 支持身份验证接口。 有关版本的详细信息，请参阅[Office 2021](/officeupdates/update-history-office-2021)或[Microsoft 365](/officeupdates/update-history-office365-proplus-by-date)的更新历史记录页以及如何查找 Office[客户端版本和更新通道](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
