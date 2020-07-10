---
title: Identity API 要求集
description: Office 外接程序的标识 API 要求集信息。
ms.date: 04/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 5bface00e0ffe89e7a403b251129867b334f7f69
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094377"
---
# <a name="identity-api-requirement-sets"></a>Identity API 要求集

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Office 外接程序在多个 Office 版本中运行。 下表列出了 Identity API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。

|  要求集  | Windows 上的 Office 2013 或更高版本<br>（一次性购买） | Windows 版 Office<br> (连接到 Microsoft 365 订阅)  |  iPad 版 Office<br> (连接到 Microsoft 365 订阅)   |  Mac 版 Office<br> (连接到 Microsoft 365 订阅)   | Office 网页版  | SharePoint Online | OneDrive.com |Outlook.com & Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 预览  | 不适用 | 预览<b>*</b> | 即将推出 | 预览<b>*</b> | 预览<b>* &#8224;</b> | 预览<b>* &#8224;</b>| 即将推出 | 即将推出 |

> **&#42;** 在预览阶段，标识 API 需要 Microsoft 365 订阅。 你应该使用来自预览体验成员频道的最新每月版本和内部版本。 你可能需要成为 Office 预览体验成员，才能获取此版本。 有关详细信息，请参阅[成为 Office 预览体验成员](https://insider.office.com)。 请注意，当内部版本进入生产半年频道时，将关闭对该内部版本的预览功能（包括 SSO）的支持。
>
> **&#8224;** 在这些平台上使用 SSO Api 的外接程序将仅在用户的租户管理员向外接程序授予许可时才有效。 用户不能同意，即使是对自己的 Azure AD 配置文件也不允许。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="identityapi-preview"></a>IdentityAPI 预览

有关此 API 的详细信息，请参阅在[tokenhelper.getaccesstoken 以便](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-)中使用承诺的版本或在[getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)中使用回调的版本。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 主机和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)
