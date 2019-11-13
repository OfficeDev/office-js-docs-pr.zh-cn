---
title: Identity API 要求集
description: ''
ms.date: 11/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 96f5c305f4ecfe0fdc0ee89aed6955e090f87b02
ms.sourcegitcommit: 88d81aa2d707105cf0eb55d9774b2e7cf468b03a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/13/2019
ms.locfileid: "38301923"
---
# <a name="identity-api-requirement-sets"></a>Identity API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Office 外接程序在多个 Office 版本中运行。 下表列出了 Identity API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。

|  要求集  | Windows 上的 Office 2013 或更高版本<br>（一次性购买） | Windows 版 Office<br>（已连接到 Office 365 订阅） |  iPad 版 Office<br>（已连接到 Office 365 订阅）  |  Mac 版 Office<br>（连接到 Office 365 订阅）  | Office 网页版  | SharePoint Online | OneDrive.com |Outlook.com & Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 预览  | 不适用 | 预览<b>*</b> | 即将推出 | 预览<b>*</b> | 预览<b>* &#8224;</b> | 预览<b>* &#8224;</b>| 即将推出 | 即将推出 |

> **&#42;** 在预览阶段，标识 API 需要 Office 365 （Office 的订阅版本）。 你应该使用来自预览体验成员频道的最新每月版本和内部版本。 你可能需要成为 Office 预览体验成员，才能获取此版本。 有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。 请注意，当内部版本进入生产半年频道时，将关闭对该内部版本的预览功能（包括 SSO）的支持。
>
> **&#8224;** 在这些平台上使用 SSO Api 的外接程序将仅在用户的租户管理员向外接程序授予许可时才有效。 用户不能同意，即使是对自己的 Azure AD 配置文件也不允许。

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="identityapi-preview"></a>IdentityAPI 预览

有关此 API 的详细信息，请参阅在[tokenhelper.getaccesstoken 以便](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-)中使用承诺的版本或在[getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)中使用回调的版本。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](/office/dev/add-ins/develop/add-in-manifests)
