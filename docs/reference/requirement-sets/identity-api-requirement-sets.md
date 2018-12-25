---
title: Identity API 要求集
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 43a220cfada5883f292edd13cc753dc6c70e3504
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433920"
---
# <a name="identity-api-requirement-sets"></a>Identity API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Office 外接程序在多个 Office 版本中运行。 下表列出了 Identity API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。

|  要求集  | Office 2013 for Windows | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com & Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | 不适用 | 预览 **&#42;** | 即将推出 | 预览 **&#42;**| 预览 | 预览| 即将推出 | 即将推出 |

> **&#42;** 在预览阶段，仅为预览体验计划中使用“快速”选项的用户提供针对 Windows 2016 和 Mac 的 Identity API 支持。 若要加入预览体验计划，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。 若要切换到快速通道，请参阅[预览体验成员快速选项](https://answers.microsoft.com/zh-CN/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961)。

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

有关公用 API 要求集的信息，请参阅 [Office 公用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="identityapi-11"></a>IdentityAPI 1.1 

单一登录 IdentityAPI 1.1 是该 API 的第一版。 有关此 API 的详细信息，请参阅[在外接程序中启用 SSO](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) 的 [SSO API 参考](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)部分。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
