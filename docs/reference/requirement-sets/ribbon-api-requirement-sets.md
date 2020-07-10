---
title: 功能区 API 要求集
description: 指定哪些 Office 平台和生成支持动态功能区 Api。
ms.date: 07/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 6a0e6af3a74b0b0402710fd66bac6c915aa4c18a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094279"
---
# <a name="ribbon-api-requirement-sets"></a>功能区 API 要求集

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

功能区 API 集支持编程控制何时自定义外接程序命令 (即启用和禁用自定义功能区按钮和菜单项) 。

Office 外接程序在多个 Office 版本中运行。 下表列出了功能区 API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本号或版本号。

|  要求集  | Windows 版 Office 2013<br>（一次性购买） | Windows 上的 Office 2016 或更高版本<br>（一次性购买）   | Windows 版 Office\*<br> (连接到 Microsoft 365 订阅)  |  iPad 版 Office<br> (连接到 Microsoft 365 订阅)   |  Mac 版 Office\*<br> (连接到 Microsoft 365 订阅)   | Office 网页版\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1。1  | 不适用 | 不适用 | 版本 2002 (内部版本 12527.20264) 或更高版本 | 16.38 或更高版本 | 不适用 | 2020 年 2 月 | 不适用|

> **&#42;** 在预览阶段，仅在 Excel 中支持功能区 API，并且它需要 Microsoft 365 订阅。 你应该使用来自预览体验成员频道的最新每月版本和内部版本。 你可能需要成为 Office 预览体验成员，才能获取此版本。 有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。 请注意，当内部版本毕业生到生产半年频道时，将为该生成关闭对预览功能（包括功能区 API）的支持。

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [适用于 Microsoft 365 客户端的更新通道版本和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Microsoft 365 客户端应用程序的版本和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="ribbon-api-11"></a>功能区 API 1。1

功能区 API 1.1 是 API 的第一个版本。 有关 API 的详细信息，请参阅 " [Office. 功能区](/javascript/api/office/office.ribbon)参考" 主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](/office/dev/add-ins/develop/add-in-manifests)
