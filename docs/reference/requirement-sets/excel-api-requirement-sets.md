---
title: Excel JavaScript API 要求集
description: 针对 Excel 内部版本的 Office 加载项要求集信息。
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e77bcb25437e082ce0fbf1b8a695db20ae9f14f1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094384"
---
# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="requirement-set-availability"></a>要求集可用性

Excel 加载项可在多个 Office 版本中运行，包括 Windows 版 Office 2016 或更高版本、Office 网页版、Mac 版 Office 和 iPad 版 Office。 下表列出了 Excel 要求集、支持各要求集的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。

> [!NOTE]
> 若要在任何带编号的要求集或 `ExcelApiOnline` 中使用 API，应引用 CDN 上的“生产”**** 库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js。
>
> 有关使用预览 API 的信息，请参阅 [Excel JavaScript 预览 API](excel-preview-apis.md) 一文。

|  要求集  |  Windows 版 Office<br>（关联至 Microsoft 365 订阅）  |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版 |
|:-----|-----|:-----|:-----|:-----|:-----|
| [预览](excel-preview-apis.md)  | 请使用最新的 Office 版本来试用预览 API（你可能需要加入 [Office 预览体验成员计划](https://insider.office.com)） |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | 不适用 | 不适用 | 不适用 | 最新（请参阅[要求集页面](./excel-api-online-requirement-set.md)） |
| [ExcelApi 1.11](excel-api-1-11-requirement-set.md) | 版本 2002（内部版本 12527.20470）或更高版本 | 16.35 或更高版本 | 16.33 或更高版本 | 2020 年 5 月 |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | 版本 1907（内部版本 11929.20306）或更高版本 | 16.0 或更高版本 | 16.30 或更高版本 | 2019 年 10 月 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | 版本 1903 (内部版本 11425.20204) 或更高版本 | 16.0 或更高版本 | 16.24 或更高版本 | 2019 年 5 月 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | 版本 1808（内部版本 10730.20102）或更高版本 | 16.0 或更高版本 | 16.17 或更高版本 | 2018 年 9 月 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | 版本 1801（内部版本 9001.2171）或更高版本   | 16.0 或更高版本  | 16.9 或更高版本  | 2018 年 4 月 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | 版本 1704（生成号 8201.2001）或更高版本   | 15.0 或更高版本  | 15.36 或更高版本 | 2017 年 4 月 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | 版本 1703（内部版本 8067.2070）或更高版本   | 15.0 或更高版本  | 15.36 或更高版本 | 2017 年 3 月 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | 版本 1701（内部版本 7870.2024）或更高版本   | 15.0 或更高版本  | 15.36 或更高版本 | 2017 年 1 月 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | 版本 1608（内部版本 7369.2055）或更高版本   | 15.0 或更高版本 | 15.27 或更高版本 | 2016 年 9 月 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | 版本 1601（内部版本 6741.2088）或更高版本   | 15.0 或更高版本 | 15.22 或更高版本 | 2016 年 1 月 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | 版本 1509（内部版本 4266.1001）或更高版本   | 15.0 或更高版本 | 15.20 或更高版本 | 2016 年 1 月 |

> [!NOTE]
> 永久版本的 Office 支持要求设置如下：
>
> - Office 2019 支持 ExcelApi 1.8 及更低版本。
> - Office 2016 仅支持 ExcelApi 1.1 要求集。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

有关 Office 版本和内部版本号的详细信息，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 主机和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)
