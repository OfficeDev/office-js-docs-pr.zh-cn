---
title: Excel JavaScript API 要求集
description: 针对 Excel 内部版本的 Office 加载项要求集信息。
ms.date: 01/14/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 466df0fcb48c49d524850e0e92803e0dc10cc3cc
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747112"
---
# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="requirement-set-availability"></a>要求集可用性

Excel 加载项跨多个版本 Office 运行，包括 Windows 版 Office 2016 或更高版本、Office 网页版、Mac 版和 iPad 版。下表列出了 Excel 要求集、支持每个要求集的 Office 客户端应用程序，以及这些应用程序的内部版本或生成号。

> [!NOTE]
> 若要在任何带编号的要求集或 `ExcelApiOnline` 中使用 API，应引用 [Office.js 内容分发网络（CDN）](https://appsforoffice.microsoft.com/lib/1/hosted/office.js)上的 **生产** 库。
>
> 有关使用预览 API 的信息，请参阅 [Excel JavaScript 预览 API](excel-preview-apis.md) 一文。

|  要求集  |  Windows 版 Office<br>（关联至 Microsoft 365 订阅）  |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版 |
|:-----|-----|:-----|:-----|:-----|:-----|
| [预览](excel-preview-apis.md)  | 请使用最新的 Office 版本来试用预览 API（你可能需要加入 [Office 预览体验成员计划](https://insider.office.com)）。 |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | 不适用 | 不适用 | 不适用 | 最新（请参阅[要求集页面](excel-api-online-requirement-set.md)） |
| [ExcelApi 1.14](excel-api-1-14-requirement-set.md) | 版本 2108（内部版本 14326.20508）或更高版本 | 16.53 或更高版本 | 16.52 或更高版本 | 2021 年 10 月 |
| [ExcelApi 1.13](excel-api-1-13-requirement-set.md) | 版本 2102（内部版本 13801.20738）或更高版本 | 16.50 或更高版本 | 16.50 或更高版本 | 2021 年 6 月 |
| [ExcelApi 1.12](excel-api-1-12-requirement-set.md) | 版本 2008（内部版本 13127.20408）或更高版本 | 16.40 或更高版本 | 16.40 或更高版本 | 2020 年 9 月 |
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
> 非订阅版本的 Office 支持如下所示的要求集：
>
> - Office 2021 支持 ExcelApi 1.14 及更低版本。
> - Office 2019 支持 ExcelApi 1.8 及更低版本。
> - Office 2016 仅支持 ExcelApi 1.1 要求集。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

有关 Office 版本和内部版本号的详细信息，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="how-to-use-excel-requirement-sets-at-runtime-and-in-the-manifest"></a>如何在运行时和清单中使用 Excel 要求集

> [!NOTE]
> 本节假定你熟悉 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)和[指定 Office 应用程序和 API 要求集](../../develop/specify-office-hosts-and-api-requirements.md)处的要求集概述。

要求集以 API 成员组命名。Office 加载项可以执行运行时检查，或使用清单中指定的要求集确定某个 Office 应用程序是否支持该加载项所需的 API。

### <a name="checking-for-requirement-set-support-at-runtime"></a>在运行时检查要求集支持

以下代码示例显示如何确定运行加载项的 Office 应用程序是否支持指定的 API 要求集。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>在清单中定义要求集支持

可以使用加载项清单中的“[要求元素](../manifest/requirements.md)”来指定的加载项需要激活的最小需求集和/或 API 方法。如果 Office 应用程序或平台不支持清单的 `Requirements` 元素中指定的要求集或 API 方法，则加载项将不会在该应用程序或平台中运行，也不会出现在“**我的加载项**”中显示的加载项列表中。如果加载项需要特定要求集以实现完整功能，但是即使在不支持该要求集的平台上也可以为用户提供值，则建议在运行时按照上述方式检查要求支持，而不是在清单中定义要求集支持。

以下代码示例显示加载项清单中的 `Requirements` 元素，该元素指定应在支持 ExcelApi 要求集版本 1.3 或更高版本的所有 Office 客户端应用程序中加载该加载项。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
