---
title: Word JavaScript API 要求集
description: 面向 Word 的 Office 加载项要求集信息。
ms.date: 01/14/2022
ms.prod: word
ms.localizationpriority: high
ms.openlocfilehash: 25a698a82669210a596026807d9e30be38a68762
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746169"
---
# <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API 要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="requirement-set-availability"></a>要求集可用性

Word 加载项可跨多个版本的 Office 运行，包括 Windows 上的 Office 2016 或更高版本，以及 Office 网页版、iPad 和 Mac 版。下表列出了 Word 要求集、支持该要求集的 Office 客户端应用程序，以及这些应用程序的内部版本或版本号。

> [!NOTE]
> 若要在任何编号的要求集内使用 API，应引用 [ Office.js 内容分发网络（CDN）](https://appsforoffice.microsoft.com/lib/1/hosted/office.js)上的 **生产** 库。
>
> 有关使用预览 API 的信息，请参阅 [Excel JavaScript 预览 API](word-preview-apis.md) 一文。

|  要求集  |   Windows 版 Office\*<br>（关联至 Microsoft 365 订阅）  |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |
|:-----|-----|:-----|:-----|:-----|
| [预览](word-preview-apis.md) | 请使用最新的 Office 版本来试用预览 API（你可能需要加入 [Office 预览体验成员计划](https://insider.office.com)） |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | 版本 1612（内部版本 7668.1000）或更高版本| 2017 年 3 月，2.22 或更高版本 | 2017 年 3 月，15.32 或更高版本| 2017 年 3 月 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | 2015 年 12 月更新，版本 1601（内部版本 6568.1000）或更高版本 | 2016 年 1 月，1.18 或更高版本 | 2016 年 1 月，15.19 或更高版本| 2016 年 9 月 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | 版本 1509（内部版本 4266.1001）或更高版本| 2016 年 1 月，1.18 或更高版本 | 2016 年 1 月，15.19 或更高版本| 2016 年 9 月 |

> [!NOTE]
> 非订阅版本的 Office 支持如下所示的要求集：
>
> - Office 2019 和 Office 2021 支持 WordApi 1.3 及更低版本。
> - Office 2016 仅支持 ExcelApi 1.1 要求集。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

有关 Office 版本和内部版本号的详细信息，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
