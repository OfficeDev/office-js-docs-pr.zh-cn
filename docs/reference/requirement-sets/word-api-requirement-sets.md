---
title: Word JavaScript API 要求集
description: 针对 Word 内部版本的 Office 加载项要求集信息。
ms.date: 04/16/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 69addff587b026ab717a438e204b1258b06266b6
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547218"
---
# <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

## <a name="requirement-set-availability"></a>要求集可用性

Word 加载项可在多个 Office 版本中运行，包括 Windows 版 Office 2016 或更高版本、Office 网页版、iPad 版 Office 和 Mac 版 Office。 下表列出了 Word 要求集、支持该要求集的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。

> [!NOTE]
> 若要在任何编号的要求集中使用 API，你应该引用 CDN 上的**生产**库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js。
>
> 有关使用预览 API 的信息，请参阅 [Excel JavaScript 预览 API](word-preview-apis.md) 一文。

|  要求集  |   Windows 版 Office\*<br>（连接到 Office 365 订阅）  |  iPad 版 Office<br>（已连接到 Office 365 订阅）  |  Mac 版 Office<br>（已连接到 Office 365 订阅）  | Office 网页版  |
|:-----|-----|:-----|:-----|:-----|
| [预览](word-preview-apis.md) | 请使用最新的 Office 版本来试用预览 API（你可能需要加入 [Office 预览体验成员计划](https://insider.office.com)） |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | 版本 1612（内部版本 7668.1000）或更高版本| 2017 年 3 月，2.22 或更高版本 | 2017 年 3 月，15.32 或更高版本| 2017 年 3 月 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | 2015 年 12 月更新，版本 1601（内部版本 6568.1000）或更高版本 | 2016 年 1 月，1.18 或更高版本 | 2016 年 1 月，15.19 或更高版本| 2016 年 9 月 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | 版本 1509（内部版本 4266.1001）或更高版本| 2016 年 1 月，1.18 或更高版本 | 2016 年 1 月，15.19 或更高版本| 2016 年 9 月 |

> [!NOTE]
> 永久版本的 Office 支持要求集如下：
>
> - Office 2019 支持 ExcelApi 1.3 及更低版本。
> - Office 2016 仅支持 ExcelApi 1.1 要求集。

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

有关 Office 版本和内部版本号的详细信息，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 主机和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)
