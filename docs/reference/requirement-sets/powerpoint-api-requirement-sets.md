---
title: PowerPoint JavaScript API 要求集
description: 了解有关 PowerPoint JavaScript API 要求集的详细信息。
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: high
ms.openlocfilehash: 2381252ef0d0a4e5b757b38534a826c77108a380
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514003"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>PowerPoint JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

下表列出了 PowerPoint 要求集、支持这些要求集的 Office 客户端应用程序，以及这些应用程序的内部版本或发布日期。

|  要求集  |  Windows 版 Office<br>（关联至 Microsoft 365 订阅）  |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版 |
|:-----|-----|:-----|:-----|:-----|:-----|
| [PowerPointApi 1.3](powerpoint-api-1-3-requirement-set.md)  | 版本 2111 (内部版本 14701.20060) 或更高版本| 尚不可以<br>支持 | 16.55 或更高版本 | 2021 年 12 月 |
| [PowerPointApi 1.2](powerpoint-api-1-2-requirement-set.md)  | 版本 2011（内部版本 13426.20184）或更高版本| 尚不可以<br>支持 | 16.43 或更高版本 | 2020 年 10 月 |
| [PowerPointApi 1.1](powerpoint-api-1-1-requirement-set.md) | 版本 1810（内部版本 11001.20074）或更高版本 | 2.17 或更高版本 | 16.19 或更高版本 | 2018 年 10 月 |

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

有关 Office 版本和内部版本号的详细信息，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a>PowerPoint JavaScript API 1.1

PowerPoint JavaScript API 1.1 包含[用于创建新演示文稿的单一 API](/javascript/api/powerpoint#PowerPoint_createPresentation_base64File_)。 有关 API 的详细信息，请参阅[创建演示文稿](../../powerpoint/powerpoint-add-ins.md#create-a-presentation)。

## <a name="powerpoint-javascript-api-12"></a>PowerPoint JavaScript API 1.2

PowerPoint JavaScript API 1.2 增加了对将其他 PowerPoint 演示文稿中的幻灯片插入当前演示文稿以及删除幻灯片的支持. 有关 API 的详细信息，请参阅[在 PowerPoint 演示文稿中插入和删除幻灯片](../../powerpoint/insert-slides-into-presentation.md)。

## <a name="powerpoint-javascript-api-13"></a>PowerPoint JavaScript API 1.3

PowerPoint JavaScript API 1.3 增加了对添加和删除幻灯片的额外支持。 它还允许外接程序应用自定义元数据标记。 有关 API 的详细信息，请参阅[在 PowerPoint 中添加和删除幻灯片](../../powerpoint/add-slides.md)和[在 PowerPoint 中为演示文稿、幻灯片和形状使用自定义标记](../../powerpoint/tagging-presentations-slides-shapes.md)。

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a>如何在运行时和清单中使用 PowerPoint 要求集

> [!NOTE]
> 本节假定你熟悉 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)和[指定 Office 应用程序和 API 要求集](../../develop/specify-office-hosts-and-api-requirements.md)处的要求集概述。

要求集以 API 成员组命名。Office 加载项可以执行运行时检查，或使用清单中指定的要求集确定某个 Office 应用程序是否支持该加载项所需的 API。

### <a name="checking-for-requirement-set-support-at-runtime"></a>在运行时检查要求集支持

以下代码示例显示如何确定运行加载项的 Office 应用程序是否支持指定的 API 要求集。

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>在清单中定义要求集支持

可以使用加载项清单中的 [“要求元素”](../manifest/requirements.md) 来指定的加载项需要激活的最小需求集和/或 API 方法。如果 Office 应用程序或平台不支持清单的 `Requirements` 元素中指定的要求集或 API 方法，则加载项将不会在该应用程序或平台中运行，也不会出现在 **“我的加载项”** 中显示的加载项列表中。如果加载项需要特定要求集以实现完整功能，但是即使在不支持该要求集的平台上也可以为用户提供值，则建议在运行时按照上述方式检查要求支持，而不是在清单中定义要求集支持。

以下代码示例显示加载项清单中的 `Requirements` 元素，该元素指定应在支持 PowerPointApi 要求集版本 1.1 或更高版本的所有 Office 客户端应用程序中加载该加载项。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

大多数 PowerPoint 加载项功能都来自通用 API 集。 若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [PowerPoint JavaScript API 参考文档](/javascript/api/powerpoint)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
