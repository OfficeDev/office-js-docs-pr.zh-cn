---
title: PowerPoint JavaScript API 要求集
description: 了解有关 PowerPoint JavaScript API 要求集的详细信息。
ms.date: 07/10/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: b2b5d4b7b5a0677812f227b6a32683c35bbf1662
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293504"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>PowerPoint JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

下表列出了 PowerPoint 要求集、支持这些要求集的 Office 客户端应用程序，以及这些应用程序的内部版本或发布日期。

|  要求集  |  Windows 版 Office<br>（关联至 Microsoft 365 订阅）  |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版 |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1.1 | 版本 1810（内部版本 11001.20074）或更高版本 | 2.17 或更高版本 | 16.19 或更高版本 | 2018 年 10 月 |

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

有关 Office 版本和内部版本号的详细信息，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a>PowerPoint JavaScript API 1.1

PowerPoint JavaScript API 1.1 包含用于创建新演示文稿的单一 API。 有关该 API 的详细信息，请参阅[适用于 PowerPoint 的 JavaScript API](../../powerpoint/powerpoint-add-ins.md)。

## <a name="runtime-requirement-support-check"></a>运行时要求支持检查

在运行时，加载项可以执行下列检查，确定特定应用程序是否支持 API 要求集。

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>基于清单的要求支持检查

使用加载项清单中的 `Requirements` 元素指定加载项必须使用的关键要求集或 API 成员。 如果 Office 应用程序或平台不支持 `Requirements` 元素中指定的要求集或 API 成员，则加载项将无法在该应用程序或平台上运行，并且不会显示在“我的加载项”中。

下面的代码示例展示了加载所有支持第 1.1 版 OneNoteApi 要求集的 Office 客户端应用程序的加载项。

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
- [指定 Office 应用程序和 API 要求集](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
