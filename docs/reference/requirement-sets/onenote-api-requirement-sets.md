---
title: OneNote JavaScript API 要求集
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 00bf9f23c307a6094345b753d7cccf1c10be7c32
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950962"
---
# <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

下表列出了 OneNote 要求集、支持这些要求集的 Office 主机应用程序，以及这些应用程序的内部版本或发布日期。

|  要求集  |  Office 网页版 |
|:-----|:-----|
| [OneNoteApi 1.1](/javascript/api/onenote?view=onenote-js-1.1)  | 2016 年 9 月 |  

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="onenote-javascript-api-11"></a>OneNote JavaScript API 1.1

OneNote JavaScript API 1.1 是该 API 的第一版。 有关此 API 的详细信息，请参阅 [OneNote JavaScript API 编程概述](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)。

## <a name="runtime-requirement-support-check"></a>运行时要求支持检查

在运行时，加载项可以执行下列检查，确定特定主机是否支持 API 要求集。

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>基于清单的要求支持检查

使用加载项清单中的 `Requirements` 元素指定加载项必须使用的关键要求集或 API 成员。 如果 Office 主机或平台不支持 `Requirements` 元素中指定的要求集或 API 成员，则加载项将无法在该主机或平台上运行，并且不会显示在“我的加载项”中。

下面的代码示例展示了加载所有 支持第 1.1 版 OneNoteApi 要求集的 Office 主机应用程序的外接程序。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [OneNote JavaScript API 参考文档](/javascript/api/onenote)
- [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](/office/dev/add-ins/develop/add-in-manifests)
