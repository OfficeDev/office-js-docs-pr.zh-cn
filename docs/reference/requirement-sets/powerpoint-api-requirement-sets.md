---
title: PowerPoint JavaScript API 要求集
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 4f64654a4130cc0d4bf96d9c59e364e77c808748
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/31/2019
ms.locfileid: "35941143"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a>PowerPoint JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

下表列出了 PowerPoint 要求集、支持这些要求集的 Office 主机应用程序, 以及内部版本或可用性日期。

|  要求集  |  Windows 版 Office<br>(连接到 Office 365 订阅)  |  IPad 上的 Office<br>(连接到 Office 365 订阅)  |  Mac 上的 Office<br>(连接到 Office 365 订阅)  | 网上的 Office |
|:-----|-----|:-----|:-----|:-----|:-----|
| PowerPointApi 1。1 | 版本 1810 (内部版本 11001.20074) 或更高版本 | 2.17 或更高版本 | 16.19 或更高版本 | 2018 年 10 月 |

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

有关 Office 版本和内部版本号的详细信息, 请参阅:

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="powerpoint-javascript-api-11"></a>PowerPoint JavaScript API 1。1

PowerPoint JavaScript API 1.1 包含一个用于创建新演示文稿的 API。 有关 API 的详细信息, 请参阅[JAVASCRIPT API For PowerPoint](../../powerpoint/powerpoint-add-ins.md)。

## <a name="runtime-requirement-support-check"></a>运行时要求支持检查

在运行时, 外接程序可以通过执行以下操作来检查特定主机是否支持 API 要求集。

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a>基于清单的要求支持检查

使用外`Requirements`接程序清单中的元素指定你的外接程序必须使用的关键要求集或 API 成员。 如果 Office 主机或平台不支持`Requirements`元素中指定的要求集或 API 成员, 则外接程序将不会在该主机或平台中运行, 并且不会显示在我的外接程序中。

下面的代码示例展示了加载所有 支持第 1.1 版 OneNoteApi 要求集的 Office 主机应用程序的外接程序。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

大多数 PowerPoint 加载项功能来自通用 API 集。 若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [PowerPoint JavaScript API 参考文档](/javascript/api/powerpoint)
- [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](/office/dev/add-ins/develop/add-in-manifests)
