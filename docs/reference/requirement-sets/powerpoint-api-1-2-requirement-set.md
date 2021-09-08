---
title: PowerPointJavaScript API 要求集 1.2
description: 有关 PowerPointApi 1.2 要求集的详细信息。
ms.date: 01/27/2021
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: fac472e9b88b78f52fe939f883d88cded8b1702c
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938913"
---
# <a name="whats-new-in-powerpoint-javascript-api-12"></a>JavaScript API 1.2 PowerPoint的新增功能

PowerPointApi 1.2 增加了对将另一个演示文稿中的幻灯片插入当前演示文稿以及删除幻灯片的支持。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [插入和删除幻灯片](../../powerpoint/insert-slides-into-presentation.md) | 允许将现有幻灯片从另一个演示文稿插入当前演示文稿，以及删除幻灯片。 | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--) [、Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API PowerPoint集 1.2。 有关所有 JavaScript POWERPOINT的完整列表 (包括预览 API 和以前发布的 API) ，请参阅所有 PowerPoint [JavaScript API。](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[格式设置](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|指定在幻灯片插入过程中使用的格式。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceSlideIds)|指定将插入到当前演示文稿的源演示文稿中的幻灯片。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetSlideId)|指定演示文稿中新幻灯片的插入位置。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64File： string， options？： PowerPoint。InsertSlideOptions) ](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)|将演示文稿中的指定幻灯片插入到当前演示文稿中。|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|返回演示文稿中幻灯片的有序集合。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete__)|从演示文稿中删除幻灯片。|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|获取幻灯片的唯一 ID。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getCount__)|获取集合中幻灯片的数量。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItem_key_)|使用其唯一 ID 获取幻灯片。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)|使用集合中从零开始索引获取幻灯片。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemOrNullObject_id_)|使用其唯一 ID 获取幻灯片。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [PowerPointJavaScript API 参考文档](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [PowerPoint JavaScript API 要求集](powerpoint-api-requirement-sets.md)
