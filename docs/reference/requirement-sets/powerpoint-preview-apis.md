---
title: PowerPoint JavaScript 预览 Api
description: 有关即将推出的 PowerPoint JavaScript Api 的详细信息。
ms.date: 11/09/2020
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: b53b6638b16b2028342003b9a77aa59e7406d5f3
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996520"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint JavaScript 预览 Api

新的 PowerPoint JavaScript Api 是在 "预览" 中首次引入的，并在进行了充分的测试并获得用户反馈之后成为特定的编号要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 插入和删除幻灯片 | 允许将现有幻灯片从另一个演示文稿插入当前演示文稿，以及删除 sildes 的功能。 | [幻灯片. 删除](/javascript/api/powerpoint/powerpoint.slide#delete--)、 [insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|

## <a name="api-list"></a>API 列表

下表列出了当前预览中的 PowerPoint JavaScript Api。 有关所有 PowerPoint JavaScript Api 的完整列表 (包括预览 Api 和以前发布的 Api) ，请参阅 [所有 Powerpoint Javascript api](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[编排](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|指定在幻灯片插入过程中要使用的格式。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|指定将插入到当前演示文稿中的源演示文稿中的幻灯片。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|指定将在演示文稿中插入新幻灯片的位置。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64File： string，options？： InsertSlideOptions) ](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|将演示文稿中指定的幻灯片插入到当前演示文稿中。|
||[页面](/javascript/api/powerpoint/powerpoint.presentation#slides)|返回演示文稿中的幻灯片的已排序集合。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|删除演示文稿中的幻灯片。|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|获取幻灯片的唯一 ID。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|获取集合中的幻灯片数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|使用其唯一 ID 获取幻灯片。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|使用集合中的从零开始的索引获取幻灯片。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|使用其唯一 ID 获取幻灯片。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [PowerPoint JavaScript API 参考文档](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API 要求集](powerpoint-api-requirement-sets.md)
