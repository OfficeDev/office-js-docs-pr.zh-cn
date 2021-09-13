---
title: PowerPointJavaScript 预览 API
description: 有关即将推出的 JavaScript PowerPoint的详细信息。
ms.date: 01/27/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: d9cb28c56a84829d87ba30e494aa46b927e0bc64
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152658"
---
# <a name="powerpoint-javascript-preview-apis"></a>PowerPointJavaScript 预览 API

JavaScript API PowerPoint在"预览"中首次引入，之后在经过充分测试且获得用户反馈后，它将成为特定编号要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 幻灯片管理 | 添加对添加幻灯片以及管理幻灯片版式和幻灯片母版的支持。 | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| 形状 | 添加对获取对幻灯片中形状的引用的支持。 | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>API 列表

下表列出了当前预览PowerPoint JavaScript API 的列表。 有关所有 JavaScript POWERPOINT API 的完整列表 (包括预览 API 和以前发布的 API) ，请参阅所有 Excel [JavaScript API。](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|指定要用于新幻灯片的幻灯片版式 ID。|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|指定要用于新幻灯片的幻灯片母版的 ID。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|返回演示文稿 `SlideMaster` 中的对象的集合。|
||[标记](/javascript/api/powerpoint/powerpoint.presentation#tags)|返回附加到演示文稿的标记的集合。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|从形状集合中删除形状。|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|获取形状的唯一 ID。|
||[标记](/javascript/api/powerpoint/powerpoint.shape#tags)|返回形状中的标记集合。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|获取集合中的形状数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|使用形状的唯一 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|使用形状在集合中从零开始编制的索引获取形状。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|使用形状的唯一 ID 获取形状。|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|获取此集合中已加载的子项。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|获取幻灯片的版式。|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|返回幻灯片中形状的集合。|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|获取 `SlideMaster` 表示幻灯片的默认内容的对象。|
||[标记](/javascript/api/powerpoint/powerpoint.slide#tags)|返回幻灯片中的标记集合。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[添加 (选项？：PowerPoint。AddSlideOptions) ](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|在集合的末尾添加新幻灯片。|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|获取幻灯片版式的唯一 ID。|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|获取幻灯片版式的名称。|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|获取集合中的布局数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|使用唯一 ID 获取布局。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|获取一个布局，该布局使用集合中从零开始编制的索引。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|使用唯一 ID 获取布局。|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|获取此集合中已加载的子项。|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|获取幻灯片母版的唯一 ID。|
||[布局](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|获取幻灯片母版提供的幻灯片版式的集合。|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|获取幻灯片母版的唯一名称。|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|获取集合中幻灯片母版的数量。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|使用幻灯片母版的唯一 ID 获取幻灯片母版。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|使用集合中从零开始编制的索引获取幻灯片母版。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|使用幻灯片母版的唯一 ID 获取幻灯片母版。|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|获取此集合中已加载的子项。|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|获取标记的唯一 ID。|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|获取标记的值。|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add (key： string， value： string) ](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|在集合的末尾添加新标记。|
||[delete (key： string) ](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|删除此集合中给定 `key` 标记。|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|获取集合中的标记数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|使用其唯一 ID 获取标记。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|获取一个标记，该标记使用集合中从零开始编制的索引。|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|使用其唯一 ID 获取标记。|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [PowerPointJavaScript API 参考文档](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [PowerPoint JavaScript API 要求集](powerpoint-api-requirement-sets.md)