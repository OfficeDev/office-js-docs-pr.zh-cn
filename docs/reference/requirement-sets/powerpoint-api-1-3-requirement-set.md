---
title: PowerPoint JavaScript API 要求集 1.3
description: 有关 PowerPointApi 1.3 要求集的详细信息。
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# <a name="whats-new-in-powerpoint-javascript-api-13"></a>JavaScript API 1.3 PowerPoint新增功能

PowerPointApi 1.3 增加了对幻灯片管理和自定义标记的额外支持。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [幻灯片管理](../../powerpoint/add-slides.md) | 添加对添加幻灯片以及管理幻灯片版式和幻灯片母版的支持。 | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| [Tags](../../powerpoint/tagging-presentations-slides-shapes.md) | 允许加载项以键值对的形式附加自定义元数据。 | [Tag](/javascript/api/powerpoint/powerpoint.tag) |

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API PowerPoint集 1.3。 有关所有 JavaScript POWERPOINT的完整列表 (包括预览 API 和以前发布的 API) ，请参阅所有 PowerPoint [JavaScript API](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-layoutid-member)|指定要用于新幻灯片的幻灯片版式 ID。|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-slidemasterid-member)|指定要用于新幻灯片的幻灯片母版的 ID。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slidemasters-member)|返回演示文稿 `SlideMaster` 中的对象的集合。|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-tags-member)|返回附加到演示文稿的标记的集合。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-delete-member(1))|从形状集合中删除形状。|
||[id](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-id-member)|获取形状的唯一 ID。|
||[tags](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-tags-member)|返回形状中的标记集合。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getcount-member(1))|获取集合中的形状数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitem-member(1))|使用形状的唯一 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemat-member(1))|使用形状在集合中从零开始编制的索引获取形状。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemornullobject-member(1))|使用形状的唯一 ID 获取形状。|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-items-member)|获取此集合中已加载的子项。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-layout-member)|获取幻灯片的版式。|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-shapes-member)|返回幻灯片中形状的集合。|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-slidemaster-member)|`SlideMaster`获取表示幻灯片的默认内容的对象。|
||[tags](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-tags-member)|返回幻灯片中的标记集合。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[添加 (选项？：PowerPoint。AddSlideOptions) ](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1))|在集合的末尾添加新幻灯片。|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-id-member)|获取幻灯片版式的唯一 ID。|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-name-member)|获取幻灯片版式的名称。|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-shapes-member)|返回幻灯片版式中形状的集合。|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getcount-member(1))|获取集合中的布局数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitem-member(1))|使用唯一 ID 获取布局。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemat-member(1))|获取一个布局，该布局使用集合中从零开始编制的索引。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemornullobject-member(1))|使用唯一 ID 获取布局。|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-items-member)|获取此集合中已加载的子项。|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-id-member)|获取幻灯片母版的唯一 ID。|
||[布局](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-layouts-member)|获取幻灯片母版提供的幻灯片版式的集合。|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-name-member)|获取幻灯片母版的唯一名称。|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-shapes-member)|返回幻灯片母版中的形状的集合。|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getcount-member(1))|获取集合中幻灯片母版的数量。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitem-member(1))|使用幻灯片母版的唯一 ID 获取幻灯片母版。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemat-member(1))|使用集合中从零开始编制的索引获取幻灯片母版。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemornullobject-member(1))|使用幻灯片母版的唯一 ID 获取幻灯片母版。|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-items-member)|获取此集合中已加载的子项。|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-key-member)|获取标记的唯一 ID。|
||[value](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-value-member)|获取标记的值。|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add (key： string， value： string) ](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1))|在集合的末尾添加新标记。|
||[delete (key： string) ](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-delete-member(1))|删除此集合中给定 `key` 标记。|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getcount-member(1))|获取集合中的标记数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitem-member(1))|使用其唯一 ID 获取标记。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemat-member(1))|获取一个标记，该标记使用集合中从零开始编制的索引。|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemornullobject-member(1))|使用其唯一 ID 获取标记。|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-items-member)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [PowerPoint JavaScript API 参考文档](/javascript/api/powerpoint?view=powerpoint-js-1.3&preserve-view=true)
- [PowerPoint JavaScript API 要求集](powerpoint-api-requirement-sets.md)
