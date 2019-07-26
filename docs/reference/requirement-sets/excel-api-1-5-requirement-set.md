---
title: Excel JavaScript API 要求集1。5
description: 有关 ExcelApi 1.5 要求集的详细信息
ms.date: 07/15/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b8f767a83b7e373b422b6fc0d9ac65de90c04f5
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771965"
---
#  <a name="whats-new-in-excel-javascript-api-15"></a>Excel JavaScript API 1.5 的最近更新

ExcelApi 1.5 添加自定义 XML 部件。 可通过工作簿对象中的[自定义 XML 部件集合](/javascript/api/excel/excel.workbook#customxmlparts)访问这些项。

## <a name="custom-xml-part"></a>自定义 XML 部件

* 使用其 ID 获取自定义 XML 部件。
* 获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。
* 获取与部件相关联的 XML 字符串。
* 提供部件的 ID 和命名空间。
* 将新的自定义 XML 部件添加到工作簿中。
* 设置完整的 XML 部件。
* 删除自定义 XML 部件。
* 删除其给定名称来自由 xpath 标识的元素的属性。
* 按 xpath 查询 XML 内容。
* 插入、更新和删除属性。

## <a name="api-list"></a>API 列表

| Class | 域 | 说明 |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|删除自定义 XML 部件。|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getxml--)|获取自定义 XML 部件的完整 XML 内容。|
||[id](/javascript/api/excel/excel.customxmlpart#id)|自定义 XML 部件的 ID。 只读。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|自定义 XML 部件的命名空间 URI。 只读。|
||[setXml (xml: string)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|设置自定义 XML 部件的完整 XML 内容。|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add (xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|向工作簿添加新的自定义 XML 部件。|
||[getByNamespace (namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|获取此集合中 CustomXml 部件的数量。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|获取基于其 ID 的自定义 XML 部件。|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|获取此集合中已加载的子项。|
|[CustomXmlPartCollectionLoadOptions](/javascript/api/excel/excel.customxmlpartcollectionloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#id)|对于集合中的每一项: 自定义 XML 部件的 ID。 只读。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartcollectionloadoptions#namespaceuri)|对于集合中的每一项: 自定义 XML 部件的命名空间 URI。 只读。|
|[CustomXmlPartData](/javascript/api/excel/excel.customxmlpartdata)|[id](/javascript/api/excel/excel.customxmlpartdata#id)|自定义 XML 部件的 ID。 只读。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartdata#namespaceuri)|自定义 XML 部件的命名空间 URI。 只读。|
|[CustomXmlPartLoadOptions](/javascript/api/excel/excel.customxmlpartloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartloadoptions#id)|自定义 XML 部件的 ID。 只读。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartloadoptions#namespaceuri)|自定义 XML 部件的命名空间 URI。 只读。|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|获取此集合中 CustomXML 部件的数量。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|获取基于其 ID 的自定义 XML 部件。|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|如果集合仅包含一个项，则此方法返回该项。|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|如果集合仅包含一个项，则此方法返回该项。|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|获取此集合中已加载的子项。|
|[CustomXmlPartScopedCollectionLoadOptions](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#id)|对于集合中的每一项: 自定义 XML 部件的 ID。 只读。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpartscopedcollectionloadoptions#namespaceuri)|对于集合中的每一项: 自定义 XML 部件的命名空间 URI。 只读。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|数据透视表的 ID。 只读。|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[id](/javascript/api/excel/excel.pivottablecollectionloadoptions#id)|对于集合中的每一项: 数据透视表的 Id。 只读。|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[id](/javascript/api/excel/excel.pivottabledata#id)|数据透视表的 ID。 只读。|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[id](/javascript/api/excel/excel.pivottableloadoptions#id)|数据透视表的 ID。 只读。|
|[语言](/javascript/api/excel/excel.runtime)|[set (properties: Excel. Runtime)](/javascript/api/excel/excel.runtime#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RuntimeUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.runtime#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[$all](/javascript/api/excel/excel.runtimeloadoptions#$all)||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|表示此工作簿包含的自定义 XML 部件的集合。 只读。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[customXmlParts](/javascript/api/excel/excel.workbookdata#customxmlparts)|表示此工作簿包含的自定义 XML 部件的集合。 只读。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly？: 布尔值)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|获取此工作表的后面的工作表。 如果此方法后面没有任何工作表, 则此方法将引发错误。|
||[getNextOrNullObject (visibleOnly？: 布尔值)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|获取此工作表的后面的工作表。 如果此方法后面没有任何工作表, 则此方法将返回一个 null 对象。|
||[getPrevious (visibleOnly？: 布尔值)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|获取此项之前的工作表。 如果没有以前的工作表, 此方法将引发错误。|
||[getPreviousOrNullObject (visibleOnly？: 布尔值)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|获取此项之前的工作表。 如果没有以前的工作表, 则此方法将返回一个空的 objet。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly？: 布尔值)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|获取集合中的第一个工作表。|
||[getLast (visibleOnly？: 布尔值)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|获取集合中的最后一个工作表。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)