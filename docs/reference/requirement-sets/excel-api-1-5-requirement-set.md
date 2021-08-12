---
title: ExcelJavaScript API 要求集 1.5
description: 有关 ExcelApi 1.5 要求集的详细信息。
ms.date: 03/19/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1ce97ae8acbcd943027ea799e0c8183790ce3c733c73cc8ed0f39303957e5093
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084907"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Excel JavaScript API 1.5 的最近更新

ExcelApi 1.5 添加自定义 XML 部件。 可通过 workbook 对象中的 [自定义 XML 部件](/javascript/api/excel/excel.workbook#customxmlparts) 集合访问这些部件。

## <a name="custom-xml-part"></a>自定义 XML 部件

* 使用其 ID 获取自定义 XML 部件。
* 获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。
* 获取与部件关联的 XML 字符串。
* 提供部分的 ID 和命名空间。
* 向工作簿添加新的自定义 XML 部件。
* 设置整个 XML 部件。
* 删除自定义 XML 部件。
* 删除其给定名称来自由 xpath 标识的元素的属性。
* 按 xpath 查询 XML 内容。
* 插入、更新和删除属性。

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.5 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.5 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.5](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete__)|删除自定义 XML 部件。|
||[getXml () ](/javascript/api/excel/excel.customxmlpart#getXml__)|获取自定义 XML 部件的完整 XML 内容。|
||[id](/javascript/api/excel/excel.customxmlpart#id)|自定义 XML 部分的 ID。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceUri)|自定义 XML 部分的命名空间 URI。|
||[setXml (xml： string) ](/javascript/api/excel/excel.customxmlpart#setXml_xml_)|设置自定义 XML 部件的完整 XML 内容。|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[添加 (xml：string) ](/javascript/api/excel/excel.customxmlpartcollection#add_xml_)|向工作簿添加新的自定义 XML 部件。|
||[getByNamespace (namespaceUri：string) ](/javascript/api/excel/excel.customxmlpartcollection#getByNamespace_namespaceUri_)|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getCount__)|获取集合中自定义 XML 部件的数量。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getItem_id_)|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getItemOrNullObject_id_)|获取基于其 ID 的自定义 XML 部件。|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|获取此集合中已加载的子项。|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getCount__)|获取此集合中 CustomXML 部件的数量。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItem_id_)|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItemOrNullObject_id_)|获取基于其 ID 的自定义 XML 部件。|
||[getOnlyItem () ](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItem__)|如果集合仅包含一个项，则此方法返回该项。|
||[getOnlyItemOrNullObject () ](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|如果集合仅包含一个项，则此方法返回该项。|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|获取此集合中已加载的子项。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|数据透视表的 ID。|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)||
|[运行时](/javascript/api/excel/excel.runtime)|||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customXmlParts)|表示此工作簿包含的自定义 XML 部件的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getNext_visibleOnly_)|获取此工作表后跟的工作表。|
||[getNextOrNullObject (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getNextOrNullObject_visibleOnly_)|获取此工作表后跟的工作表。|
||[getPrevious (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getPrevious_visibleOnly_)|获取此工作表之前的工作表。|
||[getPreviousOrNullObject (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getPreviousOrNullObject_visibleOnly_)|获取此工作表之前的工作表。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheetcollection#getFirst_visibleOnly_)|获取集合中的第一个工作表。|
||[getLast (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheetcollection#getLast_visibleOnly_)|获取集合中的最后一个工作表。|

## <a name="see-also"></a>另请参阅

* [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
