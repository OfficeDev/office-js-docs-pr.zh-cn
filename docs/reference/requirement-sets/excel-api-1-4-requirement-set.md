---
title: ExcelJavaScript API 要求集 1.4
description: 有关 ExcelApi 1.4 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: be71d1e0c063bd3902bf57ba8f2024ae5a78ff1d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937157"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 的最近更新

下面介绍了要求集 1.4 中 Excel JavaScript API 的新增内容。

## <a name="named-item-add-and-new-properties"></a>添加了已命名项和新属性

新属性：

* `comment`
* `scope` - 工作表或工作簿范围的项目。
* `worksheet` - 返回已命名项的作用域所基于的工作表。

新方法：

* `add(name: string, reference: Range or string, comment: string)` - 将新名称添加到给定范围的集合。
* `addFormulaLocal(name: string, formula: string, comment: string)` - 使用公式的用户区域设置将新名称添加到给定范围的集合。

## <a name="settings-api-in-the-excel-namespace"></a>Excel 命名空间中的设置 API

[Setting](/javascript/api/excel/excel.setting) 对象表示文档保留设置的键值对。 `Excel.Setting` 的功能等同于 `Office.Settings`，但使用批处理 API 语法，而不是通用 API 的回调模型。

API 包括 `getItem()` 通过键获取设置项，以及将指定的 `add()` key：value 设置对添加到工作簿。

## <a name="others"></a>其他

* 设置表列名称。
* 将表格列添加到表格的末尾。
* 一次向表中添加多行。
* `range.getColumnsAfter(count: number)` 和 `range.getColumnsBefore(count: number)` 分别用于获取当前 Range 对象的右/左侧的一定数量的列。
* [ \* OrNullObject 方法和属性](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties)：此功能允许使用键获取对象。 如果对象不存在，则返回对象的 `isNullObject` 属性为 true。 这允许开发人员检查对象是否存在，而无需通过异常处理来处理它。 方法 `*OrNullObject` 可用于大多数集合对象。

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求Excel集 1.4 中的 API。 若要查看受 Excel JavaScript API 要求集 1.4 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.4](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getCount__)|获取集合中的绑定数量。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getItemOrNullObject_id_)|按 ID 获取绑定对象。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getCount__)|返回工作表中的图表数。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getItemOrNullObject_name_)|使用图表名称获取图表。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getCount__)|返回系列中的图表点数。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getCount__)|返回集合中的系列数量。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|指定与此名称关联的注释。|
||[delete()](/javascript/api/excel/excel.nameditem#delete__)|删除给定的名称。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getRangeOrNullObject__)|返回与名称相关的 range 对象。|
||[scope](/javascript/api/excel/excel.nameditem#scope)|指定名称的范围是工作簿还是特定工作表。|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|返回已命名项限定到的工作表。|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetOrNullObject)|返回已命名项的作用域的工作表。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add (name： string， reference： Range \| string， comment？： string) ](/javascript/api/excel/excel.nameditemcollection#add_name__reference__comment_)|将新名称添加到给定范围的集合。|
||[addFormulaLocal (name： string， formula： string， comment？： string) ](/javascript/api/excel/excel.nameditemcollection#addFormulaLocal_name__formula__comment_)|使用用户的公式区域设置，将新名称添加到给定范围的集合。|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getCount__)|获取集合中已命名项的数量。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getItemOrNullObject_name_)|使用 `NamedItem` 对象的名称获取对象。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getCount__)|获取集合中的数据透视表的数量。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getItemOrNullObject_name_)|按名称获取 PivotTable 对象。|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#getIntersectionOrNullObject_anotherRange_)|获取表示指定区域的矩形交集的 range 对象。|
||[getUsedRangeOrNullObject (valuesOnly？： boolean) ](/javascript/api/excel/excel.range#getUsedRangeOrNullObject_valuesOnly_)|返回指定 range 对象的所用区域。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getCount__)|获取集合 `RangeView` 中对象的数量。|
|[设置](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete__)|删除 Setting 对象。|
||[key](/javascript/api/excel/excel.setting#key)|表示设置的 ID 的键。|
||[value](/javascript/api/excel/excel.setting#value)|表示为此设置存储的值。|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add (key： string， value： string \| number \| boolean \| Date Array any \| <any> \|) ](/javascript/api/excel/excel.settingcollection#add_key__value_)|设置指定的 Setting 对象，或将其添加到工作簿中。|
||[getCount()](/javascript/api/excel/excel.settingcollection#getCount__)|获取集合中的设置数。|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getItem_key_)|通过 键获取设置条目。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getItemOrNullObject_key_)|通过 键获取设置条目。|
||[items](/javascript/api/excel/excel.settingcollection#items)|获取此集合中已加载的子项。|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onSettingsChanged)|更改文档中的设置时发生。|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[设置](/javascript/api/excel/excel.settingschangedeventargs#settings)|获取 `Setting` 表示引发设置更改事件的绑定的对象|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getCount__)|获取集合中的表数量。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getItemOrNullObject_key_)|按名称或 ID 获取表。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getCount__)|获取表中的列数。|
||[getItemOrNullObject (键：数字 \| 字符串) ](/javascript/api/excel/excel.tablecolumncollection#getItemOrNullObject_key_)|按名称或 ID 获取 column 对象。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getCount__)|获取表格中的行数。|
|[Workbook](/javascript/api/excel/excel.workbook)|[设置](/javascript/api/excel/excel.workbook#settings)|表示与工作簿关联的设置的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getUsedRangeOrNullObject_valuesOnly_)|使用的区域是包含分配了值或格式化的任何单元格的最小区域。|
||[names](/javascript/api/excel/excel.worksheet#names)|一组范围限定到当前工作表的名称。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheetcollection#getCount_visibleOnly_)|获取集合中的工作表数量。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getItemOrNullObject_key_)|使用其名称或 ID 获取 worksheet 对象。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
