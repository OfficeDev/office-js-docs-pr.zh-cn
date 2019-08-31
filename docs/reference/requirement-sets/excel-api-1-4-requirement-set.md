---
title: Excel JavaScript API 要求集1。4
description: 有关 ExcelApi 1.4 要求集的详细信息
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b0f74d4de5ec867e21e4bec1cd9ab1983a87bab1
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695831"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 的最近更新

下面介绍了要求集 1.4 中 Excel JavaScript API 的新增内容。

## <a name="named-item-add-and-new-properties"></a>添加了已命名项和新属性

新属性：

* `comment`
* `scope`-工作表或工作簿的限定项。
* `worksheet`-返回命名项目的作用范围为的工作表。

新方法：

* `add(name: string, reference: Range or string, comment: string)`-将新名称添加到给定范围的集合。
* `addFormulaLocal(name: string, formula: string, comment: string)`-使用用户的公式区域设置, 将新名称添加到给定范围的集合。

## <a name="settings-api-in-the-excel-namespace"></a>Excel 命名空间中的设置 API

[Setting](/javascript/api/excel/excel.setting) 对象表示文档保留设置的键值对。 `Excel.Setting` 的功能等同于 `Office.Settings`，但使用批处理 API 语法，而不是通用 API 的回调模型。

Api 包括`getItem()`通过键获取设置条目并`add()`将指定的键: value 设置对添加到工作簿中。

## <a name="others"></a>其他

* 设置表的列名称。
* 将表格列添加到表的末尾。
* 一次向表中添加多个行。
* `range.getColumnsAfter(count: number)` 和 `range.getColumnsBefore(count: number)` 分别用于获取当前 Range 对象的右/左侧的一定数量的列。
* [Get item 或 null 对象函数](../../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods): 此功能允许使用键获取对象。 如果该对象不存在, 则返回的对象的`isNullObject`属性将为 true。 这样, 开发人员就可以检查某个对象是否存在, 而无需通过异常处理处理它。 此`*OrNullObject`方法可用于大多数集合对象。

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集1.4 中的 Api。 若要查看 Excel JavaScript API 要求集1.4 或更早版本支持的所有 Api 的 API 参考文档, 请参阅[要求集1.4 或更早版本中的 Excel api](/javascript/api/excel?view=excel-js-1.4)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|获取集合中的绑定数量。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|按 ID 获取 Binding 对象。 如果没有 Binding 对象，将返回 NULL 对象。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|返回工作表中的图表数。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|使用图表名称获取图表。 如果存在多个名称相同的图表，将返回第一个图表。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|返回系列中的图表点数。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|返回集合中的系列数量。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|表示与此名称相关联的注释。|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|删除给定的名称。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|返回与名称相关联的 Range 对象。 如果已命名项的类型不是 Range，将返回 NULL 对象。|
||[scope](/javascript/api/excel/excel.nameditem#scope)|指明是否将 name 限定到工作簿或特定工作表。 可能的值为: 工作表、工作簿。 只读。|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|返回已命名项限定到的工作表。 如果项目的作用域改为工作簿, 则会引发错误。|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|返回已命名项限定到的工作表。 如果项改为限定到工作簿，将返回 NULL 对象。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add (name: string, reference: Range \| string, comment？: string)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|将新名称添加到给定范围的集合。|
||[addFormulaLocal (name: string, formula: string, comment？: string)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|使用用户的公式区域设置，将新名称添加到给定范围的集合。|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|获取集合中已命名项的数量。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|使用其名称获取 NamedItem 对象。 如果没有 NamedItem 对象，将返回 NULL 对象。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|获取集合中的数据透视表的数量。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|按 PivotTable 对象的名称获取此对象。 如果没有 PivotTable 对象，将返回 NULL 对象。|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: Range \|字符串)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|获取表示指定区域的矩形交集的 range 对象。 如果找不到任何交集，则此方法返回空对象。|
||[getUsedRangeOrNullObject (valuesOnly？: 布尔值)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|返回指定 Range 对象的所用区域。如果区域内没有使用单元格，此函数将返回 NULL 对象。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|获取集合中 RangeView 对象的数量。|
|[设置](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|删除 Setting 对象。|
||[key](/javascript/api/excel/excel.setting#key)|返回表示 setting 对象的 ID 的键。 只读。|
||[value](/javascript/api/excel/excel.setting#value)|表示为此设置存储的值。|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add (key: string, value: string \| number \| boolean \| Date \| Array<any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|设置指定的 Setting 对象，或将其添加到工作簿中。|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|获取集合中的 Setting 对象的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|按键获取 Setting 项。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|按键获取 Setting 项。 如果没有 Setting 项，将返回 NULL 对象。|
||[items](/javascript/api/excel/excel.settingcollection#items)|获取此集合中已加载的子项。|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|当文档中的设置变化时发生。|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#settings)|获取表示引发了 SettingsChanged 事件的 binding 的 setting 对象。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|获取集合中的表数量。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|按名称或 ID 获取表。 如果没有表，将返回 NULL 对象。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|获取表中的列数。|
||[getItemOrNullObject (key: 数字\|字符串)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|按名称或 ID 获取 column 对象。 如果没有 column 对象，将返回 NULL 对象。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|获取表格中的行数。|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#settings)|表示一组与 workbook 相关联的 setting 对象。 只读。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly？: 布尔值)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|使用的区域是包含分配了值或格式的任意单元格的最小区域。如果整个工作表为空，此函数将返回 NULL 对象。|
||[名称](/javascript/api/excel/excel.worksheet#names)|一组范围限定到当前工作表的名称。 只读。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount (visibleOnly？: 布尔值)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|获取集合中的工作表数量。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|按 Worksheet 对象的名称或 ID 获取此对象。 如果没有 Worksheet 对象，将返回 NULL 对象。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.4)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
