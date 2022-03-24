---
title: Excel JavaScript API 要求集 1.4
description: 有关 ExcelApi 1.4 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: bcdbd044c5de562b7c2cc2bc9971af31179f8a9b
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746545"
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

API 包括 `getItem()` 通过键获取设置项 `add()` ，以及将指定的 key：value 设置对添加到工作簿。

## <a name="others"></a>其他

* 设置表列名称。
* 将表格列添加到表格的末尾。
* 一次向表中添加多行。
* `range.getColumnsAfter(count: number)` 和 `range.getColumnsBefore(count: number)` 分别用于获取当前 Range 对象的右/左侧的一定数量的列。
* [OrNullObject\* 方法和属性](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties)：此功能允许使用键获取对象。 如果对象不存在，则返回对象的 属性 `isNullObject` 为 true。 这允许开发人员检查对象是否存在，而无需通过异常处理来处理它。 方法 `*OrNullObject` 可用于大多数集合对象。

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.4 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.4 或更早版本支持的所有 API 的 API 参考文档，请参阅[要求集 1.4](/javascript/api/excel?view=excel-js-1.4&preserve-view=true) 或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getcount-member(1))|获取集合中的绑定数量。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemornullobject-member(1))|按 ID 获取绑定对象。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getcount-member(1))|返回工作表中的图表数。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemornullobject-member(1))|使用图表名称获取图表。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getcount-member(1))|返回系列中的图表点数。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getcount-member(1))|返回集合中的系列数量。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-comment-member)|指定与此名称关联的注释。|
||[delete()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-delete-member(1))|删除给定的名称。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrangeornullobject-member(1))|返回与名称相关的 range 对象。|
||[scope](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-scope-member)|指定名称的范围是工作簿还是特定工作表。|
||[worksheet](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheet-member)|返回已命名项限定到的工作表。|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheetornullobject-member)|返回已命名项的作用域的工作表。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add (name： string， reference： Range \| string， comment？： string) ](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-add-member(1))|将新名称添加到给定范围的集合。|
||[addFormulaLocal (name： string， formula： string， comment？： string) ](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-addformulalocal-member(1))|使用用户的公式区域设置，将新名称添加到给定范围的集合。|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getcount-member(1))|获取集合中已命名项的数量。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitemornullobject-member(1))|使用对象 `NamedItem` 的名称获取对象。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getcount-member(1))|获取集合中的数据透视表的数量。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitemornullobject-member(1))|按名称获取 PivotTable 对象。|
|[范围](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#excel-excel-range-getintersectionornullobject-member(1))|获取表示指定区域的矩形交集的 range 对象。|
||[getUsedRangeOrNullObject (值Only？： boolean) ](/javascript/api/excel/excel.range#excel-excel-range-getusedrangeornullobject-member(1))|返回指定 range 对象的所用区域。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getcount-member(1))|获取集合中 `RangeView` 对象的数量。|
|[设置](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#excel-excel-setting-delete-member(1))|删除 Setting 对象。|
||[key](/javascript/api/excel/excel.setting#excel-excel-setting-key-member)|表示设置的 ID 的键。|
||[value](/javascript/api/excel/excel.setting#excel-excel-setting-value-member)|表示为此设置存储的值。|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add (key： string， value： string \| number \| boolean \| Date \| Array \| any) ](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-add-member(1))|设置指定的 Setting 对象，或将其添加到工作簿中。|
||[getCount()](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getcount-member(1))|获取集合中的设置数。|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitem-member(1))|通过 键获取设置条目。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitemornullobject-member(1))|通过 键获取设置条目。|
||[items](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-items-member)|获取此集合中已加载的子项。|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member)|更改文档中的设置时发生。|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[设置](/javascript/api/excel/excel.settingschangedeventargs#excel-excel-settingschangedeventargs-settings-member)|获取表示 `Setting` 引发设置更改事件的绑定的对象|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getcount-member(1))|获取集合中的表数量。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemornullobject-member(1))|按名称或 ID 获取表。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getcount-member(1))|获取表中的列数。|
||[getItemOrNullObject (键：数字 \| 字符串) ](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemornullobject-member(1))|按名称或 ID 获取 column 对象。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getcount-member(1))|获取表格中的行数。|
|[Workbook](/javascript/api/excel/excel.workbook)|[设置](/javascript/api/excel/excel.workbook#excel-excel-workbook-settings-member)|表示与工作簿关联的设置的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (值Only？： boolean) ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getusedrangeornullobject-member(1))|使用的区域是包含分配了值或格式化的任何单元格的最小区域。|
||[names](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-names-member)|一组范围限定到当前工作表的名称。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getcount-member(1))|获取集合中的工作表数量。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitemornullobject-member(1))|使用其名称或 ID 获取 worksheet 对象。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
