---
title: Excel JavaScript API 要求集 1.14
description: 有关 ExcelApi 1.14 要求集的详细信息。
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-114"></a>JavaScript API 1.14 Excel的新增功能

ExcelApi 1.14 添加了对象来控制图表的表功能、用于查找公式的所有引用单元格的方法以及用于跟踪工作表保护状态更改的工作表保护事件。 它还为 对象添加了多种 [`getItemOrNullObject`](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties) 方法，如 `CommentCollection`、 `ShapeCollection`和 `StyleCollection` ，以改进错误处理。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [图表数据表](../../excel/excel-add-ins-charts.md#add-and-format-a-chart-data-table) | 控制图表上数据表的外观、格式和可见性。 | [Chart](/javascript/api/excel/excel.chart)、 [ChartDataTable](/javascript/api/excel/excel.chartdatatable)、 [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [公式引用单元格](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-precedents-of-a-formula) | 返回公式的所有引用单元格。 | [区域](/javascript/api/excel/excel.range) |
| 查询 | 检索 Power Query 属性，如名称、刷新日期和查询计数。 | [Query](/javascript/api/excel/excel.query)、 [QueryCollection](/javascript/api/excel/excel.querycollection)|
| [工作表保护事件](../../excel/excel-add-ins-worksheets.md#detect-changes-to-the-worksheet-protection-state) | 跟踪工作表的保护状态更改以及这些更改的来源。 | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)、 [Worksheet](/javascript/api/excel/excel.worksheet)、 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.14 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.14 或更早版本支持的所有 API 的 API 参考文档，请参阅[要求集 1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true) 或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria (columnIndex： number) ](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcolumncriteria-member(1))|清除自动筛选的列筛选条件。|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-deleteshiftdirection-member)|代表在 (单元格时剩余单元格) 移动的方向，例如向上或向左移动。|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-insertshiftdirection-member)|代表插入 (单元格时现有单元格) 向右或向下移动的方向。|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable () ](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatable-member(1))|获取图表上的数据表。|
||[getDataTableOrNullObject () ](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatableornullobject-member(1))|获取图表上的数据表。|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-format-member)|表示图表数据表的格式，包括填充、字体和边框格式。|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showhorizontalborder-member)|指定是否显示数据表的水平边框。|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showlegendkey-member)|指定是否显示数据表的图例项键。|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showoutlineborder-member)|指定是否显示数据表的外边框。|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showverticalborder-member)|指定是否显示数据表的垂直边框。|
||[visible](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-visible-member)|指定是否显示图表的数据表。|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-border-member)|表示图表数据表的边框格式，其中包括颜色、线条样式和粗细。|
||[fill](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-fill-member)|表示对象的填充格式，包括背景格式信息。|
||[font](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-font-member)|代表当前 (字体名称、字号和颜色) 字体属性。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject (commentId： string) ](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemornullobject-member(1))|根据其 ID 从集合中获取批注。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject (commentReplyId： string) ](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemornullobject-member(1))|返回由其 ID 标识的批注回复。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemornullobject-member(1))|返回由 ID 标识的条件格式。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemornullobject-member(1))|使用形状的名称或 ID 获取形状。|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#excel-excel-query-error-member)|从上次刷新查询时获取查询错误消息。|
||[loadedTo](/javascript/api/excel/excel.query#excel-excel-query-loadedto-member)|获取加载到 对象类型。|
||[loadedToDataModel](/javascript/api/excel/excel.query#excel-excel-query-loadedtodatamodel-member)|指定是否将查询加载到数据模型。|
||[name](/javascript/api/excel/excel.query#excel-excel-query-name-member)|获取查询的名称。|
||[refreshDate](/javascript/api/excel/excel.query#excel-excel-query-refreshdate-member)|获取上次刷新查询的日期和时间。|
||[rowsLoadedCount](/javascript/api/excel/excel.query#excel-excel-query-rowsloadedcount-member)|获取上次刷新查询时加载的行数。|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getcount-member(1))|获取工作簿中的查询数。|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getitem-member(1))|根据名称从集合获取查询。|
||[items](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-items-member)|获取此集合中已加载的子项。|
|[区域](/javascript/api/excel/excel.range)|[getPrecedents () ](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1))|返回一 `WorkbookRangeAreas` 个对象，该对象表示包含同一工作表或多个工作表中单元格的所有引用单元格的范围。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemornullobject-member(1))|使用形状的名称或 ID 获取形状。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemornullobject-member(1))|按名称获取样式。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitemornullobject-member(1))|按名称或 ID 获取表。|
|[Workbook](/javascript/api/excel/excel.workbook)|[查询](/javascript/api/excel/excel.workbook#excel-excel-workbook-queries-member)|返回属于工作簿的 Power Query 查询的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member)|工作表保护状态更改时发生。|
||[tabId](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabid-member)|返回一个值，该值代表此工作表，该工作表可通过 Open Office XML 读取。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changedirectionstate-member)|表示工作表中单元格在删除或插入时移动的方向的变化。|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-triggersource-member)|表示事件的触发源。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member)|工作表保护状态更改时发生。|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-isprotected-member)|获取工作表的当前保护状态。|
||[源](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-source-member)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-worksheetid-member)|获取其中保护状态发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
