---
title: Excel JavaScript API 要求集1。6
description: 有关 ExcelApi 1.6 要求集的详细信息
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9e1a3375d19d8c1cb0fbddac50fabf826b96d7cc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771972"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Excel JavaScript API 1.6 的最近更新

## <a name="conditional-formatting"></a>条件格式

引入了区域的条件格式。 允许以下条件格式类型：

* 色阶
* 数据栏
* 图标集
* 自定义

此外：

* 返回应用条件格式的区域。
* 删除条件格式。
* 提供优先级和 stopifTrue 功能。
* 获取给定区域内所有条件格式的集合。
* 清除当前指定区域中处于活动状态的所有条件格式。

## <a name="api-list"></a>API 列表

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|在下一次调用“context.sync()”前暂停计算。设置后，开发者负责重新计算工作簿，以确保传播所有依赖项。|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|表示此条件格式中的 Rule 对象。|
||[set (properties: CellValueConditionalFormat)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: CellValueConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.cellvalueconditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[CellValueConditionalFormatData](/javascript/api/excel/excel.cellvalueconditionalformatdata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatdata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.cellvalueconditionalformatdata#rule)|表示此条件格式中的 Rule 对象。|
|[CellValueConditionalFormatLoadOptions](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.cellvalueconditionalformatloadoptions#rule)|表示此条件格式中的 Rule 对象。|
|[CellValueConditionalFormatUpdateData](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata)|[format](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.cellvalueconditionalformatupdatedata#rule)|表示此条件格式中的 Rule 对象。|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|色阶的条件。 使用两个点的色阶时, 中点是可选的。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|如果为 true, 则色阶将具有三个点 (最小值、中点和最大值), 否则它将有两个 (最小值, 最大值)。|
||[set (properties: ColorScaleConditionalFormat)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ColorScaleConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.colorscaleconditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ColorScaleConditionalFormatData](/javascript/api/excel/excel.colorscaleconditionalformatdata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatdata#criteria)|色阶的条件。 使用两个点的色阶时, 中点是可选的。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformatdata#threecolorscale)|如果为 true, 则色阶将具有三个点 (最小值、中点和最大值), 否则它将有两个 (最小值, 最大值)。|
|[ColorScaleConditionalFormatLoadOptions](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#criteria)|色阶的条件。 使用两个点的色阶时, 中点是可选的。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformatloadoptions#threecolorscale)|如果为 true, 则色阶将具有三个点 (最小值、中点和最大值), 否则它将有两个 (最小值, 最大值)。|
|[ColorScaleConditionalFormatUpdateData](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformatupdatedata#criteria)|色阶的条件。 使用两个点的色阶时, 中点是可选的。|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|如果需要，公式可对条件格式规则进行求值。|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|如果需要，公式可对条件格式规则进行求值。|
||[接线员](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|文本条件格式的运算符。|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|最大点色阶条件。|
||[放置](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|色阶为 3 色阶时的中点色阶条件。|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|最小点色阶条件。|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|色阶颜色的 HTML 颜色代码表示形式。 例如， #FF0000 代表红色。|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|数字、公式或 null（如果类型为 LowestValue）。|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|条件条件公式应基于什么。|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|该布尔值表示负 DataBar 是否与正 DataBar 具有相同边框颜色。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|该布尔值表示负 DataBar 是否与正 DataBar 具有相同填充颜色。|
||[set (properties: ConditionalDataBarNegativeFormat)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ConditionalDataBarNegativeFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.conditionaldatabarnegativeformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ConditionalDataBarNegativeFormatData](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#bordercolor)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#fillcolor)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivebordercolor)|该布尔值表示负 DataBar 是否与正 DataBar 具有相同边框颜色。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatdata#matchpositivefillcolor)|该布尔值表示负 DataBar 是否与正 DataBar 具有相同填充颜色。|
|[ConditionalDataBarNegativeFormatLoadOptions](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#bordercolor)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#fillcolor)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivebordercolor)|该布尔值表示负 DataBar 是否与正 DataBar 具有相同边框颜色。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatloadoptions#matchpositivefillcolor)|该布尔值表示负 DataBar 是否与正 DataBar 具有相同填充颜色。|
|[ConditionalDataBarNegativeFormatUpdateData](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#bordercolor)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#fillcolor)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivebordercolor)|该布尔值表示负 DataBar 是否与正 DataBar 具有相同边框颜色。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformatupdatedata#matchpositivefillcolor)|该布尔值表示负 DataBar 是否与正 DataBar 具有相同填充颜色。|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|该布尔值表示 DataBar 是否具有渐变。|
||[set (properties: ConditionalDataBarPositiveFormat)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ConditionalDataBarPositiveFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.conditionaldatabarpositiveformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ConditionalDataBarPositiveFormatData](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#bordercolor)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#fillcolor)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatdata#gradientfill)|该布尔值表示 DataBar 是否具有渐变。|
|[ConditionalDataBarPositiveFormatLoadOptions](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions)|[$all](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#$all)||
||[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#bordercolor)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#fillcolor)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatloadoptions#gradientfill)|该布尔值表示 DataBar 是否具有渐变。|
|[ConditionalDataBarPositiveFormatUpdateData](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#bordercolor)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#fillcolor)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformatupdatedata#gradientfill)|该布尔值表示 DataBar 是否具有渐变。|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|如果需要，公式可对 databar 规则进行求值。|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Databar 的规则类型。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|删除此条件格式。|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|返回应用条件格式的范围。 如果将条件格式应用于多个区域, 则会引发错误。 只读。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|返回条件格式应用于的区域; 或者, 如果将条件格式应用于多个区域, 则返回 null 对象。 只读。|
||[首选](/javascript/api/excel/excel.conditionalformat#priority)|条件格式集合中当前存在此条件格式的优先级 (或索引)。 同时更改此|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|如果当前条件格式为 CellValue 类型, 则返回单元格值条件格式属性。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|如果当前条件格式为 CellValue 类型, 则返回单元格值条件格式属性。|
||[色阶](/javascript/api/excel/excel.conditionalformat#colorscale)|如果当前条件格式为色阶类型, 则返回色阶条件格式属性。 只读。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|如果当前条件格式为色阶类型, 则返回色阶条件格式属性。 只读。|
||[自](/javascript/api/excel/excel.conditionalformat#custom)|如果当前条件格式为自定义类型, 则返回自定义条件格式属性。 只读。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|如果当前条件格式为自定义类型, 则返回自定义条件格式属性。 只读。|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|如果当前条件格式为数据栏, 则返回数据条属性。 只读。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|如果当前条件格式为数据栏, 则返回数据条属性。 只读。|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。 只读。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。 只读。|
||[id](/javascript/api/excel/excel.conditionalformat#id)|当前 ConditionalFormatCollection 内的条件格式的优先级。 只读。|
||[好](/javascript/api/excel/excel.conditionalformat#preset)|返回预设条件的条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|返回预设条件的条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|如果当前条件格式为 TopBottom 类型, 则返回 Top/底端条件格式属性。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|如果当前条件格式为 TopBottom 类型, 则返回 Top/底端条件格式属性。|
||[type](/javascript/api/excel/excel.conditionalformat#type)|一种条件格式。 一次只能设置一个。 只读。|
||[set (properties: ConditionalFormat)](/javascript/api/excel/excel.conditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.conditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add (type: "Custom" \| "DataBar" \|色阶\| "" \| IconSet "" TopBottom " \| " PresetCriteria " \| " ContainsText " \| " "CellValue")](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|将新的条件格式添加到集合中的第一个/最高优先级处。|
||[add (type: ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|将新的条件格式添加到集合中的第一个/最高优先级处。|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|清除当前指定区域中处于活动状态的所有条件格式。|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|返回工作簿中的条件格式数。 只读。|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|返回给定 ID 的条件格式。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|返回给定索引处的条件格式。|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|获取此集合中已加载的子项。|
|[ConditionalFormatCollectionLoadOptions](/javascript/api/excel/excel.conditionalformatcollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalue)|对于集合中的每一项: 如果当前条件格式为 CellValue 类型, 则返回 cell 值条件格式属性。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#cellvalueornullobject)|对于集合中的每一项: 如果当前条件格式为 CellValue 类型, 则返回 cell 值条件格式属性。|
||[色阶](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscale)|对于集合中的每一项: 如果当前条件格式为色阶类型, 则返回色阶条件格式属性。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#colorscaleornullobject)|对于集合中的每一项: 如果当前条件格式为色阶类型, 则返回色阶条件格式属性。|
||[自](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#custom)|对于集合中的每一项: 如果当前条件格式为自定义类型, 则返回自定义条件格式属性。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#customornullobject)|对于集合中的每一项: 如果当前条件格式为自定义类型, 则返回自定义条件格式属性。|
||[dataBar](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databar)|对于集合中的每一项: 如果当前条件格式为数据栏, 则返回数据栏属性。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#databarornullobject)|对于集合中的每一项: 如果当前条件格式为数据栏, 则返回数据栏属性。|
||[iconSet](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconset)|对于集合中的每一项: 如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#iconsetornullobject)|对于集合中的每一项: 如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。|
||[id](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#id)|对于集合中的每个项目: 当前 ConditionalFormatCollection 中条件格式的优先级。 只读。|
||[好](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#preset)|对于集合中的每一项: 返回预设条件条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#presetornullobject)|对于集合中的每一项: 返回预设条件条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[首选](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#priority)|对于集合中的每一项: 条件格式集合中当前存在此条件格式的优先级 (或索引)。 同时更改此|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#stopiftrue)|对于集合中的每个项目: 如果满足此条件格式的条件, 则不会对该单元格生效优先级较低的格式。|
||[textComparison](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparison)|对于集合中的每一项: 如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#textcomparisonornullobject)|对于集合中的每一项: 如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[topBottom](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottom)|对于集合中的每一项: 如果当前条件格式为 TopBottom 类型, 则返回顶部/底部条件格式属性。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#topbottomornullobject)|对于集合中的每一项: 如果当前条件格式为 TopBottom 类型, 则返回顶部/底部条件格式属性。|
||[type](/javascript/api/excel/excel.conditionalformatcollectionloadoptions#type)|对于集合中的每一项: 一种条件格式。 一次只能设置一个。 只读。|
|[ConditionalFormatData](/javascript/api/excel/excel.conditionalformatdata)|[cellValue](/javascript/api/excel/excel.conditionalformatdata#cellvalue)|如果当前条件格式为 CellValue 类型, 则返回单元格值条件格式属性。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatdata#cellvalueornullobject)|如果当前条件格式为 CellValue 类型, 则返回单元格值条件格式属性。|
||[色阶](/javascript/api/excel/excel.conditionalformatdata#colorscale)|如果当前条件格式为色阶类型, 则返回色阶条件格式属性。 只读。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatdata#colorscaleornullobject)|如果当前条件格式为色阶类型, 则返回色阶条件格式属性。 只读。|
||[自](/javascript/api/excel/excel.conditionalformatdata#custom)|如果当前条件格式为自定义类型, 则返回自定义条件格式属性。 只读。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatdata#customornullobject)|如果当前条件格式为自定义类型, 则返回自定义条件格式属性。 只读。|
||[dataBar](/javascript/api/excel/excel.conditionalformatdata#databar)|如果当前条件格式为数据栏, 则返回数据条属性。 只读。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatdata#databarornullobject)|如果当前条件格式为数据栏, 则返回数据条属性。 只读。|
||[iconSet](/javascript/api/excel/excel.conditionalformatdata#iconset)|如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。 只读。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatdata#iconsetornullobject)|如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。 只读。|
||[id](/javascript/api/excel/excel.conditionalformatdata#id)|当前 ConditionalFormatCollection 内的条件格式的优先级。 只读。|
||[好](/javascript/api/excel/excel.conditionalformatdata#preset)|返回预设条件的条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatdata#presetornullobject)|返回预设条件的条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[首选](/javascript/api/excel/excel.conditionalformatdata#priority)|条件格式集合中当前存在此条件格式的优先级 (或索引)。 同时更改此|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatdata#stopiftrue)|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|
||[textComparison](/javascript/api/excel/excel.conditionalformatdata#textcomparison)|如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatdata#textcomparisonornullobject)|如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[topBottom](/javascript/api/excel/excel.conditionalformatdata#topbottom)|如果当前条件格式为 TopBottom 类型, 则返回 Top/底端条件格式属性。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatdata#topbottomornullobject)|如果当前条件格式为 TopBottom 类型, 则返回 Top/底端条件格式属性。|
||[type](/javascript/api/excel/excel.conditionalformatdata#type)|一种条件格式。 一次只能设置一个。 只读。|
|[ConditionalFormatLoadOptions](/javascript/api/excel/excel.conditionalformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatloadoptions#$all)||
||[cellValue](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalue)|如果当前条件格式为 CellValue 类型, 则返回单元格值条件格式属性。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#cellvalueornullobject)|如果当前条件格式为 CellValue 类型, 则返回单元格值条件格式属性。|
||[色阶](/javascript/api/excel/excel.conditionalformatloadoptions#colorscale)|如果当前条件格式为色阶类型, 则返回色阶条件格式属性。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#colorscaleornullobject)|如果当前条件格式为色阶类型, 则返回色阶条件格式属性。|
||[自](/javascript/api/excel/excel.conditionalformatloadoptions#custom)|如果当前条件格式为自定义类型, 则返回自定义条件格式属性。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#customornullobject)|如果当前条件格式为自定义类型, 则返回自定义条件格式属性。|
||[dataBar](/javascript/api/excel/excel.conditionalformatloadoptions#databar)|如果当前条件格式为数据栏, 则返回数据条属性。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#databarornullobject)|如果当前条件格式为数据栏, 则返回数据条属性。|
||[iconSet](/javascript/api/excel/excel.conditionalformatloadoptions#iconset)|如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#iconsetornullobject)|如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。|
||[id](/javascript/api/excel/excel.conditionalformatloadoptions#id)|当前 ConditionalFormatCollection 内的条件格式的优先级。 只读。|
||[好](/javascript/api/excel/excel.conditionalformatloadoptions#preset)|返回预设条件的条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#presetornullobject)|返回预设条件的条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[首选](/javascript/api/excel/excel.conditionalformatloadoptions#priority)|条件格式集合中当前存在此条件格式的优先级 (或索引)。 同时更改此|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatloadoptions#stopiftrue)|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|
||[textComparison](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparison)|如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#textcomparisonornullobject)|如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[topBottom](/javascript/api/excel/excel.conditionalformatloadoptions#topbottom)|如果当前条件格式为 TopBottom 类型, 则返回 Top/底端条件格式属性。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatloadoptions#topbottomornullobject)|如果当前条件格式为 TopBottom 类型, 则返回 Top/底端条件格式属性。|
||[type](/javascript/api/excel/excel.conditionalformatloadoptions#type)|一种条件格式。 一次只能设置一个。 只读。|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|如果需要，公式可对条件格式规则进行求值。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|如果需要，公式可采用用户的语言对条件格式规则进行求值。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|如果需要，公式可采用 R1C1 表示法对条件格式规则进行求值。|
||[set (properties: ConditionalFormatRule)](/javascript/api/excel/excel.conditionalformatrule#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ConditionalFormatRuleUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.conditionalformatrule#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ConditionalFormatRuleData](/javascript/api/excel/excel.conditionalformatruledata)|[formula](/javascript/api/excel/excel.conditionalformatruledata#formula)|如果需要，公式可对条件格式规则进行求值。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruledata#formulalocal)|如果需要，公式可采用用户的语言对条件格式规则进行求值。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruledata#formular1c1)|如果需要，公式可采用 R1C1 表示法对条件格式规则进行求值。|
|[ConditionalFormatRuleLoadOptions](/javascript/api/excel/excel.conditionalformatruleloadoptions)|[$all](/javascript/api/excel/excel.conditionalformatruleloadoptions#$all)||
||[formula](/javascript/api/excel/excel.conditionalformatruleloadoptions#formula)|如果需要，公式可对条件格式规则进行求值。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleloadoptions#formulalocal)|如果需要，公式可采用用户的语言对条件格式规则进行求值。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleloadoptions#formular1c1)|如果需要，公式可采用 R1C1 表示法对条件格式规则进行求值。|
|[ConditionalFormatRuleUpdateData](/javascript/api/excel/excel.conditionalformatruleupdatedata)|[formula](/javascript/api/excel/excel.conditionalformatruleupdatedata#formula)|如果需要，公式可对条件格式规则进行求值。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatruleupdatedata#formulalocal)|如果需要，公式可采用用户的语言对条件格式规则进行求值。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatruleupdatedata#formular1c1)|如果需要，公式可采用 R1C1 表示法对条件格式规则进行求值。|
|[ConditionalFormatUpdateData](/javascript/api/excel/excel.conditionalformatupdatedata)|[cellValue](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalue)|如果当前条件格式为 CellValue 类型, 则返回单元格值条件格式属性。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#cellvalueornullobject)|如果当前条件格式为 CellValue 类型, 则返回单元格值条件格式属性。|
||[色阶](/javascript/api/excel/excel.conditionalformatupdatedata#colorscale)|如果当前条件格式为色阶类型, 则返回色阶条件格式属性。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#colorscaleornullobject)|如果当前条件格式为色阶类型, 则返回色阶条件格式属性。|
||[自](/javascript/api/excel/excel.conditionalformatupdatedata#custom)|如果当前条件格式为自定义类型, 则返回自定义条件格式属性。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#customornullobject)|如果当前条件格式为自定义类型, 则返回自定义条件格式属性。|
||[dataBar](/javascript/api/excel/excel.conditionalformatupdatedata#databar)|如果当前条件格式为数据栏, 则返回数据条属性。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#databarornullobject)|如果当前条件格式为数据栏, 则返回数据条属性。|
||[iconSet](/javascript/api/excel/excel.conditionalformatupdatedata#iconset)|如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#iconsetornullobject)|如果当前条件格式为 IconSet 类型, 则返回 IconSet 条件格式属性。|
||[好](/javascript/api/excel/excel.conditionalformatupdatedata#preset)|返回预设条件的条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#presetornullobject)|返回预设条件的条件格式。 有关更多详细信息, 请参阅 PresetCriteriaConditionalFormat。|
||[首选](/javascript/api/excel/excel.conditionalformatupdatedata#priority)|条件格式集合中当前存在此条件格式的优先级 (或索引)。 同时更改此|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformatupdatedata#stopiftrue)|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|
||[textComparison](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparison)|如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#textcomparisonornullobject)|如果当前条件格式是文本类型, 则返回特定的文本条件格式属性。|
||[topBottom](/javascript/api/excel/excel.conditionalformatupdatedata#topbottom)|如果当前条件格式为 TopBottom 类型, 则返回 Top/底端条件格式属性。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformatupdatedata#topbottomornullobject)|如果当前条件格式为 TopBottom 类型, 则返回 Top/底端条件格式属性。|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|如果与默认 IconSet 不同，返回当前条件的自定义图标，否则将返回 null。|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|取决于类型的数字或公式。|
||[接线员](/javascript/api/excel/excel.conditionaliconcriterion#operator)|图标条件格式的每个规则类型的 GreaterThan 或 GreaterThanOrEqual。|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|应基于的图标条件公式。|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[依据](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|条件格式的条件。|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|指示边框的特定边的常量值。 有关详细信息, 请参阅 ConditionalRangeBorderIndex。 只读。|
||[set (properties: ConditionalRangeBorder)](/javascript/api/excel/excel.conditionalrangeborder#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ConditionalRangeBorderUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.conditionalrangeborder#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息, 请参阅 BorderLineStyle。|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (index: "EdgeTop" \| "" EdgeBottom \| "" EdgeLeft \| "" EdgeRight ")](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|使用其名称获取 border 对象|
||[getItem (index: ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|使用其名称获取 border 对象|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|使用其索引获取 border 对象|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|获取下边框。 只读。|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|集合中的 border 对象数量。 只读。|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|获取此集合中已加载的子项。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|获取左边框。 只读。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|获取右边框。 只读。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|获取上边框。 只读。|
|[ConditionalRangeBorderCollectionLoadOptions](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#color)|对于集合中的每一项: HTML 颜色代码, 表示边框线的颜色, 格式 #RRGGBB (例如 "FFA500") 或命名的 HTML 颜色 (例如 "橙色")。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#sideindex)|对于集合中的每一项: 常量值, 用于指示边框的特定侧。 有关详细信息, 请参阅 ConditionalRangeBorderIndex。 只读。|
||[style](/javascript/api/excel/excel.conditionalrangebordercollectionloadoptions#style)|对于集合中的每个项目: 指定边框线条样式的线条样式的常量之一。 有关详细信息, 请参阅 BorderLineStyle。|
|[ConditionalRangeBorderCollectionUpdateData](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata)|[bottom](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#bottom)|获取下边框。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#left)|获取左边框。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#right)|获取右边框。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollectionupdatedata#top)|获取上边框。|
|[ConditionalRangeBorderData](/javascript/api/excel/excel.conditionalrangeborderdata)|[color](/javascript/api/excel/excel.conditionalrangeborderdata#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborderdata#sideindex)|指示边框的特定边的常量值。 有关详细信息, 请参阅 ConditionalRangeBorderIndex。 只读。|
||[style](/javascript/api/excel/excel.conditionalrangeborderdata#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息, 请参阅 BorderLineStyle。|
|[ConditionalRangeBorderLoadOptions](/javascript/api/excel/excel.conditionalrangeborderloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangeborderloadoptions#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborderloadoptions#sideindex)|指示边框的特定边的常量值。 有关详细信息, 请参阅 ConditionalRangeBorderIndex。 只读。|
||[style](/javascript/api/excel/excel.conditionalrangeborderloadoptions#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息, 请参阅 BorderLineStyle。|
|[ConditionalRangeBorderUpdateData](/javascript/api/excel/excel.conditionalrangeborderupdatedata)|[color](/javascript/api/excel/excel.conditionalrangeborderupdatedata#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[style](/javascript/api/excel/excel.conditionalrangeborderupdatedata#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息, 请参阅 BorderLineStyle。|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|重置填充。|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[set (properties: ConditionalRangeFill)](/javascript/api/excel/excel.conditionalrangefill#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ConditionalRangeFillUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.conditionalrangefill#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ConditionalRangeFillData](/javascript/api/excel/excel.conditionalrangefilldata)|[color](/javascript/api/excel/excel.conditionalrangefilldata#color)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
|[ConditionalRangeFillLoadOptions](/javascript/api/excel/excel.conditionalrangefillloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.conditionalrangefillloadoptions#color)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
|[ConditionalRangeFillUpdateData](/javascript/api/excel/excel.conditionalrangefillupdatedata)|[color](/javascript/api/excel/excel.conditionalrangefillupdatedata#color)|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|表示字体的加粗状态。|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|重置字体格式。|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|表示字体的斜体状态。|
||[set (properties: ConditionalRangeFont)](/javascript/api/excel/excel.conditionalrangefont#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ConditionalRangeFontUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.conditionalrangefont#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|表示字体的删除线状态。|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 ConditionalRangeFontUnderlineStyle。|
|[ConditionalRangeFontData](/javascript/api/excel/excel.conditionalrangefontdata)|[bold](/javascript/api/excel/excel.conditionalrangefontdata#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.conditionalrangefontdata#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.conditionalrangefontdata#italic)|表示字体的斜体状态。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontdata#strikethrough)|表示字体的删除线状态。|
||[underline](/javascript/api/excel/excel.conditionalrangefontdata#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 ConditionalRangeFontUnderlineStyle。|
|[ConditionalRangeFontLoadOptions](/javascript/api/excel/excel.conditionalrangefontloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.conditionalrangefontloadoptions#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.conditionalrangefontloadoptions#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.conditionalrangefontloadoptions#italic)|表示字体的斜体状态。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontloadoptions#strikethrough)|表示字体的删除线状态。|
||[underline](/javascript/api/excel/excel.conditionalrangefontloadoptions#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 ConditionalRangeFontUnderlineStyle。|
|[ConditionalRangeFontUpdateData](/javascript/api/excel/excel.conditionalrangefontupdatedata)|[bold](/javascript/api/excel/excel.conditionalrangefontupdatedata#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.conditionalrangefontupdatedata#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.conditionalrangefontupdatedata#italic)|表示字体的斜体状态。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefontupdatedata#strikethrough)|表示字体的删除线状态。|
||[underline](/javascript/api/excel/excel.conditionalrangefontupdatedata#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 ConditionalRangeFontUnderlineStyle。|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|表示给定范围的 Excel 数字格式代码。 如果传入 null, 则清除。|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|应用于整体条件格式范围的 border 对象的集合。 只读。|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|返回在整体条件格式范围上定义的 fill 对象。 只读。|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|返回在整体条件格式区域上定义的 font 对象。 只读。|
||[set (properties: ConditionalRangeFormat)](/javascript/api/excel/excel.conditionalrangeformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ConditionalRangeFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.conditionalrangeformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ConditionalRangeFormatData](/javascript/api/excel/excel.conditionalrangeformatdata)|[Borders](/javascript/api/excel/excel.conditionalrangeformatdata#borders)|应用于整体条件格式范围的 border 对象的集合。 只读。|
||[fill](/javascript/api/excel/excel.conditionalrangeformatdata#fill)|返回在整体条件格式范围上定义的 fill 对象。 只读。|
||[font](/javascript/api/excel/excel.conditionalrangeformatdata#font)|返回在整体条件格式区域上定义的 font 对象。 只读。|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatdata#numberformat)|表示给定范围的 Excel 数字格式代码。 如果传入 null, 则清除。|
|[ConditionalRangeFormatLoadOptions](/javascript/api/excel/excel.conditionalrangeformatloadoptions)|[$all](/javascript/api/excel/excel.conditionalrangeformatloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.conditionalrangeformatloadoptions#borders)|应用于整体条件格式范围的 border 对象的集合。|
||[fill](/javascript/api/excel/excel.conditionalrangeformatloadoptions#fill)|返回在整体条件格式范围上定义的 fill 对象。|
||[font](/javascript/api/excel/excel.conditionalrangeformatloadoptions#font)|返回在整体条件格式区域上定义的 font 对象。|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatloadoptions#numberformat)|表示给定范围的 Excel 数字格式代码。 如果传入 null, 则清除。|
|[ConditionalRangeFormatUpdateData](/javascript/api/excel/excel.conditionalrangeformatupdatedata)|[Borders](/javascript/api/excel/excel.conditionalrangeformatupdatedata#borders)|应用于整体条件格式范围的 border 对象的集合。|
||[fill](/javascript/api/excel/excel.conditionalrangeformatupdatedata#fill)|返回在整体条件格式范围上定义的 fill 对象。|
||[font](/javascript/api/excel/excel.conditionalrangeformatupdatedata#font)|返回在整体条件格式区域上定义的 font 对象。|
||[numberFormat](/javascript/api/excel/excel.conditionalrangeformatupdatedata#numberformat)|表示给定范围的 Excel 数字格式代码。 如果传入 null, 则清除。|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[接线员](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|文本条件格式的运算符。|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|条件格式的文本值。|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|1 和 1000 之间的数字排名或 1 和 100 之间的百分比排名。|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|根据顶部或底部排名设置值的格式。|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。 只读。|
||[标尺](/javascript/api/excel/excel.customconditionalformat#rule)|表示此条件格式中的 Rule 对象。 只读。|
||[set (properties: CustomConditionalFormat)](/javascript/api/excel/excel.customconditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: CustomConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.customconditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[CustomConditionalFormatData](/javascript/api/excel/excel.customconditionalformatdata)|[format](/javascript/api/excel/excel.customconditionalformatdata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。 只读。|
||[标尺](/javascript/api/excel/excel.customconditionalformatdata#rule)|表示此条件格式中的 Rule 对象。 只读。|
|[CustomConditionalFormatLoadOptions](/javascript/api/excel/excel.customconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.customconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.customconditionalformatloadoptions#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.customconditionalformatloadoptions#rule)|表示此条件格式中的 Rule 对象。|
|[CustomConditionalFormatUpdateData](/javascript/api/excel/excel.customconditionalformatupdatedata)|[format](/javascript/api/excel/excel.customconditionalformatupdatedata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.customconditionalformatupdatedata#rule)|表示此条件格式中的 Rule 对象。|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|表示窗体 #RRGGBB（例如 "FFA500"）的轴行颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|为 Excel 数据栏确定轴的方式的表示形式。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|表示数据条图形应基于的方向。|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|构成数据栏的下限（以及如何计算，如果适用）的规则。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|在 Excel 数据栏中的轴左侧的所有值的表示形式。 只读。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|在 Excel 数据栏中的轴右侧的所有值的表示形式。 只读。|
||[set (properties: DataBarConditionalFormat)](/javascript/api/excel/excel.databarconditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: DataBarConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.databarconditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|如果为 true，则对应用数据栏的单元格隐藏值。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|构成数据栏的上限（以及如何计算，如果适用）的规则。|
|[DataBarConditionalFormatData](/javascript/api/excel/excel.databarconditionalformatdata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatdata#axiscolor)|表示窗体 #RRGGBB（例如 "FFA500"）的轴行颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatdata#axisformat)|为 Excel 数据栏确定轴的方式的表示形式。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatdata#bardirection)|表示数据条图形应基于的方向。|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatdata#lowerboundrule)|构成数据栏的下限（以及如何计算，如果适用）的规则。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatdata#negativeformat)|在 Excel 数据栏中的轴左侧的所有值的表示形式。 只读。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatdata#positiveformat)|在 Excel 数据栏中的轴右侧的所有值的表示形式。 只读。|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatdata#showdatabaronly)|如果为 true，则对应用数据栏的单元格隐藏值。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatdata#upperboundrule)|构成数据栏的上限（以及如何计算，如果适用）的规则。|
|[DataBarConditionalFormatLoadOptions](/javascript/api/excel/excel.databarconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.databarconditionalformatloadoptions#$all)||
||[axisColor](/javascript/api/excel/excel.databarconditionalformatloadoptions#axiscolor)|表示窗体 #RRGGBB（例如 "FFA500"）的轴行颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#axisformat)|为 Excel 数据栏确定轴的方式的表示形式。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatloadoptions#bardirection)|表示数据条图形应基于的方向。|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatloadoptions#lowerboundrule)|构成数据栏的下限（以及如何计算，如果适用）的规则。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#negativeformat)|在 Excel 数据栏中的轴左侧的所有值的表示形式。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatloadoptions#positiveformat)|在 Excel 数据栏中的轴右侧的所有值的表示形式。|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatloadoptions#showdatabaronly)|如果为 true，则对应用数据栏的单元格隐藏值。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatloadoptions#upperboundrule)|构成数据栏的上限（以及如何计算，如果适用）的规则。|
|[DataBarConditionalFormatUpdateData](/javascript/api/excel/excel.databarconditionalformatupdatedata)|[axisColor](/javascript/api/excel/excel.databarconditionalformatupdatedata#axiscolor)|表示窗体 #RRGGBB（例如 "FFA500"）的轴行颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#axisformat)|为 Excel 数据栏确定轴的方式的表示形式。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformatupdatedata#bardirection)|表示数据条图形应基于的方向。|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformatupdatedata#lowerboundrule)|构成数据栏的下限（以及如何计算，如果适用）的规则。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#negativeformat)|在 Excel 数据栏中的轴左侧的所有值的表示形式。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformatupdatedata#positiveformat)|在 Excel 数据栏中的轴右侧的所有值的表示形式。|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformatupdatedata#showdatabaronly)|如果为 true，则对应用数据栏的单元格隐藏值。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformatupdatedata#upperboundrule)|构成数据栏的上限（以及如何计算，如果适用）的规则。|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|用于条件图标的规则和潜在自定义图标的条件和 IconSets 的数组。 请注意, 对于第一个条件, 只有自定义图标可以修改, 而类型、公式和运算符在设置时将被忽略。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|如果为 true, 则反转 IconSet 的图标订单。 请注意, 如果使用自定义图标, 则不能设置此设置。|
||[set (properties: IconSetConditionalFormat)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: IconSetConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.iconsetconditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|如果为 true，则隐藏值并仅显示图标。|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|如果设置, 则显示条件格式的 IconSet 选项。|
|[IconSetConditionalFormatData](/javascript/api/excel/excel.iconsetconditionalformatdata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatdata#criteria)|用于条件图标的规则和潜在自定义图标的条件和 IconSets 的数组。 请注意, 对于第一个条件, 只有自定义图标可以修改, 而类型、公式和运算符在设置时将被忽略。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatdata#reverseiconorder)|如果为 true, 则反转 IconSet 的图标订单。 请注意, 如果使用自定义图标, 则不能设置此设置。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatdata#showicononly)|如果为 true，则隐藏值并仅显示图标。|
||[style](/javascript/api/excel/excel.iconsetconditionalformatdata#style)|如果设置, 则显示条件格式的 IconSet 选项。|
|[IconSetConditionalFormatLoadOptions](/javascript/api/excel/excel.iconsetconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#criteria)|用于条件图标的规则和潜在自定义图标的条件和 IconSets 的数组。 请注意, 对于第一个条件, 只有自定义图标可以修改, 而类型、公式和运算符在设置时将被忽略。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#reverseiconorder)|如果为 true, 则反转 IconSet 的图标订单。 请注意, 如果使用自定义图标, 则不能设置此设置。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#showicononly)|如果为 true，则隐藏值并仅显示图标。|
||[style](/javascript/api/excel/excel.iconsetconditionalformatloadoptions#style)|如果设置, 则显示条件格式的 IconSet 选项。|
|[IconSetConditionalFormatUpdateData](/javascript/api/excel/excel.iconsetconditionalformatupdatedata)|[criteria](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#criteria)|用于条件图标的规则和潜在自定义图标的条件和 IconSets 的数组。 请注意, 对于第一个条件, 只有自定义图标可以修改, 而类型、公式和运算符在设置时将被忽略。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#reverseiconorder)|如果为 true, 则反转 IconSet 的图标订单。 请注意, 如果使用自定义图标, 则不能设置此设置。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#showicononly)|如果为 true，则隐藏值并仅显示图标。|
||[style](/javascript/api/excel/excel.iconsetconditionalformatupdatedata#style)|如果设置, 则显示条件格式的 IconSet 选项。|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|条件格式的规则。|
||[set (properties: PresetCriteriaConditionalFormat)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: PresetCriteriaConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.presetcriteriaconditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[PresetCriteriaConditionalFormatData](/javascript/api/excel/excel.presetcriteriaconditionalformatdata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.presetcriteriaconditionalformatdata#rule)|条件格式的规则。|
|[PresetCriteriaConditionalFormatLoadOptions](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.presetcriteriaconditionalformatloadoptions#rule)|条件格式的规则。|
|[PresetCriteriaConditionalFormatUpdateData](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.presetcriteriaconditionalformatupdatedata#rule)|条件格式的规则。|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|计算工作表上的单元格区域。|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|与该范围相交的 ConditionalFormats 的集合。 只读。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[conditionalFormats](/javascript/api/excel/excel.rangedata#conditionalformats)|与该范围相交的 ConditionalFormats 的集合。 只读。|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。 只读。|
||[标尺](/javascript/api/excel/excel.textconditionalformat#rule)|条件格式的规则。|
||[set (properties: TextConditionalFormat)](/javascript/api/excel/excel.textconditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TextConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.textconditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[TextConditionalFormatData](/javascript/api/excel/excel.textconditionalformatdata)|[format](/javascript/api/excel/excel.textconditionalformatdata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。 只读。|
||[标尺](/javascript/api/excel/excel.textconditionalformatdata#rule)|条件格式的规则。|
|[TextConditionalFormatLoadOptions](/javascript/api/excel/excel.textconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.textconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.textconditionalformatloadoptions#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.textconditionalformatloadoptions#rule)|条件格式的规则。|
|[TextConditionalFormatUpdateData](/javascript/api/excel/excel.textconditionalformatupdatedata)|[format](/javascript/api/excel/excel.textconditionalformatupdatedata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.textconditionalformatupdatedata#rule)|条件格式的规则。|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。 只读。|
||[标尺](/javascript/api/excel/excel.topbottomconditionalformat#rule)|顶部/底部条件格式的条件。|
||[set (properties: TopBottomConditionalFormat)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TopBottomConditionalFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.topbottomconditionalformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[TopBottomConditionalFormatData](/javascript/api/excel/excel.topbottomconditionalformatdata)|[format](/javascript/api/excel/excel.topbottomconditionalformatdata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。 只读。|
||[标尺](/javascript/api/excel/excel.topbottomconditionalformatdata#rule)|顶部/底部条件格式的条件。|
|[TopBottomConditionalFormatLoadOptions](/javascript/api/excel/excel.topbottomconditionalformatloadoptions)|[$all](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#$all)||
||[format](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.topbottomconditionalformatloadoptions#rule)|顶部/底部条件格式的条件。|
|[TopBottomConditionalFormatUpdateData](/javascript/api/excel/excel.topbottomconditionalformatupdatedata)|[format](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#format)|返回一个 format 对象, 该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.topbottomconditionalformatupdatedata#rule)|顶部/底部条件格式的条件。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[计算 (markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|计算工作表上的所有单元格。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
