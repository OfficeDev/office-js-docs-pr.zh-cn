---
title: ExcelJavaScript API 要求集 1.6
description: 有关 ExcelApi 1.6 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 74fd184c4d65f30643266ba4799406cfe6f6e8e1e275af09eb6102e4bda986fa
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088977"
---
# <a name="whats-new-in-excel-javascript-api-16"></a>Excel JavaScript API 1.6 的最近更新

## <a name="conditional-formatting"></a>条件格式

引入区域的条件格式。 允许以下类型的条件格式。

- 色阶
- 数据栏
- 图标集
- 自定义

此外：

- 返回应用条件格式的区域。
- 删除条件格式。
- 提供优先级 `stopifTrue` 和功能。
- 获取给定区域内所有条件格式的集合。
- 清除当前指定区域中处于活动状态的所有条件格式。

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.6 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.6 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.6](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync () ](/javascript/api/excel/excel.application#suspendApiCalculationUntilNextSync__)|暂停计算，直到调用下 `context.sync()` 一个。|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|返回一个 format 对象，该对象封装了条件格式的字体、填充、边框和其他属性。|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|指定此条件格式的规则对象。|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|色标的条件。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threeColorScale)|如果为 ，色标将具有三个点 (最小值、中点、最大) ，否则色标将具有 (最小、 `true` 最大) 。|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|用于计算条件格式规则的公式（如果需要）。|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|用于计算条件格式规则的公式（如果需要）。|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|单元格值条件格式的运算符。|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|色标标准的最大点。|
||[midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|色标的中点（如果色标为 3 色色度）。|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|色标条件的最小点。|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|色标颜色格式的 HTML 颜色代码表示 (例如，#FF0000红色) 。|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|一个数字、公式或 `null` `type` (（ `lowestValue` 如果) ）。|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|条件条件公式应基于什么。|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#borderColor)|HTML 颜色代码，表示边框线条的颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillColor)|表示填充颜色的 HTML 颜色代码，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchPositiveBorderColor)|指定负数据条是否与正数据条具有相同的边框颜色。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchPositiveFillColor)|指定负数据条是否与正数据条具有相同的填充颜色。|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#borderColor)|HTML 颜色代码，表示边框线条的颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillColor)|表示填充颜色的 HTML 颜色代码，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientFill)|指定数据条是否具有渐变。|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|用于计算数据栏规则（如果需要）的公式。|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|数据栏的规则类型。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete__)|删除此条件格式。|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getRange__)|返回应用条件格式的范围。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getRangeOrNullObject__)|返回应用了二元格式的范围。|
||[priority](/javascript/api/excel/excel.conditionalformat#priority)|优先级 (索引) 当前存在此条件格式的条件格式集合中。|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellValue)|如果当前条件格式为类型，则返回单元格值条件格式 `CellValue` 属性。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellValueOrNullObject)|如果当前条件格式为类型，则返回单元格值条件格式 `CellValue` 属性。|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorScale)|如果当前条件格式为类型，则返回色标条件格式 `ColorScale` 属性。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorScaleOrNullObject)|如果当前条件格式为类型，则返回色标条件格式 `ColorScale` 属性。|
||[custom](/javascript/api/excel/excel.conditionalformat#custom)|如果当前的条件格式是自定义类型，则返回自定义条件格式属性。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customOrNullObject)|如果当前的条件格式是自定义类型，则返回自定义条件格式属性。|
||[dataBar](/javascript/api/excel/excel.conditionalformat#dataBar)|如果当前的条件格式是数据栏，则返回数据栏属性。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#dataBarOrNullObject)|如果当前的条件格式是数据栏，则返回数据栏属性。|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconSet)|如果当前条件格式是类型，则返回图标集条件格式 `IconSet` 属性。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconSetOrNullObject)|如果当前条件格式是类型，则返回图标集条件格式 `IconSet` 属性。|
||[id](/javascript/api/excel/excel.conditionalformat#id)|条件格式在当前 中的优先级 `ConditionalFormatCollection` 。|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|返回预设条件条件格式。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetOrNullObject)|返回预设条件条件格式。|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textComparison)|如果当前条件格式是文本类型，则返回特定的文本条件格式属性。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textComparisonOrNullObject)|如果当前条件格式是文本类型，则返回特定的文本条件格式属性。|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topBottom)|如果当前条件格式是类型，则返回顶部/底部条件格式 `TopBottom` 属性。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topBottomOrNullObject)|如果当前条件格式是类型，则返回顶部/底部条件格式 `TopBottom` 属性。|
||[type](/javascript/api/excel/excel.conditionalformat#type)|条件格式的类型。|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopIfTrue)|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[添加 (类型：Excel。ConditionalFormatType) ](/javascript/api/excel/excel.conditionalformatcollection#add_type_)|将新的条件格式添加到第一/第一优先级的集合。|
||[clearAll () ](/javascript/api/excel/excel.conditionalformatcollection#clearAll__)|清除当前指定区域中处于活动状态的所有条件格式。|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getCount__)|返回工作簿中条件格式的数量。|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItem_id_)|返回给定 ID 的条件格式。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getItemAt_index_)|返回给定索引处的条件格式。|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|获取此集合中已加载的子项。|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|用于计算条件格式规则的公式（如果需要）。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulaLocal)|公式（如果需要）用于计算用户语言的条件格式规则。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formulaR1C1)|公式（如果需要）用于计算 R1C1 样式表示法中的条件格式规则。|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customIcon)|如果不同于默认图标集，则返回当前条件自定义 `null` 图标。|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|取决于类型的数字或公式。|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|`greaterThan` 或 `greaterThanOrEqual` 图标条件格式的每个规则类型。|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|应基于的图标条件公式。|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[criterion](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|条件格式的条件。|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML 颜色代码，表示边框线条的颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideIndex)|指示边框的特定边的常量值。|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|线条样式的常量之一，指定边框的线条样式。|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (索引：Excel。ConditionalRangeBorderIndex) ](/javascript/api/excel/excel.conditionalrangebordercollection#getItem_index_)|使用其名称获取 border 对象|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getItemAt_index_)|使用其索引获取 border 对象|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|获取底部边框。|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|集合中的 border 对象数量。|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|获取此集合中已加载的子项。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|获取左边框。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|获取右边框。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|获取上边框。|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear__)|重置填充。|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML 颜色代码，表示填充的颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|指定字体是否加粗。|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear__)|重置字体格式。|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|文本颜色格式的 HTML 颜色代码表示 (例如，#FF0000红色) 。|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|指定字体是否为 italic。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|指定字体的删除线状态。|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|应用于字体的下划线类型。|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberFormat)|表示Excel区域的电话号码格式代码。|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|应用于整体条件格式范围的 border 对象的集合。|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|返回在整体条件格式范围内定义的 fill 对象。|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|返回在整体条件格式范围内定义的 font 对象。|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|文本条件格式的运算符。|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|条件格式的文本值。|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|1 和 1000 之间的数字排名或 1 和 100 之间的百分比排名。|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|根据排名第一或最后一位设置值的格式。|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|返回一个 format 对象，该对象封装了条件格式的字体、填充、边框和其他属性。|
||[rule](/javascript/api/excel/excel.customconditionalformat#rule)|指定此 `Rule` 条件格式的对象。|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axisColor)|HTML 颜色代码，表示轴线的颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisFormat)|如何为数据条确定坐标轴Excel表示。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#barDirection)|指定数据条图形应基于的方向。|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerBoundRule)|构成数据栏的下限（以及如何计算，如果适用）的规则。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeFormat)|数据条中轴左侧的所有Excel表示。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveFormat)|数据条中轴右侧所有值的Excel表示。|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showDataBarOnly)|如果 `true` 为 ，则隐藏应用数据栏的单元格的值。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperBoundRule)|构成数据栏的上限（以及如何计算，如果适用）的规则。|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|规则和条件图标的规则和图标集数组，以及条件图标的潜在自定义图标。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseIconOrder)|如果 `true` 为 ，则反转图标集的图标顺序。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showIconOnly)|如果 `true` 为 ，则隐藏值，并且只显示图标。|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|如果设置，则显示条件格式的图标集选项。|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|返回一个 format 对象，该对象封装了条件格式的字体、填充、边框和其他属性。|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|条件格式的规则。|
|[区域](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate__)|计算工作表上的单元格区域。|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalFormats)|与区域 `ConditionalFormats` 相交的集合。|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|返回一个 format 对象，该对象封装条件格式的字体、填充、边框和其他属性。|
||[rule](/javascript/api/excel/excel.textconditionalformat#rule)|条件格式的规则。|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|返回一个 format 对象，该对象封装条件格式的字体、填充、边框和其他属性。|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|顶部/底部条件格式的条件。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[calculate (markAllDirty： boolean) ](/javascript/api/excel/excel.worksheet#calculate_markAllDirty_)|计算工作表上的所有单元格。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.6&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
