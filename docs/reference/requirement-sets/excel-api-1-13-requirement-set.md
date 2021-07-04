---
title: ExcelJavaScript API 要求集 1.13
description: 有关 ExcelApi 1.13 要求集的详细信息。
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: bfd9c23beda64565b44f16845e046fa1a2358d41
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290809"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>JavaScript API 1.13 Excel的新增功能

ExcelApi 1.13 添加了一种方法，用于从 Base64 编码的字符串将工作表插入工作簿，并添加了一个事件来检测工作簿激活。 它还通过添加 API 跟踪对公式的更改并查找公式的直接从属单元格，增加了对范围中公式的支持。 此外，它还通过添加用于替换文本、样式和空单元格管理的 PivotLayout API 来扩展数据透视表支持。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 公式已更改事件 | 跟踪对公式的更改，包括导致更改的事件的源和类型。 | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| 公式从属单元格 | 查找公式的直接从属单元格。 | [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__) |
| 插入工作表 | 将另一个工作簿中的工作表作为 Base64 编码的字符串插入到当前工作簿中。 | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_) |
| PivotTable PivotLayout | PivotLayout 类的扩展，包括对替换文字和空单元格管理的新支持。 | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求Excel集 1.13 中的 API。 若要查看受 Excel JavaScript API 要求集 1.13 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|包含已更改公式的单元格的地址。|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|表示上一个公式，在更改之前。|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|新工作表的当前工作簿中的插入位置。|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|引用参数的当前工作簿中的 `WorksheetPositionType` 工作表。|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|要插入的单个工作表的名称。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|数据透视表的替换文字说明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|数据透视表的替换文字标题。|
||[displayBlankLineAfterEachItem (显示：boolean) ](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|设置是否在每一项后显示一个空行。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|如果 为 ，则自动填充到数据透视表中任何空单元格中的文本 `fillEmptyCells == true` 。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|指定是否应该使用 填充数据透视表中的空单元格 `emptyCellText` 。|
||[repeatAllItemLabels (repeatLabels：boolean) ](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|设置数据透视表中所有字段的"重复所有项目标签"设置。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|指定数据透视表是否显示字段标题 (字段标题和筛选器下拉列表) 。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|指定工作簿打开时数据透视表是否刷新。|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents () ](/javascript/api/excel/excel.range#getdirectdependents--)|返回一个对象，该对象表示包含同一工作表或多个工作表中单元格的所有直接从属 `WorkbookRangeAreas` 单元格的范围。|
||[getExtendedRange (方向：Excel。KeyboardDirection， activeCell？： Range \| string) ](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|返回一个 range 对象，该对象包括当前区域以及区域边缘，根据提供的方向。|
||[getMergedAreasOrNullObject () ](/javascript/api/excel/excel.range#getmergedareasornullobject--)|返回一个 RangeAreas 对象，该对象代表此范围中的合并区域。|
||[getRangeEdge (方向：Excel。KeyboardDirection， activeCell？： Range \| string) ](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|返回一个 range 对象，该对象是数据区域的边缘单元格，对应于提供的方向。|
|[Table](/javascript/api/excel/excel.table)|[resize (newRange：Range \| string) ](/javascript/api/excel/excel.table#resize-newrange-)|将表格调整到新区域。|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64 (base64File： string， options？： Excel。InsertWorksheetOptions) ](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|将源工作簿中的指定工作表插入到当前工作簿中。|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|在激活工作簿时发生。|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|获取事件的类型。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|在此工作表中更改一个或多个公式时发生。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|在此集合的任何工作表中更改一个或多个公式时发生。|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|获取对象 `FormulaChangedEventDetail` 数组，其中包含有关所有已更改公式的详细信息。|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|获取公式发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
