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
# <a name="whats-new-in-excel-javascript-api-113"></a><span data-ttu-id="291f8-103">JavaScript API 1.13 Excel的新增功能</span><span class="sxs-lookup"><span data-stu-id="291f8-103">What's new in Excel JavaScript API 1.13</span></span>

<span data-ttu-id="291f8-104">ExcelApi 1.13 添加了一种方法，用于从 Base64 编码的字符串将工作表插入工作簿，并添加了一个事件来检测工作簿激活。</span><span class="sxs-lookup"><span data-stu-id="291f8-104">The ExcelApi 1.13 added a method to insert worksheets into a workbook from a Base64-encoded string and an event to detect workbook activation.</span></span> <span data-ttu-id="291f8-105">它还通过添加 API 跟踪对公式的更改并查找公式的直接从属单元格，增加了对范围中公式的支持。</span><span class="sxs-lookup"><span data-stu-id="291f8-105">It also increased support for formulas in ranges by adding APIs to track changes to formulas and locate a formula's direct dependent cells.</span></span> <span data-ttu-id="291f8-106">此外，它还通过添加用于替换文本、样式和空单元格管理的 PivotLayout API 来扩展数据透视表支持。</span><span class="sxs-lookup"><span data-stu-id="291f8-106">Additionally, it expanded PivotTable support by adding PivotLayout APIs for alt text, style, and empty cell management.</span></span>

| <span data-ttu-id="291f8-107">功能区域</span><span class="sxs-lookup"><span data-stu-id="291f8-107">Feature area</span></span> | <span data-ttu-id="291f8-108">说明</span><span class="sxs-lookup"><span data-stu-id="291f8-108">Description</span></span> | <span data-ttu-id="291f8-109">相关对象</span><span class="sxs-lookup"><span data-stu-id="291f8-109">Relevant objects</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="291f8-110">公式已更改事件</span><span class="sxs-lookup"><span data-stu-id="291f8-110">Formula changed events</span></span> | <span data-ttu-id="291f8-111">跟踪对公式的更改，包括导致更改的事件的源和类型。</span><span class="sxs-lookup"><span data-stu-id="291f8-111">Track changes to formulas, including the source and type of event that caused a change.</span></span> | [<span data-ttu-id="291f8-112">Worksheet.onFormulaChanged</span><span class="sxs-lookup"><span data-stu-id="291f8-112">Worksheet.onFormulaChanged</span></span>](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| <span data-ttu-id="291f8-113">公式从属单元格</span><span class="sxs-lookup"><span data-stu-id="291f8-113">Formula dependents</span></span> | <span data-ttu-id="291f8-114">查找公式的直接从属单元格。</span><span class="sxs-lookup"><span data-stu-id="291f8-114">Locate the direct dependent cells of a formula.</span></span> | [<span data-ttu-id="291f8-115">Range.getDirectDependents</span><span class="sxs-lookup"><span data-stu-id="291f8-115">Range.getDirectDependents</span></span>](/javascript/api/excel/excel.range#getDirectDependents__) |
| <span data-ttu-id="291f8-116">插入工作表</span><span class="sxs-lookup"><span data-stu-id="291f8-116">Insert worksheets</span></span> | <span data-ttu-id="291f8-117">将另一个工作簿中的工作表作为 Base64 编码的字符串插入到当前工作簿中。</span><span class="sxs-lookup"><span data-stu-id="291f8-117">Insert worksheets from another workbook into the current workbook as a Base64-encoded string.</span></span> | [<span data-ttu-id="291f8-118">Workbook.insertWorksheetsFromBase64</span><span class="sxs-lookup"><span data-stu-id="291f8-118">Workbook.insertWorksheetsFromBase64</span></span>](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_) |
| <span data-ttu-id="291f8-119">PivotTable PivotLayout</span><span class="sxs-lookup"><span data-stu-id="291f8-119">PivotTable PivotLayout</span></span> | <span data-ttu-id="291f8-120">PivotLayout 类的扩展，包括对替换文字和空单元格管理的新支持。</span><span class="sxs-lookup"><span data-stu-id="291f8-120">An expansion of the PivotLayout class, including new support for alt text and empty cell management.</span></span> | [<span data-ttu-id="291f8-121">PivotLayout</span><span class="sxs-lookup"><span data-stu-id="291f8-121">PivotLayout</span></span>](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a><span data-ttu-id="291f8-122">API 列表</span><span class="sxs-lookup"><span data-stu-id="291f8-122">API list</span></span>

<span data-ttu-id="291f8-123">下表列出了 JavaScript API 要求Excel集 1.13 中的 API。</span><span class="sxs-lookup"><span data-stu-id="291f8-123">The following table lists the APIs in Excel JavaScript API requirement set 1.13.</span></span> <span data-ttu-id="291f8-124">若要查看受 Excel JavaScript API 要求集 1.13 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)或更早中的 Excel API。</span><span class="sxs-lookup"><span data-stu-id="291f8-124">To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.13 or earlier, see [Excel APIs in requirement set 1.13 or earlier](/javascript/api/excel?view=excel-js-1.13&preserve-view=true).</span></span>

| <span data-ttu-id="291f8-125">类</span><span class="sxs-lookup"><span data-stu-id="291f8-125">Class</span></span> | <span data-ttu-id="291f8-126">域</span><span class="sxs-lookup"><span data-stu-id="291f8-126">Fields</span></span> | <span data-ttu-id="291f8-127">说明</span><span class="sxs-lookup"><span data-stu-id="291f8-127">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="291f8-128">FormulaChangedEventDetail</span><span class="sxs-lookup"><span data-stu-id="291f8-128">FormulaChangedEventDetail</span></span>](/javascript/api/excel/excel.formulachangedeventdetail)|[<span data-ttu-id="291f8-129">cellAddress</span><span class="sxs-lookup"><span data-stu-id="291f8-129">cellAddress</span></span>](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|<span data-ttu-id="291f8-130">包含已更改公式的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="291f8-130">The address of the cell that contains the changed formula.</span></span>|
||[<span data-ttu-id="291f8-131">previousFormula</span><span class="sxs-lookup"><span data-stu-id="291f8-131">previousFormula</span></span>](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|<span data-ttu-id="291f8-132">表示上一个公式，在更改之前。</span><span class="sxs-lookup"><span data-stu-id="291f8-132">Represents the previous formula, before it was changed.</span></span>|
|[<span data-ttu-id="291f8-133">InsertWorksheetOptions</span><span class="sxs-lookup"><span data-stu-id="291f8-133">InsertWorksheetOptions</span></span>](/javascript/api/excel/excel.insertworksheetoptions)|[<span data-ttu-id="291f8-134">positionType</span><span class="sxs-lookup"><span data-stu-id="291f8-134">positionType</span></span>](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|<span data-ttu-id="291f8-135">新工作表的当前工作簿中的插入位置。</span><span class="sxs-lookup"><span data-stu-id="291f8-135">The insert position, in the current workbook, of the new worksheets.</span></span>|
||[<span data-ttu-id="291f8-136">relativeTo</span><span class="sxs-lookup"><span data-stu-id="291f8-136">relativeTo</span></span>](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|<span data-ttu-id="291f8-137">引用参数的当前工作簿中的 `WorksheetPositionType` 工作表。</span><span class="sxs-lookup"><span data-stu-id="291f8-137">The worksheet in the current workbook that is referenced for the `WorksheetPositionType` parameter.</span></span>|
||[<span data-ttu-id="291f8-138">sheetNamesToInsert</span><span class="sxs-lookup"><span data-stu-id="291f8-138">sheetNamesToInsert</span></span>](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|<span data-ttu-id="291f8-139">要插入的单个工作表的名称。</span><span class="sxs-lookup"><span data-stu-id="291f8-139">The names of individual worksheets to insert.</span></span>|
|[<span data-ttu-id="291f8-140">PivotLayout</span><span class="sxs-lookup"><span data-stu-id="291f8-140">PivotLayout</span></span>](/javascript/api/excel/excel.pivotlayout)|[<span data-ttu-id="291f8-141">altTextDescription</span><span class="sxs-lookup"><span data-stu-id="291f8-141">altTextDescription</span></span>](/javascript/api/excel/excel.pivotlayout#alttextdescription)|<span data-ttu-id="291f8-142">数据透视表的替换文字说明。</span><span class="sxs-lookup"><span data-stu-id="291f8-142">The alt text description of the PivotTable.</span></span>|
||[<span data-ttu-id="291f8-143">altTextTitle</span><span class="sxs-lookup"><span data-stu-id="291f8-143">altTextTitle</span></span>](/javascript/api/excel/excel.pivotlayout#alttexttitle)|<span data-ttu-id="291f8-144">数据透视表的替换文字标题。</span><span class="sxs-lookup"><span data-stu-id="291f8-144">The alt text title of the PivotTable.</span></span>|
||[<span data-ttu-id="291f8-145">displayBlankLineAfterEachItem (显示：boolean) </span><span class="sxs-lookup"><span data-stu-id="291f8-145">displayBlankLineAfterEachItem(display: boolean)</span></span>](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|<span data-ttu-id="291f8-146">设置是否在每一项后显示一个空行。</span><span class="sxs-lookup"><span data-stu-id="291f8-146">Sets whether or not to display a blank line after each item.</span></span>|
||[<span data-ttu-id="291f8-147">emptyCellText</span><span class="sxs-lookup"><span data-stu-id="291f8-147">emptyCellText</span></span>](/javascript/api/excel/excel.pivotlayout#emptycelltext)|<span data-ttu-id="291f8-148">如果 为 ，则自动填充到数据透视表中任何空单元格中的文本 `fillEmptyCells == true` 。</span><span class="sxs-lookup"><span data-stu-id="291f8-148">The text that is automatically filled into any empty cell in the PivotTable if `fillEmptyCells == true`.</span></span>|
||[<span data-ttu-id="291f8-149">fillEmptyCells</span><span class="sxs-lookup"><span data-stu-id="291f8-149">fillEmptyCells</span></span>](/javascript/api/excel/excel.pivotlayout#fillemptycells)|<span data-ttu-id="291f8-150">指定是否应该使用 填充数据透视表中的空单元格 `emptyCellText` 。</span><span class="sxs-lookup"><span data-stu-id="291f8-150">Specifies whether empty cells in the PivotTable should be populated with the `emptyCellText`.</span></span>|
||[<span data-ttu-id="291f8-151">repeatAllItemLabels (repeatLabels：boolean) </span><span class="sxs-lookup"><span data-stu-id="291f8-151">repeatAllItemLabels(repeatLabels: boolean)</span></span>](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|<span data-ttu-id="291f8-152">设置数据透视表中所有字段的"重复所有项目标签"设置。</span><span class="sxs-lookup"><span data-stu-id="291f8-152">Sets the "repeat all item labels" setting across all fields in the PivotTable.</span></span>|
||[<span data-ttu-id="291f8-153">showFieldHeaders</span><span class="sxs-lookup"><span data-stu-id="291f8-153">showFieldHeaders</span></span>](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|<span data-ttu-id="291f8-154">指定数据透视表是否显示字段标题 (字段标题和筛选器下拉列表) 。</span><span class="sxs-lookup"><span data-stu-id="291f8-154">Specifies whether the PivotTable displays field headers (field captions and filter drop-downs).</span></span>|
|[<span data-ttu-id="291f8-155">PivotTable</span><span class="sxs-lookup"><span data-stu-id="291f8-155">PivotTable</span></span>](/javascript/api/excel/excel.pivottable)|[<span data-ttu-id="291f8-156">refreshOnOpen</span><span class="sxs-lookup"><span data-stu-id="291f8-156">refreshOnOpen</span></span>](/javascript/api/excel/excel.pivottable#refreshonopen)|<span data-ttu-id="291f8-157">指定工作簿打开时数据透视表是否刷新。</span><span class="sxs-lookup"><span data-stu-id="291f8-157">Specifies whether the PivotTable refreshes when the workbook opens.</span></span>|
|[<span data-ttu-id="291f8-158">Range</span><span class="sxs-lookup"><span data-stu-id="291f8-158">Range</span></span>](/javascript/api/excel/excel.range)|[<span data-ttu-id="291f8-159">getDirectDependents () </span><span class="sxs-lookup"><span data-stu-id="291f8-159">getDirectDependents()</span></span>](/javascript/api/excel/excel.range#getdirectdependents--)|<span data-ttu-id="291f8-160">返回一个对象，该对象表示包含同一工作表或多个工作表中单元格的所有直接从属 `WorkbookRangeAreas` 单元格的范围。</span><span class="sxs-lookup"><span data-stu-id="291f8-160">Returns a `WorkbookRangeAreas` object that represents the range containing all the direct dependents of a cell in the same worksheet or in multiple worksheets.</span></span>|
||[<span data-ttu-id="291f8-161">getExtendedRange (方向：Excel。KeyboardDirection， activeCell？： Range \| string) </span><span class="sxs-lookup"><span data-stu-id="291f8-161">getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)</span></span>](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|<span data-ttu-id="291f8-162">返回一个 range 对象，该对象包括当前区域以及区域边缘，根据提供的方向。</span><span class="sxs-lookup"><span data-stu-id="291f8-162">Returns a range object that includes the current range and up to the edge of the range, based on the provided direction.</span></span>|
||[<span data-ttu-id="291f8-163">getMergedAreasOrNullObject () </span><span class="sxs-lookup"><span data-stu-id="291f8-163">getMergedAreasOrNullObject()</span></span>](/javascript/api/excel/excel.range#getmergedareasornullobject--)|<span data-ttu-id="291f8-164">返回一个 RangeAreas 对象，该对象代表此范围中的合并区域。</span><span class="sxs-lookup"><span data-stu-id="291f8-164">Returns a RangeAreas object that represents the merged areas in this range.</span></span>|
||[<span data-ttu-id="291f8-165">getRangeEdge (方向：Excel。KeyboardDirection， activeCell？： Range \| string) </span><span class="sxs-lookup"><span data-stu-id="291f8-165">getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)</span></span>](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|<span data-ttu-id="291f8-166">返回一个 range 对象，该对象是数据区域的边缘单元格，对应于提供的方向。</span><span class="sxs-lookup"><span data-stu-id="291f8-166">Returns a range object that is the edge cell of the data region that corresponds to the provided direction.</span></span>|
|[<span data-ttu-id="291f8-167">Table</span><span class="sxs-lookup"><span data-stu-id="291f8-167">Table</span></span>](/javascript/api/excel/excel.table)|[<span data-ttu-id="291f8-168">resize (newRange：Range \| string) </span><span class="sxs-lookup"><span data-stu-id="291f8-168">resize(newRange: Range \| string)</span></span>](/javascript/api/excel/excel.table#resize-newrange-)|<span data-ttu-id="291f8-169">将表格调整到新区域。</span><span class="sxs-lookup"><span data-stu-id="291f8-169">Resize the table to the new range.</span></span>|
|[<span data-ttu-id="291f8-170">Workbook</span><span class="sxs-lookup"><span data-stu-id="291f8-170">Workbook</span></span>](/javascript/api/excel/excel.workbook)|[<span data-ttu-id="291f8-171">insertWorksheetsFromBase64 (base64File： string， options？： Excel。InsertWorksheetOptions) </span><span class="sxs-lookup"><span data-stu-id="291f8-171">insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)</span></span>](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|<span data-ttu-id="291f8-172">将源工作簿中的指定工作表插入到当前工作簿中。</span><span class="sxs-lookup"><span data-stu-id="291f8-172">Inserts the specified worksheets from a source workbook into the current workbook.</span></span>|
||[<span data-ttu-id="291f8-173">onActivated</span><span class="sxs-lookup"><span data-stu-id="291f8-173">onActivated</span></span>](/javascript/api/excel/excel.workbook#onactivated)|<span data-ttu-id="291f8-174">在激活工作簿时发生。</span><span class="sxs-lookup"><span data-stu-id="291f8-174">Occurs when the the workbook is activated.</span></span>|
|[<span data-ttu-id="291f8-175">WorkbookActivatedEventArgs</span><span class="sxs-lookup"><span data-stu-id="291f8-175">WorkbookActivatedEventArgs</span></span>](/javascript/api/excel/excel.workbookactivatedeventargs)|[<span data-ttu-id="291f8-176">type</span><span class="sxs-lookup"><span data-stu-id="291f8-176">type</span></span>](/javascript/api/excel/excel.workbookactivatedeventargs#type)|<span data-ttu-id="291f8-177">获取事件的类型。</span><span class="sxs-lookup"><span data-stu-id="291f8-177">Gets the type of the event.</span></span>|
|[<span data-ttu-id="291f8-178">Worksheet</span><span class="sxs-lookup"><span data-stu-id="291f8-178">Worksheet</span></span>](/javascript/api/excel/excel.worksheet)|[<span data-ttu-id="291f8-179">onFormulaChanged</span><span class="sxs-lookup"><span data-stu-id="291f8-179">onFormulaChanged</span></span>](/javascript/api/excel/excel.worksheet#onformulachanged)|<span data-ttu-id="291f8-180">在此工作表中更改一个或多个公式时发生。</span><span class="sxs-lookup"><span data-stu-id="291f8-180">Occurs when one or more formulas are changed in this worksheet.</span></span>|
|[<span data-ttu-id="291f8-181">WorksheetCollection</span><span class="sxs-lookup"><span data-stu-id="291f8-181">WorksheetCollection</span></span>](/javascript/api/excel/excel.worksheetcollection)|[<span data-ttu-id="291f8-182">onFormulaChanged</span><span class="sxs-lookup"><span data-stu-id="291f8-182">onFormulaChanged</span></span>](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|<span data-ttu-id="291f8-183">在此集合的任何工作表中更改一个或多个公式时发生。</span><span class="sxs-lookup"><span data-stu-id="291f8-183">Occurs when one or more formulas are changed in any worksheet of this collection.</span></span>|
|[<span data-ttu-id="291f8-184">WorksheetFormulaChangedEventArgs</span><span class="sxs-lookup"><span data-stu-id="291f8-184">WorksheetFormulaChangedEventArgs</span></span>](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[<span data-ttu-id="291f8-185">formulaDetails</span><span class="sxs-lookup"><span data-stu-id="291f8-185">formulaDetails</span></span>](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|<span data-ttu-id="291f8-186">获取对象 `FormulaChangedEventDetail` 数组，其中包含有关所有已更改公式的详细信息。</span><span class="sxs-lookup"><span data-stu-id="291f8-186">Gets an array of `FormulaChangedEventDetail` objects, which contain the details about the all of the changed formulas.</span></span>|
||[<span data-ttu-id="291f8-187">source</span><span class="sxs-lookup"><span data-stu-id="291f8-187">source</span></span>](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|<span data-ttu-id="291f8-188">事件的源。</span><span class="sxs-lookup"><span data-stu-id="291f8-188">The source of the event.</span></span>|
||[<span data-ttu-id="291f8-189">type</span><span class="sxs-lookup"><span data-stu-id="291f8-189">type</span></span>](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|<span data-ttu-id="291f8-190">获取事件的类型。</span><span class="sxs-lookup"><span data-stu-id="291f8-190">Gets the type of the event.</span></span>|
||[<span data-ttu-id="291f8-191">worksheetId</span><span class="sxs-lookup"><span data-stu-id="291f8-191">worksheetId</span></span>](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|<span data-ttu-id="291f8-192">获取公式发生更改的工作表的 ID。</span><span class="sxs-lookup"><span data-stu-id="291f8-192">Gets the ID of the worksheet in which the formula changed.</span></span>|

## <a name="see-also"></a><span data-ttu-id="291f8-193">另请参阅</span><span class="sxs-lookup"><span data-stu-id="291f8-193">See also</span></span>

- [<span data-ttu-id="291f8-194">Excel JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="291f8-194">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [<span data-ttu-id="291f8-195">Excel JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="291f8-195">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)