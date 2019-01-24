---
title: 使用 Excel JavaScript API 对区域执行操作（基本）
description: ''
ms.date: 12/28/2018
localization_priority: Priority
ms.openlocfilehash: 505c22d2a3230aeafaf4d0c62a371a2ab93b3a9a
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386783"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="31920-102">使用 Excel JavaScript API 处理区域</span><span class="sxs-lookup"><span data-stu-id="31920-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="31920-103">本文中的代码示例展示了如何使用 Excel JavaScript API 对区域执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="31920-103">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="31920-104">有关 **Range** 对象支持的属性和方法的完整列表，请参阅 [Range 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="31920-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

> [!NOTE]
> <span data-ttu-id="31920-105">有关如何使用区域执行更高级任务的代码示例，请参阅 [使用 Excel JavaScript API 对区域执行操作（高级）](excel-add-ins-ranges-advanced.md)。</span><span class="sxs-lookup"><span data-stu-id="31920-105">For code samples that show how to perform more advanced tasks with ranges, see [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="31920-106">获取区域</span><span class="sxs-lookup"><span data-stu-id="31920-106">Get a range</span></span>

<span data-ttu-id="31920-107">下面的示例介绍了在工作表中获取对区域的引用的不同方法。</span><span class="sxs-lookup"><span data-stu-id="31920-107">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="31920-108">按地址获取区域</span><span class="sxs-lookup"><span data-stu-id="31920-108">Get range by address</span></span>

<span data-ttu-id="31920-109">下面的代码示例从名为 **Sample** 的工作表中获取地址为 **B2:B5** 的区域，加载其 **address** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="31920-109">The following code sample gets the range with address **B2:B5** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-range-by-name"></a><span data-ttu-id="31920-110">按名称获取区域</span><span class="sxs-lookup"><span data-stu-id="31920-110">Get range by name</span></span>

<span data-ttu-id="31920-111">下面的代码示例从名为 **Sample** 的工作表中获取名为 **MyRange** 的区域，加载其 **address** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="31920-111">The following code sample gets the range named **MyRange** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-used-range"></a><span data-ttu-id="31920-112">获取使用的区域</span><span class="sxs-lookup"><span data-stu-id="31920-112">Get used range</span></span>

<span data-ttu-id="31920-113">下面的代码示例从名为 **Sample** 的工作表中获取使用的区域，加载其 **address** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="31920-113">The following code sample gets the used range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span> <span data-ttu-id="31920-114">使用的区域是包含工作表中分配了值或格式的任意单元格的最小区域。</span><span class="sxs-lookup"><span data-stu-id="31920-114">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="31920-115">如果整个工作表为空，则 **getUsedRange()** 方法返回仅由工作表左上角单元格组成的区域。</span><span class="sxs-lookup"><span data-stu-id="31920-115">If the entire worksheet is blank, the **getUsedRange()** method returns a range that consists of only the top-left cell in the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-entire-range"></a><span data-ttu-id="31920-116">获取整个区域</span><span class="sxs-lookup"><span data-stu-id="31920-116">Get entire range</span></span>

<span data-ttu-id="31920-117">下面的代码示例从名为 **Sample** 的工作表中获取整个工作表区域，加载其 **address** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="31920-117">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="31920-118">插入多个单元格</span><span class="sxs-lookup"><span data-stu-id="31920-118">Insert a range of cells</span></span>

<span data-ttu-id="31920-119">下面的代码示例将多个单元格插入位置 **B4:E4**，并将其他单元格下移，以便为新的单元格提供空间。</span><span class="sxs-lookup"><span data-stu-id="31920-119">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-120">**插入区域之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-120">**Data before range is inserted**</span></span>

![Excel 中插入区域之前的数据](../images/excel-ranges-start.png)

<span data-ttu-id="31920-122">**插入区域之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-122">**Data after range is inserted**</span></span>

![Excel 中插入区域之后的数据](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="31920-124">清除多个单元格内容</span><span class="sxs-lookup"><span data-stu-id="31920-124">Clear a range of cells</span></span>

<span data-ttu-id="31920-125">下面的代码示例清除区域 **E2:E5** 中的所有内容和单元格格式设置。</span><span class="sxs-lookup"><span data-stu-id="31920-125">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-126">**清除区域之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-126">**Data before range is cleared**</span></span>

![Excel 中清除区域之前的数据](../images/excel-ranges-start.png)

<span data-ttu-id="31920-128">**清除区域之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-128">**Data after range is cleared**</span></span>

![Excel 中清除区域之后的数据](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="31920-130">删除多个单元格</span><span class="sxs-lookup"><span data-stu-id="31920-130">Delete a range of cells</span></span>

<span data-ttu-id="31920-131">下面的代码示例删除区域 **B4:E4** 中的单元格，并将其他单元格上移以填充删除的单元格空出的空间。</span><span class="sxs-lookup"><span data-stu-id="31920-131">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-132">**删除区域之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-132">**Data before range is deleted**</span></span>

![Excel 中删除区域之前的数据](../images/excel-ranges-start.png)

<span data-ttu-id="31920-134">**删除区域之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-134">**Data after range is deleted**</span></span>

![Excel 中删除区域之后的数据](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="31920-136">设置所选区域</span><span class="sxs-lookup"><span data-stu-id="31920-136">Set the selected range</span></span>

<span data-ttu-id="31920-137">下面的代码示例选择活动工作表中的区域 **B2:E6**。</span><span class="sxs-lookup"><span data-stu-id="31920-137">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-138">**选定的区域 B2:E6**</span><span class="sxs-lookup"><span data-stu-id="31920-138">**Selected range B2:E6**</span></span>

![Excel 中选定的区域](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="31920-140">获取所选区域</span><span class="sxs-lookup"><span data-stu-id="31920-140">Get the selected range</span></span>

<span data-ttu-id="31920-141">下面的代码示例获取所选区域，加载其 **address** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="31920-141">The following code sample gets the selected range, loads its **address** property, and writes a message to the console.</span></span> 

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-values-or-formulas"></a><span data-ttu-id="31920-142">设置值或公式</span><span class="sxs-lookup"><span data-stu-id="31920-142">Set values or formulas</span></span>

<span data-ttu-id="31920-143">下面的示例演示如何为单个单元格或多个单元格设置值和公式。</span><span class="sxs-lookup"><span data-stu-id="31920-143">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="31920-144">设置单个单元格的值</span><span class="sxs-lookup"><span data-stu-id="31920-144">Set value for a single cell</span></span>

<span data-ttu-id="31920-145">下面的代码示例将单元格 **C3** 的值设置为“5”，然后设置适合数据的最佳列宽。</span><span class="sxs-lookup"><span data-stu-id="31920-145">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-146">**更新单元格值之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-146">**Data before cell value is updated**</span></span>

![Excel 中更新单元格值之前的数据](../images/excel-ranges-set-start.png)

<span data-ttu-id="31920-148">**更新单元格值之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-148">**Data after cell value is updated**</span></span>

![Excel 中更新单元格值之后的数据](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="31920-150">设置多个单元格的值</span><span class="sxs-lookup"><span data-stu-id="31920-150">Set values for a range of cells</span></span>

<span data-ttu-id="31920-151">下面的代码示例为区域 **B5:D5** 中的单元格设置值，然后设置适合数据的最佳列宽。</span><span class="sxs-lookup"><span data-stu-id="31920-151">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];
    
    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-152">**更新多个单元格值之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-152">**Data before cell values are updated**</span></span>

![Excel 中更新多个单元格值之前的数据](../images/excel-ranges-set-start.png)

<span data-ttu-id="31920-154">**更新多个单元格值之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-154">**Data after cell values are updated**</span></span>

![Excel 中更新多个单元格值之后的数据](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="31920-156">设置单个单元格的公式</span><span class="sxs-lookup"><span data-stu-id="31920-156">Set formula for a single cell</span></span>

<span data-ttu-id="31920-157">下面的代码示例为单元格 **E3** 设置公式，然后设置适合数据的最佳列宽。</span><span class="sxs-lookup"><span data-stu-id="31920-157">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-158">**设置单元格公式之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-158">**Data before cell formula is set**</span></span>

![Excel 中设置单元格公式之前的数据](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="31920-160">**设置单元格公式之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-160">**Data after cell formula is set**</span></span>

![Excel 中设置单元格公式之后的数据](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="31920-162">设置多个单元格的公式</span><span class="sxs-lookup"><span data-stu-id="31920-162">Set formulas for a range of cells</span></span>

<span data-ttu-id="31920-163">下面的代码示例为区域 **E2:E6** 中的单元格设置公式，然后设置适合数据的最佳列宽。</span><span class="sxs-lookup"><span data-stu-id="31920-163">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];
    
    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-164">**设置多个单元格公式之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-164">**Data before cell formulas are set**</span></span>

![Excel 中设置多个单元格公式之前的数据](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="31920-166">**设置多个单元格公式之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-166">**Data after cell formulas are set**</span></span>

![Excel 中设置多个单元格公式之后的数据](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="31920-168">获取值、文本或公式</span><span class="sxs-lookup"><span data-stu-id="31920-168">Get values, text, or formulas</span></span>

<span data-ttu-id="31920-169">以下示例演示如何从多个单元格获取值、文本和公式。</span><span class="sxs-lookup"><span data-stu-id="31920-169">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="31920-170">从多个单元格获取值</span><span class="sxs-lookup"><span data-stu-id="31920-170">Get values from a range of cells</span></span>

<span data-ttu-id="31920-171">下面的代码示例获取区域 **B2:E6**，加载其 **values** 属性，并向控制台写入值。</span><span class="sxs-lookup"><span data-stu-id="31920-171">The following code sample gets the range **B2:E6**, loads its **values** property, and writes the values to the console.</span></span> <span data-ttu-id="31920-172">某个区域的 **values** 属性指定单元格包含的原始值。</span><span class="sxs-lookup"><span data-stu-id="31920-172">The **values** property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="31920-173">即使某个区域中的某些单元格包含公式，该区域的 **values** 属性仍会指定这些单元格的原始值，而不是任何公式。</span><span class="sxs-lookup"><span data-stu-id="31920-173">Even if some cells in a range contain formulas, the **values** property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-174">**区域中的数据（E 列中的值为公式的结果）**</span><span class="sxs-lookup"><span data-stu-id="31920-174">**Data in range (values in column E are a result of formulas)**</span></span>

![Excel 中设置多个单元格公式之后的数据](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="31920-176">**range.values（通过上面的代码示例记录到控制台）**</span><span class="sxs-lookup"><span data-stu-id="31920-176">**range.values (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="31920-177">从多个单元格获取文本</span><span class="sxs-lookup"><span data-stu-id="31920-177">Get text from a range of cells</span></span>

<span data-ttu-id="31920-178">下面的代码示例获取区域 **B2:E6**，加载其 **text** 属性，并向控制台写入该文本。</span><span class="sxs-lookup"><span data-stu-id="31920-178">The following code sample gets the range **B2:E6**, loads its **text** property, and writes it to the console.</span></span>  <span data-ttu-id="31920-179">区域的 **text** 属性指定该区域单元格的显示值。</span><span class="sxs-lookup"><span data-stu-id="31920-179">The **text** property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="31920-180">即使某个区域中的某些单元格包含公式，该区域的 **text** 属性仍会指定这些单元格的显示值，而不是任何公式。</span><span class="sxs-lookup"><span data-stu-id="31920-180">Even if some cells in a range contain formulas, the **text** property of the range specifies the display values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-181">**区域中的数据（E 列中的值为公式的结果）**</span><span class="sxs-lookup"><span data-stu-id="31920-181">**Data in range (values in column E are a result of formulas)**</span></span>

![Excel 中设置多个单元格公式之后的数据](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="31920-183">**range.text（通过上面的代码示例记录到控制台）**</span><span class="sxs-lookup"><span data-stu-id="31920-183">**range.text (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="31920-184">从多个单元格获取公式</span><span class="sxs-lookup"><span data-stu-id="31920-184">Get formulas from a range of cells</span></span>

<span data-ttu-id="31920-185">下面的代码示例获取区域 **B2:E6**，加载其 **formulas** 属性，并向控制台写入该公式。</span><span class="sxs-lookup"><span data-stu-id="31920-185">The following code sample gets the range **B2:E6**, loads its **formulas** property, and writes it to the console.</span></span>  <span data-ttu-id="31920-186">区域的 **formulas** 属性为包含公式的区域单元格指定公式，并为不包含公式的区域单元格指定原始值。</span><span class="sxs-lookup"><span data-stu-id="31920-186">The **formulas** property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-187">**区域中的数据（E 列中的值为公式的结果）**</span><span class="sxs-lookup"><span data-stu-id="31920-187">**Data in range (values in column E are a result of formulas)**</span></span>

![Excel 中设置多个单元格公式之后的数据](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="31920-189">**range.formulas（通过上面的代码示例记录到控制台）**</span><span class="sxs-lookup"><span data-stu-id="31920-189">**range.formulas (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="set-range-format"></a><span data-ttu-id="31920-190">设置区域格式</span><span class="sxs-lookup"><span data-stu-id="31920-190">Set range format</span></span>

<span data-ttu-id="31920-191">下面的示例演示如何为区域中的单元格设置字体颜色、填充颜色和数字格式。</span><span class="sxs-lookup"><span data-stu-id="31920-191">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="31920-192">设置字体颜色和填充颜色</span><span class="sxs-lookup"><span data-stu-id="31920-192">Set font color and fill color</span></span>

<span data-ttu-id="31920-193">下面的代码示例为区域 **B2:E2** 中的单元格设置字体颜色和填充颜色。</span><span class="sxs-lookup"><span data-stu-id="31920-193">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-194">**区域中设置字体颜色和填充颜色之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-194">**Data in range before font color and fill color are set**</span></span>

![Excel 中设置格式之前的数据](../images/excel-ranges-format-before.png)

<span data-ttu-id="31920-196">**区域中设置字体颜色和填充颜色之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-196">**Data in range after font color and fill color are set**</span></span>

![Excel 中设置格式之后的数据](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="31920-198">设置数字格式</span><span class="sxs-lookup"><span data-stu-id="31920-198">Set number format</span></span>

<span data-ttu-id="31920-199">下面的代码示例为区域 **D3:E5** 中的单元格设置数字格式。</span><span class="sxs-lookup"><span data-stu-id="31920-199">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-200">**区域中设置数字格式之前的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-200">**Data in range before number format is set**</span></span>

![Excel 中设置格式之前的数据](../images/excel-ranges-format-font-and-fill.png)

<span data-ttu-id="31920-202">**区域中设置数字格式之后的数据**</span><span class="sxs-lookup"><span data-stu-id="31920-202">**Data in range after number format is set**</span></span>

![设置格式后的 Excel 数据](../images/excel-ranges-format-numbers.png)

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="31920-204">范围的条件格式</span><span class="sxs-lookup"><span data-stu-id="31920-204">Conditional formatting of ranges</span></span>

<span data-ttu-id="31920-205">范围可以根据条件将格式应用于个别单元格。</span><span class="sxs-lookup"><span data-stu-id="31920-205">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="31920-206">有关此操作的详细信息，请参阅[将条件格式应用于 Excel 范围](excel-add-ins-conditional-formatting.md)。</span><span class="sxs-lookup"><span data-stu-id="31920-206">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="find-a-cell-using-string-matching-preview"></a><span data-ttu-id="31920-207">查找使用字符串匹配 （预览） 的单元格</span><span class="sxs-lookup"><span data-stu-id="31920-207">Find a cell using string matching (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="31920-208">区域对象的 `find` 函数当前仅适用于公共预览版（beta 版本）。</span><span class="sxs-lookup"><span data-stu-id="31920-208">The Range object's `find` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="31920-209">若要使用此功能，必须使用 Office.js CDN 的 beta 版库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="31920-209">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="31920-210">如果使用的是 TypeScript 或代码编辑器将 TypeScript 类型定义文件用于 IntelliSense，则使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。</span><span class="sxs-lookup"><span data-stu-id="31920-210">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="31920-211">`Range` 对象具有 `find` 方法在区域内搜索指定字符串。</span><span class="sxs-lookup"><span data-stu-id="31920-211">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="31920-212">返回有匹配文本的第一个单元格区域。</span><span class="sxs-lookup"><span data-stu-id="31920-212">It returns the range of the first cell with matching text.</span></span> <span data-ttu-id="31920-213">以下代码示例查找值等于字符串 **食品** 的第一个单元格，并将其地址记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="31920-213">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="31920-214">请注意，若指定的字符串不存在于区域中，`find` 将引发 `ItemNotFound` 错误。</span><span class="sxs-lookup"><span data-stu-id="31920-214">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="31920-215">若您预计到指定的字符串可能不存在区域中，则可使用 [findOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) 方法，以便您的代码可正常处理该情况。</span><span class="sxs-lookup"><span data-stu-id="31920-215">If you expect that the specified string may not exist in the range, use the [findOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="31920-216">在表示一个单元格的区域调用 `find` 方法时，将在整个工作表进行搜索。</span><span class="sxs-lookup"><span data-stu-id="31920-216">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="31920-217">搜索开始于该单元格，并按照 `SearchCriteria.searchDirection` 指定的方向进行，如有需要在工作表结束的地方换行。</span><span class="sxs-lookup"><span data-stu-id="31920-217">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="31920-218">另请参阅</span><span class="sxs-lookup"><span data-stu-id="31920-218">See also</span></span>

- [<span data-ttu-id="31920-219">使用 Excel JavaScript API 对区域执行操作（高级）</span><span class="sxs-lookup"><span data-stu-id="31920-219">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
- [<span data-ttu-id="31920-220">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="31920-220">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
