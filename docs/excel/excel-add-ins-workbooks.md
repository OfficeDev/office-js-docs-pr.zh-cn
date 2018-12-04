---
title: 使用 Excel JavaScript API 处理工作簿
description: ''
ms.date: 11/27/2018
ms.openlocfilehash: 1cfde9bfdf306e35f47595f936679d9fa6e1814e
ms.sourcegitcommit: 026437bd3819f4e9cd4153ebe60c98ab04e18f4e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/30/2018
ms.locfileid: "27002336"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="c869a-102">使用 Excel JavaScript API 处理工作簿</span><span class="sxs-lookup"><span data-stu-id="c869a-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="c869a-103">本文提供了代码示例，介绍如何使用 Excel JavaScript API 对工作簿执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="c869a-103">This article provides code samples that show how to perform common tasks with tables using the Excel JavaScript API.</span></span> <span data-ttu-id="c869a-104">有关 **Workbook** 对象支持的属性和方法的完整列表，请参阅 [Workbook 对象 (Excel JavaScript API)](/javascript/api/excel/excel.workbook)。</span><span class="sxs-lookup"><span data-stu-id="c869a-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="c869a-105">此外，本文还介绍了通过 [Application](/javascript/api/excel/excel.application) 对象执行的工作簿级别的操作。</span><span class="sxs-lookup"><span data-stu-id="c869a-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="c869a-106">Workbook 对象是加载项与 Excel 交互的入口点。</span><span class="sxs-lookup"><span data-stu-id="c869a-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="c869a-107">它用于维护工作表、表、数据透视表等的集合，通过这些集合可以访问并更改 Excel 数据。</span><span class="sxs-lookup"><span data-stu-id="c869a-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="c869a-108">加载项可以通过 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) 对象访问单个工作表内的所有工作簿数据。</span><span class="sxs-lookup"><span data-stu-id="c869a-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through indivual worksheets.</span></span> <span data-ttu-id="c869a-109">具体来说，加载项可以借助它添加工作表、在工作表间导航并向工作表分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="c869a-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="c869a-110">[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md)一文介绍了如何访问并编辑工作表。</span><span class="sxs-lookup"><span data-stu-id="c869a-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="c869a-111">获取活动单元格或选定范围</span><span class="sxs-lookup"><span data-stu-id="c869a-111">Get the active cell or selected range</span></span>

<span data-ttu-id="c869a-112">Workbook 对象包含两种获取用户或加载项所选定单元格范围的方法：`getActiveCell()` 和 `getSelectedRange()`。</span><span class="sxs-lookup"><span data-stu-id="c869a-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="c869a-113">`getActiveCell()` 将活动单元格作为 [Range 对象](/javascript/api/excel/excel.range)来从工作簿中获取它。</span><span class="sxs-lookup"><span data-stu-id="c869a-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="c869a-114">下列示例演示对 `getActiveCell()` 的调用，紧随其后的是打印到控制台的单元格地址。</span><span class="sxs-lookup"><span data-stu-id="c869a-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c869a-115">`getSelectedRange()` 方法返回当前选定的单个范围。</span><span class="sxs-lookup"><span data-stu-id="c869a-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="c869a-116">若选定多个范围，将引发 InvalidSelection 错误。</span><span class="sxs-lookup"><span data-stu-id="c869a-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="c869a-117">下列示例演示对 `getSelectedRange()` 的调用，并且此方法随后会将相应范围的填充颜色设置为黄色。</span><span class="sxs-lookup"><span data-stu-id="c869a-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="c869a-118">创建工作簿</span><span class="sxs-lookup"><span data-stu-id="c869a-118">Create a workbook</span></span>

<span data-ttu-id="c869a-119">加载项可以新建一个工作簿，并独立于当前运行加载项的 Excel 实例。</span><span class="sxs-lookup"><span data-stu-id="c869a-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="c869a-120">Excel 对象包含的 `createWorkbook` 方法可用于实现此目的。</span><span class="sxs-lookup"><span data-stu-id="c869a-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="c869a-121">调用此方法时，会立即打开新的工作簿，并在新的 Excel 实例中显示它。</span><span class="sxs-lookup"><span data-stu-id="c869a-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="c869a-122">加载项保持打开状态，并随之前的工作簿一起运行。</span><span class="sxs-lookup"><span data-stu-id="c869a-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="c869a-123">此外，`createWorkbook` 方法还可以创建现有工作簿的副本。</span><span class="sxs-lookup"><span data-stu-id="c869a-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="c869a-124">此方法接受 .xlsx 文件的 base64 编码字符串表示形式作为可选参数。</span><span class="sxs-lookup"><span data-stu-id="c869a-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="c869a-125">若字符串参数为有效的 .xlsx 文件，则生成的工作簿为该文件的副本。</span><span class="sxs-lookup"><span data-stu-id="c869a-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="c869a-126">可以利用[文件切片](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)获取加载项的当前工作簿，作为一个 base64 编码字符串。</span><span class="sxs-lookup"><span data-stu-id="c869a-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="c869a-127">可以使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 类将文件转换为所需的 base64 编码字符串，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="c869a-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span> 

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var mybase64 = event.target.result.substr(startIndex + 7);

        Excel.createWorkbook(mybase64);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="c869a-128">保护工作簿的结构</span><span class="sxs-lookup"><span data-stu-id="c869a-128">Protect the workbook's structure</span></span>

<span data-ttu-id="c869a-129">加载项可以控制用户编辑工作簿结构的能力。</span><span class="sxs-lookup"><span data-stu-id="c869a-129">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="c869a-130">Workbook 对象的 `protection` 属性是一个包含 `protect()` 方法的 [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) 对象。</span><span class="sxs-lookup"><span data-stu-id="c869a-130">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="c869a-131">下列示例演示切换对工作簿结构的保护的基本方案。</span><span class="sxs-lookup"><span data-stu-id="c869a-131">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span> 

```js
Excel.run(function (context) {
    var workbook = context.workbook;
    workbook.load("protection/protected");

    return context.sync().then(function() {
        if (!workbook.protection.protected) {
            workbook.protection.protect();
        }
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c869a-132">`protect` 方法接受一个可选字符串参数。</span><span class="sxs-lookup"><span data-stu-id="c869a-132">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="c869a-133">此字符串表示用户要绕过保护并更改工作簿结构所需的密码。</span><span class="sxs-lookup"><span data-stu-id="c869a-133">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="c869a-134">此外，还可以在工作表级别设置保护，来防止不希望发生的数据编辑。</span><span class="sxs-lookup"><span data-stu-id="c869a-134">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="c869a-135">有关详细信息，请参阅[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md#data-protection)一文的“数据保护”部分。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="c869a-135">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE] 
> <span data-ttu-id="c869a-136">有关 Excel 中工作簿保护的详细信息，请参阅[保护工作簿](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)一文。</span><span class="sxs-lookup"><span data-stu-id="c869a-136">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="c869a-137">访问文档属性</span><span class="sxs-lookup"><span data-stu-id="c869a-137">Access document properties</span></span>

<span data-ttu-id="c869a-138">Workbook 对象可以访问 Office 文件元数据，即[文档属性](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)。</span><span class="sxs-lookup"><span data-stu-id="c869a-138">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="c869a-139">Workbook 对象的 `properties` 属性是一个包含这些元数据值的 [DocumentProperties](/javascript/api/excel/excel.documentproperties) 对象。</span><span class="sxs-lookup"><span data-stu-id="c869a-139">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="c869a-140">下列示例演示如何设置 author 属性。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="c869a-140">The following example shows how to set the **MetadataCatalogFileName** property declaratively.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c869a-141">此外，还可以定义自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c869a-141">You can also define custom properties.</span></span> <span data-ttu-id="c869a-142">DocumentProperties 对象保护 `custom` 属性，它表示用户定义的属性的键值对集合。</span><span class="sxs-lookup"><span data-stu-id="c869a-142">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="c869a-143">下列示例演示如何创建名称为“Introduction”且值为“Hello”的自定义属性，以及如何检索它。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="c869a-143">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

```js
Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    var customProperty = customDocProperties.getItem("Introduction");
    customProperty.load("key, value");

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="c869a-144">访问文档设置</span><span class="sxs-lookup"><span data-stu-id="c869a-144">Access document settings</span></span>

<span data-ttu-id="c869a-145">工作簿的设置类似于自定义属性集合。</span><span class="sxs-lookup"><span data-stu-id="c869a-145">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="c869a-146">区别在于：设置对于单个 Excel 文件和加载项配对而言是唯一的，而属性只是连接到文件。</span><span class="sxs-lookup"><span data-stu-id="c869a-146">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="c869a-147">下列示例演示如何创建并访问设置。</span><span class="sxs-lookup"><span data-stu-id="c869a-147">The following example shows how to create a file and add it to a folder.</span></span>

```js
Excel.run(function (context) {
    var settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    var needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    return context.sync().then(function() {
        console.log("Workbook needs review : " + needsReview.value);
    });
}).catch(errorHandlerFunction);
```

## <a name="control-calculation-behavior"></a><span data-ttu-id="c869a-148">控制计算行为</span><span class="sxs-lookup"><span data-stu-id="c869a-148">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="c869a-149">设置计算模式</span><span class="sxs-lookup"><span data-stu-id="c869a-149">Set calculation mode</span></span>

<span data-ttu-id="c869a-150">默认情况下，当引用的单元格发生更改时，Excel 会重新计算公式结果。</span><span class="sxs-lookup"><span data-stu-id="c869a-150">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="c869a-151">调整此计算行为可以改进加载项的性能。</span><span class="sxs-lookup"><span data-stu-id="c869a-151">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="c869a-152">Application 对象包含一个 `CalculationMode` 类型的 `calculationMode` 属性。</span><span class="sxs-lookup"><span data-stu-id="c869a-152">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="c869a-153">可以将此属性设置为下列值：</span><span class="sxs-lookup"><span data-stu-id="c869a-153">It can be set to the following values:</span></span>

 - <span data-ttu-id="c869a-154">`automatic`：默认的重新计算行为，每当相关数据发生更改时 Excel 都会计算新的公式结果。</span><span class="sxs-lookup"><span data-stu-id="c869a-154">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
 - <span data-ttu-id="c869a-155">`automaticExceptTables`：与 `automatic` 相同，但会忽略对表中值的任何更改。</span><span class="sxs-lookup"><span data-stu-id="c869a-155">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
 - <span data-ttu-id="c869a-156">`manual`：仅在用户或加载项请求计算时，才会进行计算。</span><span class="sxs-lookup"><span data-stu-id="c869a-156">`manual`: Calculations only occur when the user or add-in requests them.</span></span>
 
### <a name="set-calculation-type"></a><span data-ttu-id="c869a-157">设置计算类型</span><span class="sxs-lookup"><span data-stu-id="c869a-157">Set calculation type</span></span>

<span data-ttu-id="c869a-158">[Application](/javascript/api/excel/excel.application) 对象提供了一个用于强制立即进行重新计算的方法。</span><span class="sxs-lookup"><span data-stu-id="c869a-158">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="c869a-159">`Application.calculate(calculationType)` 将基于指定的 `calculationType` 启动手动重新计算。</span><span class="sxs-lookup"><span data-stu-id="c869a-159">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="c869a-160">可以指定下列值：</span><span class="sxs-lookup"><span data-stu-id="c869a-160">Specifies the operation to perform. The following table describes values that can be specified.</span></span>

 - <span data-ttu-id="c869a-161">`full`：重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。</span><span class="sxs-lookup"><span data-stu-id="c869a-161">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
 - <span data-ttu-id="c869a-162">`fullRebuild`：检查从属的公式，然后重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。</span><span class="sxs-lookup"><span data-stu-id="c869a-162">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
 - <span data-ttu-id="c869a-163">`recalculate`：重新计算所有活动工作簿中自上次计算后发生更改（或已以编程方式将其标记为重新计算目标）的公式，以及从属于它们的公式。</span><span class="sxs-lookup"><span data-stu-id="c869a-163">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>
 
> [!NOTE] 
> <span data-ttu-id="c869a-164">有关重新计算的详细信息，请参阅[更改公式重新计算、迭代或精度](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)一文。</span><span class="sxs-lookup"><span data-stu-id="c869a-164">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="c869a-165">暂停计算</span><span class="sxs-lookup"><span data-stu-id="c869a-165">Temporarily suspend calculations</span></span>

<span data-ttu-id="c869a-166">借助 Excel API，加载项还可以在调用 `RequestContext.sync()` 前禁用计算。</span><span class="sxs-lookup"><span data-stu-id="c869a-166">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="c869a-167">此操作通过 `suspendApiCalculationUntilNextSync()` 完成。</span><span class="sxs-lookup"><span data-stu-id="c869a-167">This is done to provide compatibility with InfoPath 2003.</span></span> <span data-ttu-id="c869a-168">加载项在编辑较大范围且无需访问两次编辑之间的数据时，使用此方法。</span><span class="sxs-lookup"><span data-stu-id="c869a-168">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="see-also"></a><span data-ttu-id="c869a-169">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c869a-169">See also</span></span>

- [<span data-ttu-id="c869a-170">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="c869a-170">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c869a-171">使用 Excel JavaScript API 处理工作表</span><span class="sxs-lookup"><span data-stu-id="c869a-171">Work with Worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="c869a-172">使用 Excel JavaScript API 处理特定范围</span><span class="sxs-lookup"><span data-stu-id="c869a-172">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)