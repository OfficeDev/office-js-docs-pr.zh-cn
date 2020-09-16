---
title: 使用 Excel JavaScript API 处理工作簿
description: 说明如何使用 Excel JavaScript API 对工作簿或应用程序级别的功能执行常见任务的代码示例。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: f0af6cc889a110406d987664575a6f3d1b30aa7b
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819502"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="33cab-103">使用 Excel JavaScript API 处理工作簿</span><span class="sxs-lookup"><span data-stu-id="33cab-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="33cab-104">本文提供了代码示例，介绍如何使用 Excel JavaScript API 对工作簿执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="33cab-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="33cab-105">有关该对象支持的属性和方法的完整列表 `Workbook` ，请参阅 [ (适用于 Excel 的 JavaScript API) 的工作簿对象 ](/javascript/api/excel/excel.workbook)。</span><span class="sxs-lookup"><span data-stu-id="33cab-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="33cab-106">此外，本文还介绍了通过 [Application](/javascript/api/excel/excel.application) 对象执行的工作簿级别的操作。</span><span class="sxs-lookup"><span data-stu-id="33cab-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="33cab-107">Workbook 对象是加载项与 Excel 交互的入口点。</span><span class="sxs-lookup"><span data-stu-id="33cab-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="33cab-108">它用于维护工作表、表、数据透视表等的集合，通过这些集合可以访问并更改 Excel 数据。</span><span class="sxs-lookup"><span data-stu-id="33cab-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="33cab-109">加载项可以通过 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) 对象访问单个工作表内的所有工作簿数据。</span><span class="sxs-lookup"><span data-stu-id="33cab-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="33cab-110">具体来说，加载项可以借助它添加工作表、在工作表间导航并向工作表分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="33cab-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="33cab-111">[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md)一文介绍了如何访问并编辑工作表。</span><span class="sxs-lookup"><span data-stu-id="33cab-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="33cab-112">获取活动单元格或选定范围</span><span class="sxs-lookup"><span data-stu-id="33cab-112">Get the active cell or selected range</span></span>

<span data-ttu-id="33cab-113">Workbook 对象包含两种获取用户或加载项所选定单元格范围的方法：`getActiveCell()` 和 `getSelectedRange()`。</span><span class="sxs-lookup"><span data-stu-id="33cab-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="33cab-114">`getActiveCell()` 将活动单元格作为 [Range 对象](/javascript/api/excel/excel.range)来从工作簿中获取它。</span><span class="sxs-lookup"><span data-stu-id="33cab-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="33cab-115">下列示例演示对 `getActiveCell()` 的调用，紧随其后的是打印到控制台的单元格地址。</span><span class="sxs-lookup"><span data-stu-id="33cab-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="33cab-116">`getSelectedRange()` 方法返回当前选定的单个范围。</span><span class="sxs-lookup"><span data-stu-id="33cab-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="33cab-117">若选定多个范围，将引发 InvalidSelection 错误。</span><span class="sxs-lookup"><span data-stu-id="33cab-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="33cab-118">下列示例演示对 `getSelectedRange()` 的调用，并且此方法随后会将相应范围的填充颜色设置为黄色。</span><span class="sxs-lookup"><span data-stu-id="33cab-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="33cab-119">创建工作簿</span><span class="sxs-lookup"><span data-stu-id="33cab-119">Create a workbook</span></span>

<span data-ttu-id="33cab-120">加载项可以新建一个工作簿，并独立于当前运行加载项的 Excel 实例。</span><span class="sxs-lookup"><span data-stu-id="33cab-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="33cab-121">Excel 对象包含的 `createWorkbook` 方法可用于实现此目的。</span><span class="sxs-lookup"><span data-stu-id="33cab-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="33cab-122">调用此方法时，会立即打开新的工作簿，并在新的 Excel 实例中显示它。</span><span class="sxs-lookup"><span data-stu-id="33cab-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="33cab-123">加载项保持打开状态，并随之前的工作簿一起运行。</span><span class="sxs-lookup"><span data-stu-id="33cab-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="33cab-124">此外，`createWorkbook` 方法还可以创建现有工作簿的副本。</span><span class="sxs-lookup"><span data-stu-id="33cab-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="33cab-125">此方法接受 .xlsx 文件的 base64 编码字符串表示形式作为可选参数。</span><span class="sxs-lookup"><span data-stu-id="33cab-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="33cab-126">若字符串参数为有效的 .xlsx 文件，则生成的工作簿为该文件的副本。</span><span class="sxs-lookup"><span data-stu-id="33cab-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="33cab-127">可以使用 [文件切片](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)以 base64 编码的字符串形式获取外接程序的当前工作簿。</span><span class="sxs-lookup"><span data-stu-id="33cab-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="33cab-128">可以使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 类将文件转换为所需的 base64 编码字符串，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="33cab-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="33cab-129">将现有工作簿副本插入到当前工作簿中（预览版）</span><span class="sxs-lookup"><span data-stu-id="33cab-129">Insert a copy of an existing workbook into the current one (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="33cab-130">`WorksheetCollection.addFromBase64` 方法当前仅在公共预览版中可用，并且仅适用于 Windows 和 Mac 上的 Office。</span><span class="sxs-lookup"><span data-stu-id="33cab-130">The `WorksheetCollection.addFromBase64` method is currently only available in public preview and only for Office on Windows and Mac.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="33cab-131">上一示例显示从现有工作簿创建的新工作簿。</span><span class="sxs-lookup"><span data-stu-id="33cab-131">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="33cab-132">此外，还可以将所有或部分现有工作簿复制到当前与加载项关联的工作簿中。</span><span class="sxs-lookup"><span data-stu-id="33cab-132">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="33cab-133">工作簿的 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) 可通过 `addFromBase64` 方法将目标工作簿的工作表副本插入到其本身。</span><span class="sxs-lookup"><span data-stu-id="33cab-133">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="33cab-134">其他工作簿文件将作为 base64 编码字符串传递，如 `Excel.createWorkbook` 调用一样。</span><span class="sxs-lookup"><span data-stu-id="33cab-134">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="33cab-135">在以下示例中，工作簿的工作表将插入到当前工作簿的活动工作表之后。</span><span class="sxs-lookup"><span data-stu-id="33cab-135">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="33cab-136">请注意，将为 `sheetNamesToInsert?: string[]` 参数传递 `null`。</span><span class="sxs-lookup"><span data-stu-id="33cab-136">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="33cab-137">这意味着将插入所有工作表。</span><span class="sxs-lookup"><span data-stu-id="33cab-137">This means all the worksheets are being inserted.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        var sheets = context.workbook.worksheets;
        sheets.addFromBase64(
            workbookContents,
            null, // get all the worksheets
            Excel.WorksheetPositionType.after, // insert them after the worksheet specified by the next parameter
            sheets.getActiveWorksheet() // insert them after the active worksheet
        );
        return context.sync();
    });
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="33cab-138">保护工作簿的结构</span><span class="sxs-lookup"><span data-stu-id="33cab-138">Protect the workbook's structure</span></span>

<span data-ttu-id="33cab-139">加载项可以控制用户编辑工作簿结构的能力。</span><span class="sxs-lookup"><span data-stu-id="33cab-139">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="33cab-140">Workbook 对象的 `protection` 属性是一个包含 `protect()` 方法的 [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) 对象。</span><span class="sxs-lookup"><span data-stu-id="33cab-140">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="33cab-141">下列示例演示切换对工作簿结构的保护的基本方案。</span><span class="sxs-lookup"><span data-stu-id="33cab-141">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="33cab-142">`protect` 方法接受一个可选字符串参数。</span><span class="sxs-lookup"><span data-stu-id="33cab-142">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="33cab-143">此字符串表示用户要绕过保护并更改工作簿结构所需的密码。</span><span class="sxs-lookup"><span data-stu-id="33cab-143">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="33cab-144">此外，还可以在工作表级别设置保护，来防止不希望发生的数据编辑。</span><span class="sxs-lookup"><span data-stu-id="33cab-144">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="33cab-145">有关详细信息，请参阅[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md#data-protection)一文的“数据保护”部分。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="33cab-145">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="33cab-146">有关 Excel 中工作簿保护的详细信息，请参阅[保护工作簿](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)一文。</span><span class="sxs-lookup"><span data-stu-id="33cab-146">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="33cab-147">访问文档属性</span><span class="sxs-lookup"><span data-stu-id="33cab-147">Access document properties</span></span>

<span data-ttu-id="33cab-148">Workbook 对象可以访问 Office 文件元数据，即[文档属性](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)。</span><span class="sxs-lookup"><span data-stu-id="33cab-148">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="33cab-149">Workbook 对象的 `properties` 属性是一个包含这些元数据值的 [DocumentProperties](/javascript/api/excel/excel.documentproperties) 对象。</span><span class="sxs-lookup"><span data-stu-id="33cab-149">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="33cab-150">下面的示例演示如何设置 `author` 属性。</span><span class="sxs-lookup"><span data-stu-id="33cab-150">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a><span data-ttu-id="33cab-151">自定义属性</span><span class="sxs-lookup"><span data-stu-id="33cab-151">Custom properties</span></span>

<span data-ttu-id="33cab-152">此外，还可以定义自定义属性。</span><span class="sxs-lookup"><span data-stu-id="33cab-152">You can also define custom properties.</span></span> <span data-ttu-id="33cab-153">DocumentProperties 对象保护 `custom` 属性，它表示用户定义的属性的键值对集合。</span><span class="sxs-lookup"><span data-stu-id="33cab-153">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="33cab-154">下列示例演示如何创建名称为“Introduction”且值为“Hello”的自定义属性，以及如何检索它。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="33cab-154">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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
    customProperty.load(["key, value"]);

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

#### <a name="worksheet-level-custom-properties"></a><span data-ttu-id="33cab-155">工作表级自定义属性</span><span class="sxs-lookup"><span data-stu-id="33cab-155">Worksheet-level custom properties</span></span>

<span data-ttu-id="33cab-156">此外，还可以在工作表级别设置自定义属性。</span><span class="sxs-lookup"><span data-stu-id="33cab-156">Custom properties can also be set at the worksheet level.</span></span> <span data-ttu-id="33cab-157">这些属性与文档级自定义属性相似，不同之处在于可以在不同的工作表中重复相同的键。</span><span class="sxs-lookup"><span data-stu-id="33cab-157">These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.</span></span> <span data-ttu-id="33cab-158">下面的示例演示如何使用当前工作表上的值 "Alpha" 创建名为 **WorksheetGroup** 的自定义属性，然后检索该属性。</span><span class="sxs-lookup"><span data-stu-id="33cab-158">The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.</span></span>

```js
Excel.run(function (context) {
    // Add the custom property.
    var customWorksheetProperties = context.workbook.worksheets.getActiveWorksheet().customProperties;
    customWorksheetProperties.add("WorksheetGroup", "Alpha");

    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    // Load the keys and values of all custom properties in the current worksheet.
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    var customWorksheetProperties = worksheet.customProperties;
    var customWorksheetProperty = customWorksheetProperties.getItem("WorksheetGroup");
    customWorksheetProperty.load(["key", "value"]);

    return context.sync().then(function() {
        // Log the WorksheetGroup custom property to the console.
        console.log(worksheet.name + ": " + customWorksheetProperty.key); // "WorksheetGroup"
        console.log("  Custom value : " + customWorksheetProperty.value); // "Alpha"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="33cab-159">访问文档设置</span><span class="sxs-lookup"><span data-stu-id="33cab-159">Access document settings</span></span>

<span data-ttu-id="33cab-160">工作簿的设置类似于自定义属性集合。</span><span class="sxs-lookup"><span data-stu-id="33cab-160">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="33cab-161">区别在于：设置对于单个 Excel 文件和加载项配对而言是唯一的，而属性只是连接到文件。</span><span class="sxs-lookup"><span data-stu-id="33cab-161">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="33cab-162">下列示例演示如何创建并访问设置。</span><span class="sxs-lookup"><span data-stu-id="33cab-162">The following example shows how to create and access a setting.</span></span>

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

## <a name="access-application-culture-settings"></a><span data-ttu-id="33cab-163">Access 应用程序的区域性设置</span><span class="sxs-lookup"><span data-stu-id="33cab-163">Access application culture settings</span></span>

<span data-ttu-id="33cab-164">工作簿具有可影响特定数据显示方式的语言和区域性设置。</span><span class="sxs-lookup"><span data-stu-id="33cab-164">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="33cab-165">当您的外接程序的用户在不同语言和区域性中共享工作簿时，这些设置可以帮助本地化数据。</span><span class="sxs-lookup"><span data-stu-id="33cab-165">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="33cab-166">您的外接程序可以使用字符串分析根据系统区域性设置本地化数字、日期和时间的格式，这样每个用户都可以看到自己的区域性格式的数据。</span><span class="sxs-lookup"><span data-stu-id="33cab-166">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="33cab-167">`Application.cultureInfo` 将系统区域性设置定义为 [CultureInfo](/javascript/api/excel/excel.cultureinfo) 对象。</span><span class="sxs-lookup"><span data-stu-id="33cab-167">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="33cab-168">这包含数字小数分隔符或日期格式等设置。</span><span class="sxs-lookup"><span data-stu-id="33cab-168">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="33cab-169">某些区域性设置可以 [通过 EXCEL UI 进行更改](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)。</span><span class="sxs-lookup"><span data-stu-id="33cab-169">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="33cab-170">系统设置将保留在对象中 `CultureInfo` 。</span><span class="sxs-lookup"><span data-stu-id="33cab-170">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="33cab-171">任何本地更改都将保留为 [应用程序](/javascript/api/excel/excel.application)级属性，例如 `Application.decimalSeparator` 。</span><span class="sxs-lookup"><span data-stu-id="33cab-171">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="33cab-172">下面的示例将数字字符串的十进制分隔符字符从 "，" 更改为系统设置所用的字符。</span><span class="sxs-lookup"><span data-stu-id="33cab-172">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var decimalSource = sheet.getRange("B2");
    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");

    return context.sync().then(function() {
        var systemDecimalSeparator =
            context.application.cultureInfo.numberFormat.numberDecimalSeparator;
        var oldDecimalString = decimalSource.values[0][0];

        // This assumes the input column is standardized to use "," as the decimal separator.
        var newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

        var resultRange = sheet.getRange("C2");
        resultRange.values = [[newDecimalString]];
        resultRange.format.autofitColumns();
        return context.sync();
    });
});
```

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="33cab-173">向工作簿添加自定义 XML 数据</span><span class="sxs-lookup"><span data-stu-id="33cab-173">Add custom XML data to the workbook</span></span>

<span data-ttu-id="33cab-174">通过 Excel 的 Open XML **.xlsx** 文件格式，可以让加载项将自定义 XML 数据嵌入到工作簿中。</span><span class="sxs-lookup"><span data-stu-id="33cab-174">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="33cab-175">此类数据将一直位于工作簿中，具体取决于加载项。</span><span class="sxs-lookup"><span data-stu-id="33cab-175">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="33cab-176">工作簿包含 [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)，它是一个 [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) 列表。</span><span class="sxs-lookup"><span data-stu-id="33cab-176">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="33cab-177">通过这些部件可以访问 XML 字符串并获得对应的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="33cab-177">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="33cab-178">将这些 ID 存储为设置后，加载项可以维护会话之间的 XML 部件密钥。</span><span class="sxs-lookup"><span data-stu-id="33cab-178">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="33cab-179">以下示例展示了如何使用自定义 XML 部件。</span><span class="sxs-lookup"><span data-stu-id="33cab-179">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="33cab-180">第一个代码块演示了如何将 XML 数据嵌入到文档中。</span><span class="sxs-lookup"><span data-stu-id="33cab-180">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="33cab-181">它将会存储一个审阅者列表，然后使用工作簿的设置保存 XML 的 `id`，以供后续检索。</span><span class="sxs-lookup"><span data-stu-id="33cab-181">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="33cab-182">第二个代码块演示后续如何访问该 XML。</span><span class="sxs-lookup"><span data-stu-id="33cab-182">The second block shows how to access that XML later.</span></span> <span data-ttu-id="33cab-183">“ContosoReviewXmlPartId”设置将被加载和传递到工作簿的 `customXmlParts`。</span><span class="sxs-lookup"><span data-stu-id="33cab-183">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="33cab-184">XML 数据随后将打印至控制台。</span><span class="sxs-lookup"><span data-stu-id="33cab-184">The XML data is then printed to the console.</span></span>

```js
Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    var originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    var customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");

    return context.sync().then(function() {
        // Store the XML part's ID in a setting
        var settings = context.workbook.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);
    });
}).catch(errorHandlerFunction);
```

```js
Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    var settings = context.workbook.settings;
    var xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    return context.sync().then(function () {
        if (xmlPartIDSetting.value) {
            var customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
            var xmlBlob = customXmlPart.getXml();

            return context.sync().then(function () {
                // Add spaces to make more human readable in the console
                var readableXML = xmlBlob.value.replace(/></g, "> <");
                console.log(readableXML);
            });
        }
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="33cab-185">仅当顶级自定义 XML 元素包含 `xmlns` 属性时才会填充 `CustomXMLPart.namespaceUri`。</span><span class="sxs-lookup"><span data-stu-id="33cab-185">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="33cab-186">控制计算行为</span><span class="sxs-lookup"><span data-stu-id="33cab-186">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="33cab-187">设置计算模式</span><span class="sxs-lookup"><span data-stu-id="33cab-187">Set calculation mode</span></span>

<span data-ttu-id="33cab-188">默认情况下，当引用的单元格发生更改时，Excel 会重新计算公式结果。</span><span class="sxs-lookup"><span data-stu-id="33cab-188">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="33cab-189">调整此计算行为可以改进加载项的性能。</span><span class="sxs-lookup"><span data-stu-id="33cab-189">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="33cab-190">Application 对象包含一个 `CalculationMode` 类型的 `calculationMode` 属性。</span><span class="sxs-lookup"><span data-stu-id="33cab-190">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="33cab-191">可以将此属性设置为下列值：</span><span class="sxs-lookup"><span data-stu-id="33cab-191">It can be set to the following values:</span></span>

- <span data-ttu-id="33cab-192">`automatic`：默认的重新计算行为，每当相关数据发生更改时 Excel 都会计算新的公式结果。</span><span class="sxs-lookup"><span data-stu-id="33cab-192">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="33cab-193">`automaticExceptTables`：与 `automatic` 相同，但会忽略对表中值的任何更改。</span><span class="sxs-lookup"><span data-stu-id="33cab-193">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="33cab-194">`manual`：仅在用户或加载项请求计算时，才会进行计算。</span><span class="sxs-lookup"><span data-stu-id="33cab-194">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="33cab-195">设置计算类型</span><span class="sxs-lookup"><span data-stu-id="33cab-195">Set calculation type</span></span>

<span data-ttu-id="33cab-196">[Application](/javascript/api/excel/excel.application) 对象提供了一个用于强制立即进行重新计算的方法。</span><span class="sxs-lookup"><span data-stu-id="33cab-196">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="33cab-197">`Application.calculate(calculationType)` 将基于指定的 `calculationType` 启动手动重新计算。</span><span class="sxs-lookup"><span data-stu-id="33cab-197">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="33cab-198">可以指定下列值：</span><span class="sxs-lookup"><span data-stu-id="33cab-198">The following values can be specified:</span></span>

- <span data-ttu-id="33cab-199">`full`：重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。</span><span class="sxs-lookup"><span data-stu-id="33cab-199">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="33cab-200">`fullRebuild`：检查从属的公式，然后重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。</span><span class="sxs-lookup"><span data-stu-id="33cab-200">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="33cab-201">`recalculate`：重新计算所有活动工作簿中自上次计算后发生更改（或已以编程方式将其标记为重新计算目标）的公式，以及从属于它们的公式。</span><span class="sxs-lookup"><span data-stu-id="33cab-201">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="33cab-202">有关重新计算的详细信息，请参阅[更改公式重新计算、迭代或精度](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)一文。</span><span class="sxs-lookup"><span data-stu-id="33cab-202">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="33cab-203">暂停计算</span><span class="sxs-lookup"><span data-stu-id="33cab-203">Temporarily suspend calculations</span></span>

<span data-ttu-id="33cab-204">借助 Excel API，加载项还可以在调用 `RequestContext.sync()` 前禁用计算。</span><span class="sxs-lookup"><span data-stu-id="33cab-204">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="33cab-205">此操作通过 `suspendApiCalculationUntilNextSync()` 完成。</span><span class="sxs-lookup"><span data-stu-id="33cab-205">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="33cab-206">加载项在编辑较大范围且无需访问两次编辑之间的数据时，使用此方法。</span><span class="sxs-lookup"><span data-stu-id="33cab-206">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="save-the-workbook"></a><span data-ttu-id="33cab-207">保存工作簿</span><span class="sxs-lookup"><span data-stu-id="33cab-207">Save the workbook</span></span>

<span data-ttu-id="33cab-208">`Workbook.save` 会将工作簿保存到永久存储区。</span><span class="sxs-lookup"><span data-stu-id="33cab-208">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="33cab-209">`save` 方法采用单个可选 `saveBehavior` 参数，该参数可为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="33cab-209">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="33cab-210">`Excel.SaveBehavior.save`（默认）：保存文件，但不提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="33cab-210">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="33cab-211">如果之前未保存文件，则文件保存到默认位置。</span><span class="sxs-lookup"><span data-stu-id="33cab-211">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="33cab-212">如果之前保存过文件，则保存到之前的位置。</span><span class="sxs-lookup"><span data-stu-id="33cab-212">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="33cab-213">`Excel.SaveBehavior.prompt`：如果之前未保存文件，则将提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="33cab-213">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="33cab-214">如果之前已保存文件，则保存到之前的位置且不提示用户。</span><span class="sxs-lookup"><span data-stu-id="33cab-214">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="33cab-215">如果提示用户保存并取消操作，则 `save` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="33cab-215">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="33cab-216">关闭工作簿</span><span class="sxs-lookup"><span data-stu-id="33cab-216">Close the workbook</span></span>

<span data-ttu-id="33cab-217">`Workbook.close` 会关闭工作簿，一并关闭与该工作簿关联的加载项（Excel 应用程序仍保持打开状态）。</span><span class="sxs-lookup"><span data-stu-id="33cab-217">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="33cab-218">`close` 方法采用单个可选 `closeBehavior` 参数，该参数可为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="33cab-218">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="33cab-219">`Excel.CloseBehavior.save`（默认）：在关闭前保存文件。</span><span class="sxs-lookup"><span data-stu-id="33cab-219">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="33cab-220">如果之前未保存文件，则将提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="33cab-220">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="33cab-221">`Excel.CloseBehavior.skipSave`：立即关闭文件但不保存。</span><span class="sxs-lookup"><span data-stu-id="33cab-221">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="33cab-222">所有未保存的更改均将丢失。</span><span class="sxs-lookup"><span data-stu-id="33cab-222">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="33cab-223">另请参阅</span><span class="sxs-lookup"><span data-stu-id="33cab-223">See also</span></span>

- [<span data-ttu-id="33cab-224">Office 外接程序中的 Excel JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="33cab-224">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="33cab-225">使用 Excel JavaScript API 处理工作表</span><span class="sxs-lookup"><span data-stu-id="33cab-225">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="33cab-226">使用 Excel JavaScript API 处理特定范围</span><span class="sxs-lookup"><span data-stu-id="33cab-226">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
