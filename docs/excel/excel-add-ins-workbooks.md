---
title: 使用 Excel JavaScript API 处理工作簿
description: 了解如何使用 JavaScript API 对工作簿或应用程序级别功能执行Excel任务。
ms.date: 06/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 638384a1e08af182db042638c655d8d74354c637
ms.sourcegitcommit: ba4fb7087b9841d38bb46a99a63e88df49514a4d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/05/2021
ms.locfileid: "52779346"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="fe505-103">使用 Excel JavaScript API 处理工作簿</span><span class="sxs-lookup"><span data-stu-id="fe505-103">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="fe505-104">本文提供了代码示例，介绍如何使用 Excel JavaScript API 对工作簿执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="fe505-104">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="fe505-105">有关对象支持的属性和方法的完整列表，请参阅 `Workbook` Workbook Object [ (JavaScript API for Excel) 。 ](/javascript/api/excel/excel.workbook)</span><span class="sxs-lookup"><span data-stu-id="fe505-105">For the complete list of properties and methods that the `Workbook` object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="fe505-106">此外，本文还介绍了通过 [Application](/javascript/api/excel/excel.application) 对象执行的工作簿级别的操作。</span><span class="sxs-lookup"><span data-stu-id="fe505-106">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="fe505-107">Workbook 对象是加载项与 Excel 交互的入口点。</span><span class="sxs-lookup"><span data-stu-id="fe505-107">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="fe505-108">它用于维护工作表、表、数据透视表等的集合，通过这些集合可以访问并更改 Excel 数据。</span><span class="sxs-lookup"><span data-stu-id="fe505-108">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="fe505-109">加载项可以通过 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) 对象访问单个工作表内的所有工作簿数据。</span><span class="sxs-lookup"><span data-stu-id="fe505-109">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="fe505-110">具体来说，加载项可以借助它添加工作表、在工作表间导航并向工作表分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="fe505-110">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="fe505-111">[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md)一文介绍了如何访问并编辑工作表。</span><span class="sxs-lookup"><span data-stu-id="fe505-111">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="fe505-112">获取活动单元格或选定范围</span><span class="sxs-lookup"><span data-stu-id="fe505-112">Get the active cell or selected range</span></span>

<span data-ttu-id="fe505-113">Workbook 对象包含两种获取用户或加载项所选定单元格范围的方法：`getActiveCell()` 和 `getSelectedRange()`。</span><span class="sxs-lookup"><span data-stu-id="fe505-113">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="fe505-114">`getActiveCell()` 将活动单元格作为 [Range 对象](/javascript/api/excel/excel.range)来从工作簿中获取它。</span><span class="sxs-lookup"><span data-stu-id="fe505-114">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="fe505-115">下列示例演示对 `getActiveCell()` 的调用，紧随其后的是打印到控制台的单元格地址。</span><span class="sxs-lookup"><span data-stu-id="fe505-115">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="fe505-116">`getSelectedRange()` 方法返回当前选定的单个范围。</span><span class="sxs-lookup"><span data-stu-id="fe505-116">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="fe505-117">若选定多个范围，将引发 InvalidSelection 错误。</span><span class="sxs-lookup"><span data-stu-id="fe505-117">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="fe505-118">下列示例演示对 `getSelectedRange()` 的调用，并且此方法随后会将相应范围的填充颜色设置为黄色。</span><span class="sxs-lookup"><span data-stu-id="fe505-118">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="fe505-119">创建工作簿</span><span class="sxs-lookup"><span data-stu-id="fe505-119">Create a workbook</span></span>

<span data-ttu-id="fe505-120">加载项可以新建一个工作簿，并独立于当前运行加载项的 Excel 实例。</span><span class="sxs-lookup"><span data-stu-id="fe505-120">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="fe505-121">Excel 对象包含的 `createWorkbook` 方法可用于实现此目的。</span><span class="sxs-lookup"><span data-stu-id="fe505-121">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="fe505-122">调用此方法时，会立即打开新的工作簿，并在新的 Excel 实例中显示它。</span><span class="sxs-lookup"><span data-stu-id="fe505-122">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="fe505-123">加载项保持打开状态，并随之前的工作簿一起运行。</span><span class="sxs-lookup"><span data-stu-id="fe505-123">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="fe505-124">此外，`createWorkbook` 方法还可以创建现有工作簿的副本。</span><span class="sxs-lookup"><span data-stu-id="fe505-124">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="fe505-125">此方法接受 .xlsx 文件的 base64 编码字符串表示形式作为可选参数。</span><span class="sxs-lookup"><span data-stu-id="fe505-125">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="fe505-126">若字符串参数为有效的 .xlsx 文件，则生成的工作簿为该文件的副本。</span><span class="sxs-lookup"><span data-stu-id="fe505-126">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="fe505-127">可以使用文件切片 将加载项的当前工作簿作为 base64 编码的 [字符串获取](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="fe505-127">You can get your add-in's current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="fe505-128">可以使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 类将文件转换为所需的 base64 编码字符串，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="fe505-128">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="fe505-129">将现有工作簿副本插入到当前工作簿中（预览版）</span><span class="sxs-lookup"><span data-stu-id="fe505-129">Insert a copy of an existing workbook into the current one (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="fe505-130">`Workbook.insertWorksheetsFromBase64`方法当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="fe505-130">The `Workbook.insertWorksheetsFromBase64` method is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="fe505-131">上一示例显示从现有工作簿创建的新工作簿。</span><span class="sxs-lookup"><span data-stu-id="fe505-131">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="fe505-132">此外，还可以将所有或部分现有工作簿复制到当前与加载项关联的工作簿中。</span><span class="sxs-lookup"><span data-stu-id="fe505-132">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="fe505-133">[Workbook](/javascript/api/excel/excel.workbook)具有将目标工作簿的工作表 `insertWorksheetsFromBase64` 副本插入自身的方法。</span><span class="sxs-lookup"><span data-stu-id="fe505-133">A [Workbook](/javascript/api/excel/excel.workbook) has the `insertWorksheetsFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="fe505-134">另一个工作簿的文件作为 base64 编码的字符串传递，就像调用 `Excel.createWorkbook` 一样。</span><span class="sxs-lookup"><span data-stu-id="fe505-134">The other workbook's file is passed as a base64-encoded string, just like the `Excel.createWorkbook` call.</span></span> 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> <span data-ttu-id="fe505-135">此方法 `insertWorksheetsFromBase64` 在 Excel、Windows 和 Web 上受支持。</span><span class="sxs-lookup"><span data-stu-id="fe505-135">The `insertWorksheetsFromBase64` method is supported for Excel on Windows, Mac, and the web.</span></span> <span data-ttu-id="fe505-136">iOS 不支持它。</span><span class="sxs-lookup"><span data-stu-id="fe505-136">It's not supported for iOS.</span></span> <span data-ttu-id="fe505-137">此外，Excel web 版此方法不支持包含数据透视表、图表、注释或 Slicer 元素的源工作表。</span><span class="sxs-lookup"><span data-stu-id="fe505-137">Additionally, in Excel on the web this method doesn't support source worksheets with PivotTable, Chart, Comment, or Slicer elements.</span></span> <span data-ttu-id="fe505-138">如果存在这些对象， `insertWorksheetsFromBase64` 该方法将返回 `UnsupportedFeature` Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="fe505-138">If those objects are present, the `insertWorksheetsFromBase64` method returns the `UnsupportedFeature` error in Excel on the web.</span></span> 

<span data-ttu-id="fe505-139">下面的代码示例演示如何将另一个工作簿中的工作表插入当前工作簿。</span><span class="sxs-lookup"><span data-stu-id="fe505-139">The following code sample shows how to insert worksheets from another workbook into the current workbook.</span></span> <span data-ttu-id="fe505-140">此代码示例首先处理包含对象的工作簿文件并提取 base64 编码的字符串，然后将此 base64 编码的字符串插入 [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) 当前工作簿中。</span><span class="sxs-lookup"><span data-stu-id="fe505-140">This code sample first processes a workbook file with a [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) object and extracts a base64-encoded string, and then it inserts this base64-encoded string into the current workbook.</span></span> <span data-ttu-id="fe505-141">新工作表插入到工作表 **Sheet1** 之后。</span><span class="sxs-lookup"><span data-stu-id="fe505-141">The new worksheets are inserted after the worksheet named **Sheet1**.</span></span> <span data-ttu-id="fe505-142">请注意， `[]` 作为 [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) 属性的参数传递。</span><span class="sxs-lookup"><span data-stu-id="fe505-142">Note that `[]` is passed as the parameter for the [InsertWorksheetOptions.sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) property.</span></span> <span data-ttu-id="fe505-143">这意味着目标工作簿的所有工作表都插入到当前工作簿中。</span><span class="sxs-lookup"><span data-stu-id="fe505-143">This means that all the worksheets from the target workbook are inserted into the current workbook.</span></span>

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);
            
        // Retrieve the current workbook.
        var workbook = context.workbook;
            
        // Set up the insert options. 
        var options = { 
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
            relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
        }; 
            
         // Insert the new worksheets into the current workbook.
         workbook.insertWorksheetsFromBase64(externalWorkbook, options);
         return context.sync();
    });
};

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="fe505-144">保护工作簿的结构</span><span class="sxs-lookup"><span data-stu-id="fe505-144">Protect the workbook's structure</span></span>

<span data-ttu-id="fe505-145">加载项可以控制用户编辑工作簿结构的能力。</span><span class="sxs-lookup"><span data-stu-id="fe505-145">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="fe505-146">Workbook 对象的 `protection` 属性是一个包含 `protect()` 方法的 [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) 对象。</span><span class="sxs-lookup"><span data-stu-id="fe505-146">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="fe505-147">下列示例演示切换对工作簿结构的保护的基本方案。</span><span class="sxs-lookup"><span data-stu-id="fe505-147">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="fe505-148">`protect` 方法接受一个可选字符串参数。</span><span class="sxs-lookup"><span data-stu-id="fe505-148">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="fe505-149">此字符串表示用户要绕过保护并更改工作簿结构所需的密码。</span><span class="sxs-lookup"><span data-stu-id="fe505-149">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="fe505-150">此外，还可以在工作表级别设置保护，来防止不希望发生的数据编辑。</span><span class="sxs-lookup"><span data-stu-id="fe505-150">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="fe505-151">有关详细信息，请参阅[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md#data-protection)一文的“数据保护”部分。</span><span class="sxs-lookup"><span data-stu-id="fe505-151">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="fe505-152">有关 Excel 中工作簿保护的详细信息，请参阅[保护工作簿](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)一文。</span><span class="sxs-lookup"><span data-stu-id="fe505-152">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="fe505-153">访问文档属性</span><span class="sxs-lookup"><span data-stu-id="fe505-153">Access document properties</span></span>

<span data-ttu-id="fe505-154">Workbook 对象可以访问 Office 文件元数据，即[文档属性](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)。</span><span class="sxs-lookup"><span data-stu-id="fe505-154">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="fe505-155">Workbook 对象的 `properties` 属性是一个包含这些元数据值的 [DocumentProperties](/javascript/api/excel/excel.documentproperties) 对象。</span><span class="sxs-lookup"><span data-stu-id="fe505-155">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="fe505-156">以下示例演示如何设置 `author` 属性。</span><span class="sxs-lookup"><span data-stu-id="fe505-156">The following example shows how to set the `author` property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a><span data-ttu-id="fe505-157">自定义属性</span><span class="sxs-lookup"><span data-stu-id="fe505-157">Custom properties</span></span>

<span data-ttu-id="fe505-158">此外，还可以定义自定义属性。</span><span class="sxs-lookup"><span data-stu-id="fe505-158">You can also define custom properties.</span></span> <span data-ttu-id="fe505-159">DocumentProperties 对象保护 `custom` 属性，它表示用户定义的属性的键值对集合。</span><span class="sxs-lookup"><span data-stu-id="fe505-159">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="fe505-160">下列示例演示如何创建名称为“Introduction”且值为“Hello”的自定义属性，以及如何检索它。</span><span class="sxs-lookup"><span data-stu-id="fe505-160">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

#### <a name="worksheet-level-custom-properties"></a><span data-ttu-id="fe505-161">工作表级别的自定义属性</span><span class="sxs-lookup"><span data-stu-id="fe505-161">Worksheet-level custom properties</span></span>

<span data-ttu-id="fe505-162">还可以在工作表级别设置自定义属性。</span><span class="sxs-lookup"><span data-stu-id="fe505-162">Custom properties can also be set at the worksheet level.</span></span> <span data-ttu-id="fe505-163">这些属性类似于文档级别的自定义属性，不同工作表之间可以重复相同的键。</span><span class="sxs-lookup"><span data-stu-id="fe505-163">These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.</span></span> <span data-ttu-id="fe505-164">以下示例演示如何在当前工作表上创建名为 **WorksheetGroup** 的自定义属性，其值为"Alpha"，然后进行检索。</span><span class="sxs-lookup"><span data-stu-id="fe505-164">The following example shows how to create a custom property named **WorksheetGroup** with the value "Alpha" on the current worksheet, then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="fe505-165">访问文档设置</span><span class="sxs-lookup"><span data-stu-id="fe505-165">Access document settings</span></span>

<span data-ttu-id="fe505-166">工作簿的设置类似于自定义属性集合。</span><span class="sxs-lookup"><span data-stu-id="fe505-166">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="fe505-167">区别在于：设置对于单个 Excel 文件和加载项配对而言是唯一的，而属性只是连接到文件。</span><span class="sxs-lookup"><span data-stu-id="fe505-167">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="fe505-168">下列示例演示如何创建并访问设置。</span><span class="sxs-lookup"><span data-stu-id="fe505-168">The following example shows how to create and access a setting.</span></span>

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

## <a name="access-application-culture-settings"></a><span data-ttu-id="fe505-169">访问应用程序区域性设置</span><span class="sxs-lookup"><span data-stu-id="fe505-169">Access application culture settings</span></span>

<span data-ttu-id="fe505-170">工作簿具有影响特定数据的显示方式的语言和区域性设置。</span><span class="sxs-lookup"><span data-stu-id="fe505-170">A workbook has language and culture settings that affect how certain data is displayed.</span></span> <span data-ttu-id="fe505-171">当外接程序的用户跨不同语言和文化共享工作簿时，这些设置可帮助本地化数据。</span><span class="sxs-lookup"><span data-stu-id="fe505-171">These settings can help localize data when your add-in's users are sharing workbooks across different languages and cultures.</span></span> <span data-ttu-id="fe505-172">您的外接程序可以使用字符串分析来基于系统区域性设置本地化数字、日期和时间的格式，以便每个用户都可以查看其区域性格式的数据。</span><span class="sxs-lookup"><span data-stu-id="fe505-172">Your add-in can use string parsing to localize the format of numbers, dates, and times based on the system culture settings so that each user sees data in their own culture's format.</span></span>

<span data-ttu-id="fe505-173">`Application.cultureInfo` 将系统区域性设置定义为 [CultureInfo](/javascript/api/excel/excel.cultureinfo) 对象。</span><span class="sxs-lookup"><span data-stu-id="fe505-173">`Application.cultureInfo` defines the system culture settings as a [CultureInfo](/javascript/api/excel/excel.cultureinfo) object.</span></span> <span data-ttu-id="fe505-174">这包括数字小数分隔符或日期格式等设置。</span><span class="sxs-lookup"><span data-stu-id="fe505-174">This contains settings like the numerical decimal separator or the date format.</span></span>

<span data-ttu-id="fe505-175">某些区域性设置可以通过自定义[UI Excel更改](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e)。</span><span class="sxs-lookup"><span data-stu-id="fe505-175">Some culture settings can be [changed through the Excel UI](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e).</span></span> <span data-ttu-id="fe505-176">系统设置保留在 对象 `CultureInfo` 中。</span><span class="sxs-lookup"><span data-stu-id="fe505-176">The system settings are preserved in the `CultureInfo` object.</span></span> <span data-ttu-id="fe505-177">任何本地更改都保留为 [应用程序](/javascript/api/excel/excel.application)级属性，例如 `Application.decimalSeparator` 。</span><span class="sxs-lookup"><span data-stu-id="fe505-177">Any local changes are kept as [Application](/javascript/api/excel/excel.application)-level properties, such as `Application.decimalSeparator`.</span></span>

<span data-ttu-id="fe505-178">以下示例将数字字符串的十进制分隔符字符从""更改为系统设置所使用的字符。</span><span class="sxs-lookup"><span data-stu-id="fe505-178">The following sample changes the decimal separator character of a numerical string from a ',' to the character used by the system settings.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="fe505-179">向工作簿添加自定义 XML 数据</span><span class="sxs-lookup"><span data-stu-id="fe505-179">Add custom XML data to the workbook</span></span>

<span data-ttu-id="fe505-180">通过 Excel 的 Open XML **.xlsx** 文件格式，可以让加载项将自定义 XML 数据嵌入到工作簿中。</span><span class="sxs-lookup"><span data-stu-id="fe505-180">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="fe505-181">此类数据将一直位于工作簿中，具体取决于加载项。</span><span class="sxs-lookup"><span data-stu-id="fe505-181">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="fe505-182">工作簿包含 [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)，它是一个 [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) 列表。</span><span class="sxs-lookup"><span data-stu-id="fe505-182">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="fe505-183">通过这些部件可以访问 XML 字符串并获得对应的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="fe505-183">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="fe505-184">将这些 ID 存储为设置后，加载项可以维护会话之间的 XML 部件密钥。</span><span class="sxs-lookup"><span data-stu-id="fe505-184">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="fe505-185">以下示例展示了如何使用自定义 XML 部件。</span><span class="sxs-lookup"><span data-stu-id="fe505-185">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="fe505-186">第一个代码块演示了如何将 XML 数据嵌入到文档中。</span><span class="sxs-lookup"><span data-stu-id="fe505-186">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="fe505-187">它将会存储一个审阅者列表，然后使用工作簿的设置保存 XML 的 `id`，以供后续检索。</span><span class="sxs-lookup"><span data-stu-id="fe505-187">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="fe505-188">第二个代码块演示后续如何访问该 XML。</span><span class="sxs-lookup"><span data-stu-id="fe505-188">The second block shows how to access that XML later.</span></span> <span data-ttu-id="fe505-189">“ContosoReviewXmlPartId”设置将被加载和传递到工作簿的 `customXmlParts`。</span><span class="sxs-lookup"><span data-stu-id="fe505-189">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="fe505-190">XML 数据随后将打印至控制台。</span><span class="sxs-lookup"><span data-stu-id="fe505-190">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="fe505-191">仅当顶级自定义 XML 元素包含 `xmlns` 属性时才会填充 `CustomXMLPart.namespaceUri`。</span><span class="sxs-lookup"><span data-stu-id="fe505-191">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="fe505-192">控制计算行为</span><span class="sxs-lookup"><span data-stu-id="fe505-192">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="fe505-193">设置计算模式</span><span class="sxs-lookup"><span data-stu-id="fe505-193">Set calculation mode</span></span>

<span data-ttu-id="fe505-194">默认情况下，当引用的单元格发生更改时，Excel 会重新计算公式结果。</span><span class="sxs-lookup"><span data-stu-id="fe505-194">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="fe505-195">调整此计算行为可以改进加载项的性能。</span><span class="sxs-lookup"><span data-stu-id="fe505-195">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="fe505-196">Application 对象包含一个 `CalculationMode` 类型的 `calculationMode` 属性。</span><span class="sxs-lookup"><span data-stu-id="fe505-196">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="fe505-197">可以将此属性设置为下列值：</span><span class="sxs-lookup"><span data-stu-id="fe505-197">It can be set to the following values:</span></span>

- <span data-ttu-id="fe505-198">`automatic`：默认的重新计算行为，每当相关数据发生更改时 Excel 都会计算新的公式结果。</span><span class="sxs-lookup"><span data-stu-id="fe505-198">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="fe505-199">`automaticExceptTables`：与 `automatic` 相同，但会忽略对表中值的任何更改。</span><span class="sxs-lookup"><span data-stu-id="fe505-199">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="fe505-200">`manual`：仅在用户或加载项请求计算时，才会进行计算。</span><span class="sxs-lookup"><span data-stu-id="fe505-200">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="fe505-201">设置计算类型</span><span class="sxs-lookup"><span data-stu-id="fe505-201">Set calculation type</span></span>

<span data-ttu-id="fe505-202">[Application](/javascript/api/excel/excel.application) 对象提供了一个用于强制立即进行重新计算的方法。</span><span class="sxs-lookup"><span data-stu-id="fe505-202">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="fe505-203">`Application.calculate(calculationType)` 将基于指定的 `calculationType` 启动手动重新计算。</span><span class="sxs-lookup"><span data-stu-id="fe505-203">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="fe505-204">可以指定下列值：</span><span class="sxs-lookup"><span data-stu-id="fe505-204">The following values can be specified:</span></span>

- <span data-ttu-id="fe505-205">`full`：重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。</span><span class="sxs-lookup"><span data-stu-id="fe505-205">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="fe505-206">`fullRebuild`：检查从属的公式，然后重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。</span><span class="sxs-lookup"><span data-stu-id="fe505-206">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="fe505-207">`recalculate`：重新计算所有活动工作簿中自上次计算后发生更改（或已以编程方式将其标记为重新计算目标）的公式，以及从属于它们的公式。</span><span class="sxs-lookup"><span data-stu-id="fe505-207">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="fe505-208">有关重新计算的详细信息，请参阅[更改公式重新计算、迭代或精度](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)一文。</span><span class="sxs-lookup"><span data-stu-id="fe505-208">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="fe505-209">暂停计算</span><span class="sxs-lookup"><span data-stu-id="fe505-209">Temporarily suspend calculations</span></span>

<span data-ttu-id="fe505-210">借助 Excel API，加载项还可以在调用 `RequestContext.sync()` 前禁用计算。</span><span class="sxs-lookup"><span data-stu-id="fe505-210">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="fe505-211">此操作通过 `suspendApiCalculationUntilNextSync()` 完成。</span><span class="sxs-lookup"><span data-stu-id="fe505-211">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="fe505-212">加载项在编辑较大范围且无需访问两次编辑之间的数据时，使用此方法。</span><span class="sxs-lookup"><span data-stu-id="fe505-212">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="save-the-workbook"></a><span data-ttu-id="fe505-213">保存工作簿</span><span class="sxs-lookup"><span data-stu-id="fe505-213">Save the workbook</span></span>

<span data-ttu-id="fe505-214">`Workbook.save` 会将工作簿保存到永久存储区。</span><span class="sxs-lookup"><span data-stu-id="fe505-214">`Workbook.save` saves the workbook to persistent storage.</span></span> <span data-ttu-id="fe505-215">`save` 方法采用单个可选 `saveBehavior` 参数，该参数可为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="fe505-215">The `save` method takes a single, optional `saveBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="fe505-216">`Excel.SaveBehavior.save`（默认）：保存文件，但不提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="fe505-216">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="fe505-217">如果之前未保存文件，则文件保存到默认位置。</span><span class="sxs-lookup"><span data-stu-id="fe505-217">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="fe505-218">如果之前保存过文件，则保存到之前的位置。</span><span class="sxs-lookup"><span data-stu-id="fe505-218">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="fe505-219">`Excel.SaveBehavior.prompt`：如果之前未保存文件，则将提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="fe505-219">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="fe505-220">如果之前已保存文件，则保存到之前的位置且不提示用户。</span><span class="sxs-lookup"><span data-stu-id="fe505-220">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="fe505-221">如果提示用户保存并取消操作，则 `save` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="fe505-221">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a><span data-ttu-id="fe505-222">关闭工作簿</span><span class="sxs-lookup"><span data-stu-id="fe505-222">Close the workbook</span></span>

<span data-ttu-id="fe505-223">`Workbook.close` 会关闭工作簿，一并关闭与该工作簿关联的加载项（Excel 应用程序仍保持打开状态）。</span><span class="sxs-lookup"><span data-stu-id="fe505-223">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="fe505-224">`close` 方法采用单个可选 `closeBehavior` 参数，该参数可为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="fe505-224">The `close` method takes a single, optional `closeBehavior` parameter that can be one of the following values:</span></span>

- <span data-ttu-id="fe505-225">`Excel.CloseBehavior.save`（默认）：在关闭前保存文件。</span><span class="sxs-lookup"><span data-stu-id="fe505-225">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="fe505-226">如果之前未保存文件，则将提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="fe505-226">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="fe505-227">`Excel.CloseBehavior.skipSave`：立即关闭文件但不保存。</span><span class="sxs-lookup"><span data-stu-id="fe505-227">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="fe505-228">所有未保存的更改均将丢失。</span><span class="sxs-lookup"><span data-stu-id="fe505-228">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="fe505-229">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fe505-229">See also</span></span>

- [<span data-ttu-id="fe505-230">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="fe505-230">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="fe505-231">使用 Excel JavaScript API 处理工作表</span><span class="sxs-lookup"><span data-stu-id="fe505-231">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
