---
title: 使用 Excel JavaScript API 处理工作簿
description: ''
ms.date: 05/01/2019
localization_priority: Priority
ms.openlocfilehash: e7ec76846a9097ea9e1ef6269661d51c42c21f62
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33620162"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="faa33-102">使用 Excel JavaScript API 处理工作簿</span><span class="sxs-lookup"><span data-stu-id="faa33-102">Work with workbooks using the Excel JavaScript API</span></span>

<span data-ttu-id="faa33-103">本文提供了代码示例，介绍如何使用 Excel JavaScript API 对工作簿执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="faa33-103">This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API.</span></span> <span data-ttu-id="faa33-104">有关 **Workbook** 对象支持的属性和方法的完整列表，请参阅 [Workbook 对象 (Excel JavaScript API)](/javascript/api/excel/excel.workbook)。</span><span class="sxs-lookup"><span data-stu-id="faa33-104">For the complete list of properties and methods that the **Workbook** object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="faa33-105">此外，本文还介绍了通过 [Application](/javascript/api/excel/excel.application) 对象执行的工作簿级别的操作。</span><span class="sxs-lookup"><span data-stu-id="faa33-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="faa33-106">Workbook 对象是加载项与 Excel 交互的入口点。</span><span class="sxs-lookup"><span data-stu-id="faa33-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="faa33-107">它用于维护工作表、表、数据透视表等的集合，通过这些集合可以访问并更改 Excel 数据。</span><span class="sxs-lookup"><span data-stu-id="faa33-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="faa33-108">加载项可以通过 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) 对象访问单个工作表内的所有工作簿数据。</span><span class="sxs-lookup"><span data-stu-id="faa33-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through individual worksheets.</span></span> <span data-ttu-id="faa33-109">具体来说，加载项可以借助它添加工作表、在工作表间导航并向工作表分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="faa33-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="faa33-110">[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md)一文介绍了如何访问并编辑工作表。</span><span class="sxs-lookup"><span data-stu-id="faa33-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="faa33-111">获取活动单元格或选定范围</span><span class="sxs-lookup"><span data-stu-id="faa33-111">Get the active cell or selected range</span></span>

<span data-ttu-id="faa33-112">Workbook 对象包含两种获取用户或加载项所选定单元格范围的方法：`getActiveCell()` 和 `getSelectedRange()`。</span><span class="sxs-lookup"><span data-stu-id="faa33-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="faa33-113">`getActiveCell()` 将活动单元格作为 [Range 对象](/javascript/api/excel/excel.range)来从工作簿中获取它。</span><span class="sxs-lookup"><span data-stu-id="faa33-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="faa33-114">下列示例演示对 `getActiveCell()` 的调用，紧随其后的是打印到控制台的单元格地址。</span><span class="sxs-lookup"><span data-stu-id="faa33-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="faa33-115">`getSelectedRange()` 方法返回当前选定的单个范围。</span><span class="sxs-lookup"><span data-stu-id="faa33-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="faa33-116">若选定多个范围，将引发 InvalidSelection 错误。</span><span class="sxs-lookup"><span data-stu-id="faa33-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="faa33-117">下列示例演示对 `getSelectedRange()` 的调用，并且此方法随后会将相应范围的填充颜色设置为黄色。</span><span class="sxs-lookup"><span data-stu-id="faa33-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="faa33-118">创建工作簿</span><span class="sxs-lookup"><span data-stu-id="faa33-118">Create a workbook</span></span>

<span data-ttu-id="faa33-119">加载项可以新建一个工作簿，并独立于当前运行加载项的 Excel 实例。</span><span class="sxs-lookup"><span data-stu-id="faa33-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="faa33-120">Excel 对象包含的 `createWorkbook` 方法可用于实现此目的。</span><span class="sxs-lookup"><span data-stu-id="faa33-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="faa33-121">调用此方法时，会立即打开新的工作簿，并在新的 Excel 实例中显示它。</span><span class="sxs-lookup"><span data-stu-id="faa33-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="faa33-122">加载项保持打开状态，并随之前的工作簿一起运行。</span><span class="sxs-lookup"><span data-stu-id="faa33-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="faa33-123">此外，`createWorkbook` 方法还可以创建现有工作簿的副本。</span><span class="sxs-lookup"><span data-stu-id="faa33-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="faa33-124">此方法接受 .xlsx 文件的 base64 编码字符串表示形式作为可选参数。</span><span class="sxs-lookup"><span data-stu-id="faa33-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="faa33-125">若字符串参数为有效的 .xlsx 文件，则生成的工作簿为该文件的副本。</span><span class="sxs-lookup"><span data-stu-id="faa33-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="faa33-126">可以利用[文件切片](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)获取加载项的当前工作簿，作为一个 base64 编码字符串。</span><span class="sxs-lookup"><span data-stu-id="faa33-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="faa33-127">可以使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 类将文件转换为所需的 base64 编码字符串，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="faa33-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var workbookContents = event.target.result.substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a><span data-ttu-id="faa33-128">将现有工作簿副本插入到当前工作簿中（预览版）</span><span class="sxs-lookup"><span data-stu-id="faa33-128">Insert a copy of an existing workbook into the current one</span></span>

> [!NOTE]
> <span data-ttu-id="faa33-129">`WorksheetCollection.addFromBase64` 方法当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="faa33-129">The `WorksheetCollection.addFromBase64` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="faa33-130">上一示例显示从现有工作簿创建的新工作簿。</span><span class="sxs-lookup"><span data-stu-id="faa33-130">The previous example shows a new workbook being created from an existing workbook.</span></span> <span data-ttu-id="faa33-131">此外，还可以将所有或部分现有工作簿复制到当前与加载项关联的工作簿中。</span><span class="sxs-lookup"><span data-stu-id="faa33-131">You can also copy some or all of an existing workbook into the one currently associated with your add-in.</span></span> <span data-ttu-id="faa33-132">工作簿的 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) 可通过 `addFromBase64` 方法将目标工作簿的工作表副本插入到其本身。</span><span class="sxs-lookup"><span data-stu-id="faa33-132">A workbook's [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) has the `addFromBase64` method to insert copies of the target workbook's worksheets into itself.</span></span> <span data-ttu-id="faa33-133">其他工作簿文件将作为 base64 编码字符串传递，如 `Excel.createWorkbook` 调用一样。</span><span class="sxs-lookup"><span data-stu-id="faa33-133">The other workbook's file is passed as base64-encoded string, just like the `Excel.createWorkbook` call.</span></span>

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

<span data-ttu-id="faa33-134">在以下示例中，工作簿的工作表将插入到当前工作簿的活动工作表之后。</span><span class="sxs-lookup"><span data-stu-id="faa33-134">The following example shows a workbook's worksheets being inserted in the current workbook, directly after the active worksheet.</span></span> <span data-ttu-id="faa33-135">请注意，将为 `sheetNamesToInsert?: string[]` 参数传递 `null`。</span><span class="sxs-lookup"><span data-stu-id="faa33-135">Note that `null` is passed for the `sheetNamesToInsert?: string[]` parameter.</span></span> <span data-ttu-id="faa33-136">这意味着将插入所有工作表。</span><span class="sxs-lookup"><span data-stu-id="faa33-136">This means all the worksheets are being inserted.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var workbookContents = event.target.result.substr(startIndex + 7);

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

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="faa33-137">保护工作簿的结构</span><span class="sxs-lookup"><span data-stu-id="faa33-137">Protect the workbook's structure</span></span>

<span data-ttu-id="faa33-138">加载项可以控制用户编辑工作簿结构的能力。</span><span class="sxs-lookup"><span data-stu-id="faa33-138">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="faa33-139">Workbook 对象的 `protection` 属性是一个包含 `protect()` 方法的 [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) 对象。</span><span class="sxs-lookup"><span data-stu-id="faa33-139">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="faa33-140">下列示例演示切换对工作簿结构的保护的基本方案。</span><span class="sxs-lookup"><span data-stu-id="faa33-140">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span>

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

<span data-ttu-id="faa33-141">`protect` 方法接受一个可选字符串参数。</span><span class="sxs-lookup"><span data-stu-id="faa33-141">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="faa33-142">此字符串表示用户要绕过保护并更改工作簿结构所需的密码。</span><span class="sxs-lookup"><span data-stu-id="faa33-142">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="faa33-143">此外，还可以在工作表级别设置保护，来防止不希望发生的数据编辑。</span><span class="sxs-lookup"><span data-stu-id="faa33-143">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="faa33-144">有关详细信息，请参阅[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md#data-protection)一文的“数据保护”部分。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="faa33-144">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE]
> <span data-ttu-id="faa33-145">有关 Excel 中工作簿保护的详细信息，请参阅[保护工作簿](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)一文。</span><span class="sxs-lookup"><span data-stu-id="faa33-145">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="faa33-146">访问文档属性</span><span class="sxs-lookup"><span data-stu-id="faa33-146">Access document properties</span></span>

<span data-ttu-id="faa33-147">Workbook 对象可以访问 Office 文件元数据，即[文档属性](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)。</span><span class="sxs-lookup"><span data-stu-id="faa33-147">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="faa33-148">Workbook 对象的 `properties` 属性是一个包含这些元数据值的 [DocumentProperties](/javascript/api/excel/excel.documentproperties) 对象。</span><span class="sxs-lookup"><span data-stu-id="faa33-148">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="faa33-149">下列示例演示如何设置 author 属性。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="faa33-149">The following example shows how to set the **author** property.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="faa33-150">此外，还可以定义自定义属性。</span><span class="sxs-lookup"><span data-stu-id="faa33-150">You can also define custom properties.</span></span> <span data-ttu-id="faa33-151">DocumentProperties 对象保护 `custom` 属性，它表示用户定义的属性的键值对集合。</span><span class="sxs-lookup"><span data-stu-id="faa33-151">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="faa33-152">下列示例演示如何创建名称为“Introduction”且值为“Hello”的自定义属性，以及如何检索它。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="faa33-152">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

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

## <a name="access-document-settings"></a><span data-ttu-id="faa33-153">访问文档设置</span><span class="sxs-lookup"><span data-stu-id="faa33-153">Access document settings</span></span>

<span data-ttu-id="faa33-154">工作簿的设置类似于自定义属性集合。</span><span class="sxs-lookup"><span data-stu-id="faa33-154">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="faa33-155">区别在于：设置对于单个 Excel 文件和加载项配对而言是唯一的，而属性只是连接到文件。</span><span class="sxs-lookup"><span data-stu-id="faa33-155">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="faa33-156">下列示例演示如何创建并访问设置。</span><span class="sxs-lookup"><span data-stu-id="faa33-156">The following example shows how to create and access a setting.</span></span>

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

## <a name="add-custom-xml-data-to-the-workbook"></a><span data-ttu-id="faa33-157">向工作簿添加自定义 XML 数据</span><span class="sxs-lookup"><span data-stu-id="faa33-157">Add custom XML data to the workbook</span></span>

<span data-ttu-id="faa33-158">通过 Excel 的 Open XML **.xlsx** 文件格式，可以让加载项将自定义 XML 数据嵌入到工作簿中。</span><span class="sxs-lookup"><span data-stu-id="faa33-158">Excel's Open XML **.xlsx** file format lets your add-in embed custom XML data in the workbook.</span></span> <span data-ttu-id="faa33-159">此类数据将一直位于工作簿中，具体取决于加载项。</span><span class="sxs-lookup"><span data-stu-id="faa33-159">This data persists with the workbook, independent of the add-in.</span></span>

<span data-ttu-id="faa33-160">工作簿包含 [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)，它是一个 [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) 列表。</span><span class="sxs-lookup"><span data-stu-id="faa33-160">A workbook contains a [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), which is a list of [CustomXmlParts](/javascript/api/excel/excel.customxmlpart).</span></span> <span data-ttu-id="faa33-161">通过这些部件可以访问 XML 字符串并获得对应的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="faa33-161">These give access to the XML strings and a corresponding unique ID.</span></span> <span data-ttu-id="faa33-162">将这些 ID 存储为设置后，加载项可以维护会话之间的 XML 部件密钥。</span><span class="sxs-lookup"><span data-stu-id="faa33-162">By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.</span></span>

<span data-ttu-id="faa33-163">以下示例展示了如何使用自定义 XML 部件。</span><span class="sxs-lookup"><span data-stu-id="faa33-163">The following samples show how to use custom XML parts.</span></span> <span data-ttu-id="faa33-164">第一个代码块演示了如何将 XML 数据嵌入到文档中。</span><span class="sxs-lookup"><span data-stu-id="faa33-164">The first code block demonstrates how to embed XML data in the document.</span></span> <span data-ttu-id="faa33-165">它将会存储一个审阅者列表，然后使用工作簿的设置保存 XML 的 `id`，以供后续检索。</span><span class="sxs-lookup"><span data-stu-id="faa33-165">It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval.</span></span> <span data-ttu-id="faa33-166">第二个代码块演示后续如何访问该 XML。</span><span class="sxs-lookup"><span data-stu-id="faa33-166">The second block shows how to access that XML later.</span></span> <span data-ttu-id="faa33-167">“ContosoReviewXmlPartId”设置将被加载和传递到工作簿的 `customXmlParts`。</span><span class="sxs-lookup"><span data-stu-id="faa33-167">The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`.</span></span> <span data-ttu-id="faa33-168">XML 数据随后将打印至控制台。</span><span class="sxs-lookup"><span data-stu-id="faa33-168">The XML data is then printed to the console.</span></span>

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
> <span data-ttu-id="faa33-169">仅当顶级自定义 XML 元素包含 `xmlns` 属性时才会填充 `CustomXMLPart.namespaceUri`。</span><span class="sxs-lookup"><span data-stu-id="faa33-169">`CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.</span></span>

## <a name="control-calculation-behavior"></a><span data-ttu-id="faa33-170">控制计算行为</span><span class="sxs-lookup"><span data-stu-id="faa33-170">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="faa33-171">设置计算模式</span><span class="sxs-lookup"><span data-stu-id="faa33-171">Set calculation mode</span></span>

<span data-ttu-id="faa33-172">默认情况下，当引用的单元格发生更改时，Excel 会重新计算公式结果。</span><span class="sxs-lookup"><span data-stu-id="faa33-172">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="faa33-173">调整此计算行为可以改进加载项的性能。</span><span class="sxs-lookup"><span data-stu-id="faa33-173">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="faa33-174">Application 对象包含一个 `CalculationMode` 类型的 `calculationMode` 属性。</span><span class="sxs-lookup"><span data-stu-id="faa33-174">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="faa33-175">可以将此属性设置为下列值：</span><span class="sxs-lookup"><span data-stu-id="faa33-175">It can be set to the following values:</span></span>

- <span data-ttu-id="faa33-176">`automatic`：默认的重新计算行为，每当相关数据发生更改时 Excel 都会计算新的公式结果。</span><span class="sxs-lookup"><span data-stu-id="faa33-176">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
- <span data-ttu-id="faa33-177">`automaticExceptTables`：与 `automatic` 相同，但会忽略对表中值的任何更改。</span><span class="sxs-lookup"><span data-stu-id="faa33-177">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
- <span data-ttu-id="faa33-178">`manual`：仅在用户或加载项请求计算时，才会进行计算。</span><span class="sxs-lookup"><span data-stu-id="faa33-178">`manual`: Calculations only occur when the user or add-in requests them.</span></span>

### <a name="set-calculation-type"></a><span data-ttu-id="faa33-179">设置计算类型</span><span class="sxs-lookup"><span data-stu-id="faa33-179">Set calculation type</span></span>

<span data-ttu-id="faa33-180">[Application](/javascript/api/excel/excel.application) 对象提供了一个用于强制立即进行重新计算的方法。</span><span class="sxs-lookup"><span data-stu-id="faa33-180">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="faa33-181">`Application.calculate(calculationType)` 将基于指定的 `calculationType` 启动手动重新计算。</span><span class="sxs-lookup"><span data-stu-id="faa33-181">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="faa33-182">可以指定下列值：</span><span class="sxs-lookup"><span data-stu-id="faa33-182">The following values can be specified:</span></span>

- <span data-ttu-id="faa33-183">`full`：重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。</span><span class="sxs-lookup"><span data-stu-id="faa33-183">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="faa33-184">`fullRebuild`：检查从属的公式，然后重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。</span><span class="sxs-lookup"><span data-stu-id="faa33-184">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
- <span data-ttu-id="faa33-185">`recalculate`：重新计算所有活动工作簿中自上次计算后发生更改（或已以编程方式将其标记为重新计算目标）的公式，以及从属于它们的公式。</span><span class="sxs-lookup"><span data-stu-id="faa33-185">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>

> [!NOTE]
> <span data-ttu-id="faa33-186">有关重新计算的详细信息，请参阅[更改公式重新计算、迭代或精度](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)一文。</span><span class="sxs-lookup"><span data-stu-id="faa33-186">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="faa33-187">暂停计算</span><span class="sxs-lookup"><span data-stu-id="faa33-187">Temporarily suspend calculations</span></span>

<span data-ttu-id="faa33-188">借助 Excel API，加载项还可以在调用 `RequestContext.sync()` 前禁用计算。</span><span class="sxs-lookup"><span data-stu-id="faa33-188">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="faa33-189">此操作通过 `suspendApiCalculationUntilNextSync()` 完成。</span><span class="sxs-lookup"><span data-stu-id="faa33-189">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="faa33-190">加载项在编辑较大范围且无需访问两次编辑之间的数据时，使用此方法。</span><span class="sxs-lookup"><span data-stu-id="faa33-190">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="comments-preview"></a><span data-ttu-id="faa33-191">批注（预览版）</span><span class="sxs-lookup"><span data-stu-id="faa33-191">Comments (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="faa33-192">批注 API 当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="faa33-192">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="faa33-193">工作簿中的所有[批注](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)都由 `Workbook.comments` 属性进行跟踪。</span><span class="sxs-lookup"><span data-stu-id="faa33-193">All [comments](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="faa33-194">这包括由用户创建的批注以及由加载项创建的批注。</span><span class="sxs-lookup"><span data-stu-id="faa33-194">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="faa33-195">`Workbook.comments` 属性是一个包含一系列 [Comment](/javascript/api/excel/excel.comment) 对象的 [CommentCollection](/javascript/api/excel/excel.commentcollection) 对象。</span><span class="sxs-lookup"><span data-stu-id="faa33-195">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span>

<span data-ttu-id="faa33-196">若要向工作簿添加批注，请使用 `CommentCollection.add` 方法，将批注的文本作为字符串传入，并将添加批注的单元格作为字符串或 [Range](/javascript/api/excel/excel.range) 对象传入。</span><span class="sxs-lookup"><span data-stu-id="faa33-196">To add comments to a workbook, use the `CommentCollection.add` method, passing in the comment's text, as a string, and the cell where the comment will be added, as either a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="faa33-197">下面的代码示例将向单元格 **A2** 添加批注。</span><span class="sxs-lookup"><span data-stu-id="faa33-197">The following code sample adds a comment to cell **A2**.</span></span>

```js
Excel.run(function (context) {
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("TODO: add data.", "A2");
    return context.sync();
});
```

<span data-ttu-id="faa33-198">每个批注都包含有关其创建情况的元数据，如作者和创建日期。</span><span class="sxs-lookup"><span data-stu-id="faa33-198">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="faa33-199">由加载项创建的批注将被视为是由当前用户创作的。</span><span class="sxs-lookup"><span data-stu-id="faa33-199">Comments created by your add-in are considered to be authored by the current user.</span></span> <span data-ttu-id="faa33-200">下面的示例演示如何显示 **A2** 中批注的作者电子邮件、作者姓名和创建日期。</span><span class="sxs-lookup"><span data-stu-id="faa33-200">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

```js
Excel.run(function (context) {
    // Get the comment at cell A2.
    var comment = context.workbook.comments.getItemByCell("Comments!A2");
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

<span data-ttu-id="faa33-201">每个批注包含零个或多个回复。</span><span class="sxs-lookup"><span data-stu-id="faa33-201">Each comment contains zero or more replies.</span></span> <span data-ttu-id="faa33-202">`Comment` 对象具有 `replies` 属性，后者是一个包含 [CommentReply](/javascript/api/excel/excel.commentreply) 对象的 [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) 对象。</span><span class="sxs-lookup"><span data-stu-id="faa33-202">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="faa33-203">若要向批注添加回复，请使用 `CommentReplyCollection.add` 方法，传入回复的文本。</span><span class="sxs-lookup"><span data-stu-id="faa33-203">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="faa33-204">回复将按照添加的顺序显示。</span><span class="sxs-lookup"><span data-stu-id="faa33-204">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="faa33-205">下面的代码示例向工作簿中的第一个批注添加回复。</span><span class="sxs-lookup"><span data-stu-id="faa33-205">The following code sample adds a data series to the first chart in the worksheet.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

<span data-ttu-id="faa33-206">若要编辑批注或批注回复，请设置其 `Comment.content` 属性或 `CommentReply.content` 属性。</span><span class="sxs-lookup"><span data-stu-id="faa33-206">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span> <span data-ttu-id="faa33-207">若要删除批注或批注回复，请使用 `Comment.delete` 方法或 `CommentReply.delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="faa33-207">To delete a comment or comment reply, use the `Comment.delete` method or `CommentReply.delete` method.</span></span> <span data-ttu-id="faa33-208">删除批注也会删除与该批注相关联的所有回复。</span><span class="sxs-lookup"><span data-stu-id="faa33-208">Deleting a comment also deletes all the replies associated with that comment.</span></span>

> [!TIP]
> <span data-ttu-id="faa33-209">还可以使用相同的方法在[工作表](/javascript/api/excel/excel.worksheet)级别管理批注。</span><span class="sxs-lookup"><span data-stu-id="faa33-209">Comments can also be managed at the [Worksheet](/javascript/api/excel/excel.worksheet) level using the same techniques.</span></span>

## <a name="save-the-workbook-preview"></a><span data-ttu-id="faa33-210">保存工作簿（预览版）</span><span class="sxs-lookup"><span data-stu-id="faa33-210">Save the workbook</span></span>

> [!NOTE]
> <span data-ttu-id="faa33-211">`Workbook.save` 方法当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="faa33-211">The `Workbook.save` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="faa33-212">`Workbook.save` 会将工作簿保存到永久存储区。</span><span class="sxs-lookup"><span data-stu-id="faa33-212">`Workbook.save` saves the workbook to persistent storage .</span></span> <span data-ttu-id="faa33-213">`save` 方法采用单个可选 `saveBehavior` 参数，该参数可为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="faa33-213">The `save` method takes a single, optional parameter that can be one of the following values:</span></span>

- <span data-ttu-id="faa33-214">`Excel.SaveBehavior.save`（默认）：保存文件，但不提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="faa33-214">`Excel.SaveBehavior.save` (default): The file is saved without prompting the user to specify file name and save location.</span></span> <span data-ttu-id="faa33-215">如果之前未保存文件，则文件保存到默认位置。</span><span class="sxs-lookup"><span data-stu-id="faa33-215">If the file has not been saved previously, it's saved to the default location.</span></span> <span data-ttu-id="faa33-216">如果之前保存过文件，则保存到之前的位置。</span><span class="sxs-lookup"><span data-stu-id="faa33-216">If the file has been saved previously, it's saved to the same location.</span></span>
- <span data-ttu-id="faa33-217">`Excel.SaveBehavior.prompt`：如果之前未保存文件，则将提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="faa33-217">`Excel.SaveBehavior.prompt`: If file has not been saved previously, the user will be prompted to specify file name and save location.</span></span> <span data-ttu-id="faa33-218">如果之前已保存文件，则保存到之前的位置且不提示用户。</span><span class="sxs-lookup"><span data-stu-id="faa33-218">If the file has been saved previously, it will be saved to the same location and the user will not be prompted.</span></span>

> [!CAUTION]
> <span data-ttu-id="faa33-219">如果提示用户保存并取消操作，则 `save` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="faa33-219">If the user is prompted to save and cancels the operation, `save` throws an exception.</span></span>

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook-preview"></a><span data-ttu-id="faa33-220">关闭工作簿（预览版）</span><span class="sxs-lookup"><span data-stu-id="faa33-220">Close the workbook</span></span>

> [!NOTE]
> <span data-ttu-id="faa33-221">`Workbook.close` 方法当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="faa33-221">The `Workbook.close` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="faa33-222">`Workbook.close` 会关闭工作簿，一并关闭与该工作簿关联的加载项（Excel 应用程序仍保持打开状态）。</span><span class="sxs-lookup"><span data-stu-id="faa33-222">`Workbook.close` closes the workbook, along with add-ins that are associated with the workbook (the Excel application remains open).</span></span> <span data-ttu-id="faa33-223">`close` 方法采用单个可选 `closeBehavior` 参数，该参数可为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="faa33-223">The `close` method takes a single, optional parameter that can be one of the following values:</span></span>

- <span data-ttu-id="faa33-224">`Excel.CloseBehavior.save`（默认）：在关闭前保存文件。</span><span class="sxs-lookup"><span data-stu-id="faa33-224">`Excel.CloseBehavior.save` (default): The file is saved before closing.</span></span> <span data-ttu-id="faa33-225">如果之前未保存文件，则将提示用户指示文件名和保存位置。</span><span class="sxs-lookup"><span data-stu-id="faa33-225">If the file has not been saved previously, the user will be prompted to specify file name and save location.</span></span>
- <span data-ttu-id="faa33-226">`Excel.CloseBehavior.skipSave`：立即关闭文件但不保存。</span><span class="sxs-lookup"><span data-stu-id="faa33-226">`Excel.CloseBehavior.skipSave`: The file is immediately closed, without saving.</span></span> <span data-ttu-id="faa33-227">所有未保存的更改均将丢失。</span><span class="sxs-lookup"><span data-stu-id="faa33-227">Any unsaved changes will be lost.</span></span>

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a><span data-ttu-id="faa33-228">另请参阅</span><span class="sxs-lookup"><span data-stu-id="faa33-228">See also</span></span>

- [<span data-ttu-id="faa33-229">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="faa33-229">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="faa33-230">使用 Excel JavaScript API 处理工作表</span><span class="sxs-lookup"><span data-stu-id="faa33-230">Work with worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="faa33-231">使用 Excel JavaScript API 处理特定范围</span><span class="sxs-lookup"><span data-stu-id="faa33-231">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
