---
title: 使用 Excel JavaScript API 处理工作簿
description: ''
ms.date: 09/03/2019
localization_priority: Priority
ms.openlocfilehash: eb2313203e770e173d4db12d2bbc03048a08acaa
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715618"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理工作簿

本文提供了代码示例，介绍如何使用 Excel JavaScript API 对工作簿执行常见任务。 有关 **Workbook** 对象支持的属性和方法的完整列表，请参阅 [Workbook 对象 (Excel JavaScript API)](/javascript/api/excel/excel.workbook)。 此外，本文还介绍了通过 [Application](/javascript/api/excel/excel.application) 对象执行的工作簿级别的操作。

Workbook 对象是加载项与 Excel 交互的入口点。 它用于维护工作表、表、数据透视表等的集合，通过这些集合可以访问并更改 Excel 数据。 加载项可以通过 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) 对象访问单个工作表内的所有工作簿数据。 具体来说，加载项可以借助它添加工作表、在工作表间导航并向工作表分配处理程序。 [使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md)一文介绍了如何访问并编辑工作表。

## <a name="get-the-active-cell-or-selected-range"></a>获取活动单元格或选定范围

Workbook 对象包含两种获取用户或加载项所选定单元格范围的方法：`getActiveCell()` 和 `getSelectedRange()`。 `getActiveCell()` 将活动单元格作为 [Range 对象](/javascript/api/excel/excel.range)来从工作簿中获取它。 下列示例演示对 `getActiveCell()` 的调用，紧随其后的是打印到控制台的单元格地址。

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

`getSelectedRange()` 方法返回当前选定的单个范围。 若选定多个范围，将引发 InvalidSelection 错误。 下列示例演示对 `getSelectedRange()` 的调用，并且此方法随后会将相应范围的填充颜色设置为黄色。

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a>创建工作簿

加载项可以新建一个工作簿，并独立于当前运行加载项的 Excel 实例。 Excel 对象包含的 `createWorkbook` 方法可用于实现此目的。 调用此方法时，会立即打开新的工作簿，并在新的 Excel 实例中显示它。 加载项保持打开状态，并随之前的工作簿一起运行。

```js
Excel.createWorkbook();
```

此外，`createWorkbook` 方法还可以创建现有工作簿的副本。 此方法接受 .xlsx 文件的 base64 编码字符串表示形式作为可选参数。 若字符串参数为有效的 .xlsx 文件，则生成的工作簿为该文件的副本。

可以利用[文件切片](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)获取加载项的当前工作簿，作为一个 base64 编码字符串。 可以使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 类将文件转换为所需的 base64 编码字符串，如以下示例所示。

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

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a>将现有工作簿副本插入到当前工作簿中（预览版）

> [!NOTE]
> `WorksheetCollection.addFromBase64` 方法当前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

上一示例显示从现有工作簿创建的新工作簿。 此外，还可以将所有或部分现有工作簿复制到当前与加载项关联的工作簿中。 工作簿的 [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) 可通过 `addFromBase64` 方法将目标工作簿的工作表副本插入到其本身。 其他工作簿文件将作为 base64 编码字符串传递，如 `Excel.createWorkbook` 调用一样。

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

在以下示例中，工作簿的工作表将插入到当前工作簿的活动工作表之后。 请注意，将为 `sheetNamesToInsert?: string[]` 参数传递 `null`。 这意味着将插入所有工作表。

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

## <a name="protect-the-workbooks-structure"></a>保护工作簿的结构

加载项可以控制用户编辑工作簿结构的能力。 Workbook 对象的 `protection` 属性是一个包含 `protect()` 方法的 [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) 对象。 下列示例演示切换对工作簿结构的保护的基本方案。

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

`protect` 方法接受一个可选字符串参数。 此字符串表示用户要绕过保护并更改工作簿结构所需的密码。

此外，还可以在工作表级别设置保护，来防止不希望发生的数据编辑。 有关详细信息，请参阅[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md#data-protection)一文的“数据保护”部分。****

> [!NOTE]
> 有关 Excel 中工作簿保护的详细信息，请参阅[保护工作簿](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517)一文。

## <a name="access-document-properties"></a>访问文档属性

Workbook 对象可以访问 Office 文件元数据，即[文档属性](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75)。 Workbook 对象的 `properties` 属性是一个包含这些元数据值的 [DocumentProperties](/javascript/api/excel/excel.documentproperties) 对象。 下列示例演示如何设置 author 属性。****

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

此外，还可以定义自定义属性。 DocumentProperties 对象保护 `custom` 属性，它表示用户定义的属性的键值对集合。 下列示例演示如何创建名称为“Introduction”且值为“Hello”的自定义属性，以及如何检索它。****

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

## <a name="access-document-settings"></a>访问文档设置

工作簿的设置类似于自定义属性集合。 区别在于：设置对于单个 Excel 文件和加载项配对而言是唯一的，而属性只是连接到文件。 下列示例演示如何创建并访问设置。

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

## <a name="add-custom-xml-data-to-the-workbook"></a>向工作簿添加自定义 XML 数据

通过 Excel 的 Open XML **.xlsx** 文件格式，可以让加载项将自定义 XML 数据嵌入到工作簿中。 此类数据将一直位于工作簿中，具体取决于加载项。

工作簿包含 [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)，它是一个 [CustomXmlParts](/javascript/api/excel/excel.customxmlpart) 列表。 通过这些部件可以访问 XML 字符串并获得对应的唯一 ID。 将这些 ID 存储为设置后，加载项可以维护会话之间的 XML 部件密钥。

以下示例展示了如何使用自定义 XML 部件。 第一个代码块演示了如何将 XML 数据嵌入到文档中。 它将会存储一个审阅者列表，然后使用工作簿的设置保存 XML 的 `id`，以供后续检索。 第二个代码块演示后续如何访问该 XML。 “ContosoReviewXmlPartId”设置将被加载和传递到工作簿的 `customXmlParts`。 XML 数据随后将打印至控制台。

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
> 仅当顶级自定义 XML 元素包含 `xmlns` 属性时才会填充 `CustomXMLPart.namespaceUri`。

## <a name="control-calculation-behavior"></a>控制计算行为

### <a name="set-calculation-mode"></a>设置计算模式

默认情况下，当引用的单元格发生更改时，Excel 会重新计算公式结果。 调整此计算行为可以改进加载项的性能。 Application 对象包含一个 `CalculationMode` 类型的 `calculationMode` 属性。 可以将此属性设置为下列值：

- `automatic`：默认的重新计算行为，每当相关数据发生更改时 Excel 都会计算新的公式结果。
- `automaticExceptTables`：与 `automatic` 相同，但会忽略对表中值的任何更改。
- `manual`：仅在用户或加载项请求计算时，才会进行计算。

### <a name="set-calculation-type"></a>设置计算类型

[Application](/javascript/api/excel/excel.application) 对象提供了一个用于强制立即进行重新计算的方法。 `Application.calculate(calculationType)` 将基于指定的 `calculationType` 启动手动重新计算。 可以指定下列值：

- `full`：重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。
- `fullRebuild`：检查从属的公式，然后重新计算所有打开的工作簿中的所有公式，无论它们自上次重新计算后是否发生了更改。
- `recalculate`：重新计算所有活动工作簿中自上次计算后发生更改（或已以编程方式将其标记为重新计算目标）的公式，以及从属于它们的公式。

> [!NOTE]
> 有关重新计算的详细信息，请参阅[更改公式重新计算、迭代或精度](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4)一文。

### <a name="temporarily-suspend-calculations"></a>暂停计算

借助 Excel API，加载项还可以在调用 `RequestContext.sync()` 前禁用计算。 此操作通过 `suspendApiCalculationUntilNextSync()` 完成。 加载项在编辑较大范围且无需访问两次编辑之间的数据时，使用此方法。

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="comments-preview"></a>批注（预览版）

> [!NOTE]
> 批注 API 当前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

工作簿中的所有[批注](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)都由 `Workbook.comments` 属性进行跟踪。 这包括由用户创建的批注以及由加载项创建的批注。 `Workbook.comments` 属性是一个包含一系列 [Comment](/javascript/api/excel/excel.comment) 对象的 [CommentCollection](/javascript/api/excel/excel.commentcollection) 对象。

若要向工作簿添加批注，请使用 `CommentCollection.add` 方法，将批注的文本作为字符串传入，并将添加批注的单元格作为字符串或 [Range](/javascript/api/excel/excel.range) 对象传入。 下面的代码示例将向单元格 **A2** 添加批注。

```js
Excel.run(function (context) {
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("TODO: add data.", "A2");
    return context.sync();
});
```

每个批注都包含有关其创建情况的元数据，如作者和创建日期。 由加载项创建的批注将被视为是由当前用户创作的。 下面的示例演示如何显示 **A2** 中批注的作者电子邮件、作者姓名和创建日期。

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

每个批注包含零个或多个回复。 `Comment` 对象具有 `replies` 属性，后者是一个包含 [CommentReply](/javascript/api/excel/excel.commentreply) 对象的 [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) 对象。 若要向批注添加回复，请使用 `CommentReplyCollection.add` 方法，传入回复的文本。 回复将按照添加的顺序显示。 下面的代码示例向工作簿中的第一个批注添加回复。

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

若要编辑批注或批注回复，请设置其 `Comment.content` 属性或 `CommentReply.content` 属性。 若要删除批注或批注回复，请使用 `Comment.delete` 方法或 `CommentReply.delete` 方法。 删除批注也会删除与该批注相关联的所有回复。

> [!TIP]
> 还可以使用相同的方法在[工作表](/javascript/api/excel/excel.worksheet)级别管理批注。

## <a name="save-the-workbook-preview"></a>保存工作簿（预览版）

> [!NOTE]
> `Workbook.save` 方法当前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

`Workbook.save` 会将工作簿保存到永久存储区。 `save` 方法采用单个可选 `saveBehavior` 参数，该参数可为以下值之一：

- `Excel.SaveBehavior.save`（默认）：保存文件，但不提示用户指示文件名和保存位置。 如果之前未保存文件，则文件保存到默认位置。 如果之前保存过文件，则保存到之前的位置。
- `Excel.SaveBehavior.prompt`：如果之前未保存文件，则将提示用户指示文件名和保存位置。 如果之前已保存文件，则保存到之前的位置且不提示用户。

> [!CAUTION]
> 如果提示用户保存并取消操作，则 `save` 将引发异常。

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook-preview"></a>关闭工作簿（预览版）

> [!NOTE]
> `Workbook.close` 方法当前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

`Workbook.close` 会关闭工作簿，一并关闭与该工作簿关联的加载项（Excel 应用程序仍保持打开状态）。 `close` 方法采用单个可选 `closeBehavior` 参数，该参数可为以下值之一：

- `Excel.CloseBehavior.save`（默认）：在关闭前保存文件。 如果之前未保存文件，则将提示用户指示文件名和保存位置。
- `Excel.CloseBehavior.skipSave`：立即关闭文件但不保存。 所有未保存的更改均将丢失。

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md)
- [使用 Excel JavaScript API 处理特定范围](excel-add-ins-ranges.md)
