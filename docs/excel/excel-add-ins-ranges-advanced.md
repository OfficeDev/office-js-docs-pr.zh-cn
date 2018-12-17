---
title: 使用 Excel JavaScript API 对区域执行操作（高级）
description: ''
ms.date: 12/14/2018
ms.openlocfilehash: 42b1127580c46120d337553fdb86a19a78b37567
ms.sourcegitcommit: 09f124fac7b2e711e1a8be562a99624627c0699e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/15/2018
ms.locfileid: "27283791"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a>使用 Excel JavaScript API 对区域执行操作（高级）

本文基于[使用 Excel JavaScript API 对区域执行操作（基本）](excel-add-ins-ranges.md)中包含的信息，它提供了显示如何使用 Excel JavaScript API 对区域执行更多高级任务的代码示例。 有关 **Range** 对象支持的属性和方法的完整列表，请参阅 [Range 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.range)。

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>使用 Moment-MSDate 插件处理日期

[时刻 JavaScript 库](https://momentjs.com/)提供了使用日期和时间戳的便捷方式。 [Moment-MSDate 插件](https://www.npmjs.com/package/moment-msdate)可将时刻格式转换为 Excel 所需的格式。 这是 [NOW 函数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)返回的相同格式。

以下代码显示如何将 **B4** 处的范围设置为时刻的时间戳：

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

这是一项类似于在单元格之外获取日期并将其转换为时刻或其他格式的技术，如以下代码中所示：

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp 
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

你的加载项将必须对范围进行格式化才能以更可读的形式显示日期。 `"[$-409]m/d/yy h:mm AM/PM;@"` 的示例显示类似“12/3/18 3:57 PM”的时间。 有关日期和时间数字格式的详细信息，请参阅[查看自定义数字格式的准则](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)一文中的“日期和时间格式的准则”。

## <a name="copy-and-paste"></a>复制和粘贴

> [!NOTE]
> `Range.copyFrom` 函数当前仅适用于公共预览版（beta 版本）。 若要使用此功能，必须使用 Office.js CDN 的 beta 版库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。
> 如果使用的是 TypeScript 或代码编辑器将 TypeScript 类型定义文件用于 IntelliSense，则使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。

区域的 `copyFrom` 函数将复制 Excel UI 的“复制和粘贴”行为。 调用 `copyFrom` 的区域对象是目标。
将要复制的源作为一个范围或一个表示范围的字符串地址进行传递。 以下代码示例将数据从“A1:E1”**** 复制到“G1”**** 开始的范围（粘贴到“G1:K1”**** 结束）。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom` 具有三个可选参数。

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` 指定将哪些数据从源复制到目标。

- `"Formulas"` 转换源单元格中的公式，并保留这些公式范围的相对位置。 将原样复制任何非公式条目。
- `"Values"` 复制数据值，如果是公式，则复制公式的结果。
- `"Formats"` 复制范围的格式设置（包括字体、颜色和其他格式），但不会复制任何值。
- `"All"`（默认选项）复制数据和格式设置，保留单元格的公式（如果找到）。

`skipBlanks` 设置是否将空白单元格复制到目标。 如果为 true，`copyFrom` 将跳过源范围中的空白单元格。
跳过的单元格不会覆盖目标范围中其对应单元格的现有数据。 默认值为 false。

`transpose` 确定是否将数据转置（即切换其行和列）到源位置。
转置范围沿主对角线翻转，因此，行“1”****、“2”**** 和“3”**** 将成为列“A”****、“B”**** 和“C”****。

以下代码示例和图像在一个简单的方案中演示此行为。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*在上一个函数已运行之前。*

![在区域中运行复制方法之前的 Excel 中的数据](../images/excel-range-copyfrom-skipblanks-before.png)

*在上一个函数已运行之后。*

![在区域中运行复制方法之后的 Excel 中的数据](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a>删除重复项

> [!NOTE]
> 区域对象的 `removeDuplicates` 函数当前仅适用于公共预览版（beta 版本）。 若要使用此功能，必须使用 Office.js CDN 的 beta 版库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。
> 如果使用的是 TypeScript 或代码编辑器将 TypeScript 类型定义文件用于 IntelliSense，则使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。

区域对象的 `removeDuplicates` 函数将删除在指定列中具有重复条目的行。 该函数将从区域最低值索引到最高值索引（从上到下）遍历区域中的每一行。 如果指定列中的值之前显示在区域中，则会删除该行。 在区域内位于已删除行下方的行将上移。 `removeDuplicates` 不影响该区域外的单元格位置。

`removeDuplicates` 使用 `number[]` 来表示已执行重复项检查的列索引。 此数组从零开始并且与区域而不是与工作表相关。 该函数还使用一个布尔参数来指定第一行是否为标题。 如果为 **true**，则在考虑重复项时将忽略顶行。 `removeDuplicates` 函数将返回 `RemoveDuplicatesResult` 对象，用于指定已删除的行数和剩余的唯一行数。

在使用区域的 `removeDuplicates` 函数时，应记住以下几点：

- `removeDuplicates` 会考虑单元格值，而不是函数结果。 如果两个不同的函数具有相同的求值结果，则不会将单元格值视为重复项。
- `removeDuplicates` 不会忽略空单元格。 空单元格的值与任何其他值具有相同的处理方式。 这意味着区域内所含的空行将包含在 `RemoveDuplicatesResult` 中。

以下示例显示删除第一列中具有重复值的条目。

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*在上一个函数已运行之前。*

![在区域中运行删除重复项方法之前的 Excel 中的数据](../images/excel-ranges-remove-duplicates-before.png)

*在上一个函数已运行之后。*

![在区域中运行删除重复项方法之后的 Excel 中的数据](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>另请参阅

- [使用 Excel JavaScript API 对区域执行操作](excel-add-ins-ranges.md)
- [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)