---
title: 使用 Excel JavaScript API 对区域执行操作（基本）
description: 演示如何使用 Excel JavaScript API 对区域执行常见任务的代码示例。
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: 4eb04a58fdf58425f7bb13a6dc457da28625dba5
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294162"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理区域

本文中的代码示例展示了如何使用 Excel JavaScript API 对区域执行常见任务。 有关该对象支持的属性和方法的完整列表 `Range` ，请参阅 [Range 对象 (适用于 Excel 的 JavaScript API) ](/javascript/api/excel/excel.range)。

> [!NOTE]
> 有关如何使用区域执行更高级任务的代码示例，请参阅 [使用 Excel JavaScript API 对区域执行操作（高级）](excel-add-ins-ranges-advanced.md)。

## <a name="get-a-range"></a>获取区域

下面的示例介绍了在工作表中获取对区域的引用的不同方法。

### <a name="get-range-by-address"></a>按地址获取区域

下面的代码示例从名为**sample**的工作表中获取地址为**B2： C5**的范围，加载其 `address` 属性，并向控制台写入一条消息。

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

### <a name="get-range-by-name"></a>按名称获取区域

下面的代码示例 `MyRange` 从名为 **sample**的工作表中获取名为 sample 的区域，加载其 `address` 属性，并向控制台写入一条消息。

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

### <a name="get-used-range"></a>获取使用的区域

下面的代码示例从名为 **sample**的工作表中获取所使用的范围，加载其 `address` 属性，并向控制台写入一条消息。 使用的区域是包含工作表中分配了值或格式的任意单元格的最小区域。 如果整个工作表为空，则该 `getUsedRange()` 方法返回一个区域，该区域仅包含工作表中左上角的单元格。

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

### <a name="get-entire-range"></a>获取整个区域

下面的代码示例从名为 **sample**的工作表中获取整个工作表区域，加载其 `address` 属性，并向控制台写入一条消息。

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

## <a name="insert-a-range-of-cells"></a>插入多个单元格

下面的代码示例将多个单元格插入位置 **B4:E4**，并将其他单元格下移，以便为新的单元格提供空间。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a>插入区域之前的数据

![Excel 中插入区域之前的数据](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>插入区域之后的数据

![Excel 中插入区域之后的数据](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a>清除多个单元格内容

下面的代码示例清除区域 **E2:E5** 中的所有内容和单元格格式设置。  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a>清除区域之前的数据

![Excel 中清除区域之前的数据](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>清除区域之后的数据

![Excel 中清除区域之后的数据](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>删除多个单元格

下面的代码示例删除区域 **B4:E4** 中的单元格，并将其他单元格上移以填充删除的单元格空出的空间。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a>删除区域之前的数据

![Excel 中删除区域之前的数据](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>删除区域之后的数据

![Excel 中删除区域之后的数据](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a>设置所选区域

下面的代码示例选择活动工作表中的区域 **B2:E6**。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a>选定的区域 B2:E6

![Excel 中选定的区域](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>获取所选区域

下面的代码示例获取所选区域，加载其 `address` 属性，并向控制台写入一条消息。

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

## <a name="set-values-or-formulas"></a>设置值或公式

下面的示例演示如何为单个单元格或多个单元格设置值和公式。

### <a name="set-value-for-a-single-cell"></a>设置单个单元格的值

下面的代码示例将单元格 **C3** 的值设置为“5”，然后设置适合数据的最佳列宽。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-value-is-updated"></a>更新单元格值之前的数据

![Excel 中更新单元格值之前的数据](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a>更新单元格值之后的数据

![Excel 中更新单元格值之后的数据](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a>设置多个单元格的值

下面的代码示例为区域 **B5:D5** 中的单元格设置值，然后设置适合数据的最佳列宽。

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

#### <a name="data-before-cell-values-are-updated"></a>更新多个单元格值之前的数据

![Excel 中更新多个单元格值之前的数据](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a>更新多个单元格值之后的数据

![Excel 中更新多个单元格值之后的数据](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a>设置单个单元格的公式

下面的代码示例为单元格 **E3** 设置公式，然后设置适合数据的最佳列宽。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formula-is-set"></a>设置单元格公式之前的数据

![Excel 中设置单元格公式之前的数据](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a>设置单元格公式之后的数据

![Excel 中设置单元格公式之后的数据](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a>设置多个单元格的公式

下面的代码示例为区域 **E2:E6** 中的单元格设置公式，然后设置适合数据的最佳列宽。

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

#### <a name="data-before-cell-formulas-are-set"></a>设置多个单元格公式之前的数据

![Excel 中设置多个单元格公式之前的数据](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a>设置多个单元格公式之后的数据

![Excel 中设置多个单元格公式之后的数据](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a>获取值、文本或公式

以下示例演示如何从多个单元格获取值、文本和公式。

### <a name="get-values-from-a-range-of-cells"></a>从多个单元格获取值

下面的代码示例获取区域 **B2： E6**，加载其 `values` 属性，并将这些值写入控制台。 `values`区域的属性指定单元格包含的原始值。 即使某个区域中的某些单元格包含公式，该 `values` 区域的属性也会指定这些单元格的原始值，而不是任何公式。

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>区域中的数据（E 列中的值为公式的结果）

![Excel 中设置多个单元格公式之后的数据](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a>range.values（通过上面的代码示例记录到控制台）

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

### <a name="get-text-from-a-range-of-cells"></a>从多个单元格获取文本

下面的代码示例获取区域 **B2： E6**，加载其 `text` 属性，并将其写入控制台。 `text`区域的属性指定区域中的单元格的显示值。 即使区域中的某些单元格包含公式，该 `text` 区域的属性也指定这些单元格的显示值，而不是任何公式的显示值。

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>区域中的数据（E 列中的值为公式的结果）

![Excel 中设置多个单元格公式之后的数据](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a>range.text（通过上面的代码示例记录到控制台）

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

### <a name="get-formulas-from-a-range-of-cells"></a>从多个单元格获取公式

下面的代码示例获取区域 **B2： E6**，加载其 `formulas` 属性，并将其写入控制台。 `formulas`Range 的属性指定区域中包含公式的单元格的公式，以及区域中不包含公式的单元格的原始值。

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

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>区域中的数据（E 列中的值为公式的结果）

![Excel 中设置多个单元格公式之后的数据](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a>range.formulas（通过上面的代码示例记录到控制台）

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

## <a name="set-range-format"></a>设置区域格式

下面的示例演示如何为区域中的单元格设置字体颜色、填充颜色和数字格式。

### <a name="set-font-color-and-fill-color"></a>设置字体颜色和填充颜色

下面的代码示例为区域 **B2:E2** 中的单元格设置字体颜色和填充颜色。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>区域中设置字体颜色和填充颜色之前的数据

![Excel 中设置格式之前的数据](../images/excel-ranges-format-before.png)

#### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>区域中设置字体颜色和填充颜色之后的数据

![Excel 中设置格式之后的数据](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a>设置数字格式

下面的代码示例为区域 **D3:E5** 中的单元格设置数字格式。

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

#### <a name="data-in-range-before-number-format-is-set"></a>区域中设置数字格式之前的数据

![Excel 中设置数字格式之前的数据](../images/excel-ranges-format-font-and-fill.png)

#### <a name="data-in-range-after-number-format-is-set"></a>区域中设置数字格式之后的数据

![设置数字格式后 Excel 中的数据](../images/excel-ranges-format-numbers.png)

## <a name="read-or-write-to-an-unbounded-range"></a>读取或写入无限区域

### <a name="read-an-unbounded-range"></a>读取无限区域

无限区域地址是指定整个列或整个行的区域地址。 例如：

- 由整个列组成的区域地址：<ul><li>`C:C`</li><li>`A:F`</li></ul>
- 由整行组成的区域地址：<ul><li>`2:2`</li><li>`1:4`</li></ul>

API 发出请求以检索无限区域时（例如，`getRange('C:C')`），该响应将包含单元格级别属性（如 `null`、`values`、`text` 和 `numberFormat`）的 `formula` 值。 其他区域属性（如 `address` 和 `cellCount`）将包含无限区域的有效值。

### <a name="write-to-an-unbounded-range"></a>写入一个无限区域

您无法在无限区域上设置单元格级别的属性，如 `values` 、 `numberFormat` 和， `formula` 因为输入请求过大。 例如，下面的代码段无效，因为它尝试为无限区域指定 `values`。 如果您尝试为无限区域设置单元格级别的属性，API 将返回错误。

```js
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="read-or-write-to-a-large-range"></a>读取或写入较大区域

如果区域中包含大量单元格、值、数字格式和/或公式，它可能无法在该区域运行 API 操作。 API 将始终尽量尝试在区域内运行所请求的操作（即检索或写入指定的数据），但尝试对较大区域执行读取或写入操作可能会因资源利用率过高而导致 API 错误。 为避免此类错误，建议为较大区域的较小子集运行单独的读取或写入操作，而不是尝试在较大区域内运行单个读取或写入操作。

有关系统限制的详细信息，请参阅 [适用于 Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)的 "Excel 外接程序" 部分。

### <a name="conditional-formatting-of-ranges"></a>范围的条件格式

范围可以根据条件将格式应用于个别单元格。 有关此操作的详细信息，请参阅[将条件格式应用于 Excel 范围](excel-add-ins-conditional-formatting.md)。

## <a name="find-a-cell-using-string-matching"></a>使用字符串匹配查找单元格

`Range` 对象具有 `find` 方法在区域内搜索指定字符串。 返回有匹配文本的第一个单元格区域。 以下代码示例查找值等于字符串 **食品** 的第一个单元格，并将其地址记录到控制台。 请注意，若指定的字符串不存在于区域中，`find` 将引发 `ItemNotFound` 错误。 若您预计到指定的字符串可能不存在区域中，则可使用 [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) 方法，以便您的代码可正常处理该情况。

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

在表示一个单元格的区域调用 `find` 方法时，将在整个工作表进行搜索。 搜索开始于该单元格，并按照 `SearchCriteria.searchDirection` 指定的方向进行，如有需要在工作表结束的地方换行。

## <a name="see-also"></a>另请参阅

- [使用 Excel JavaScript API 对区域执行操作（高级）](excel-add-ins-ranges-advanced.md)
- [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)
