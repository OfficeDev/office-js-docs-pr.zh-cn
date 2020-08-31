---
title: Excel JavaScript API 基本编程概念
description: 使用 Excel JavaScript API 生成 Excel 加载项。
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: dde7dc66e0746fc4d9cf91ed3df824fab05c109d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292588"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a>Excel JavaScript API 基本编程概念

本文介绍了如何使用 [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 生成 Excel 2016 或更高版本的加载项。 它引入了一些核心概念，这些概念是使用 API 的基础，并为执行特定任务提供指导，如读取或写入较大区域、更新区域内的所有单元格等等。

> [!IMPORTANT]
> 请参阅[使用特定于应用程序的 API 模型](../develop/application-specific-api-model.md)，以了解 Excel API 的异步性质以及它们如何与工作簿协同工作。  

## <a name="officejs-apis-for-excel"></a>适用于 Excel 的 Office.js API

Excel 加载项通过使用适 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API包括两个 JavaScript 对象模型：

* **Excel JavaScript API**：[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。

* **通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。

你可能会使用 Excel JavaScript API 开发面向 Excel 2016 或更高版本的加载项中的大部分功能，同时还可以使用通用 API 中的对象。 例如：

* [Context](/javascript/api/office/office.context)：`Context` 对象表示加载项的运行时环境，并提供对 API 关键对象的访问权限。 它由工作簿配置详细信息（如 `contentLanguage` 和 `officeTheme`）组成，并提供有关加载项的运行时环境（如 `host` 和 `platform`）的信息。 此外，它还提供了 `requirements.isSetSupported()` 方法，可用于检查运行加载项的 Excel 应用程序是否支持指定的要求集。
* [Document](/javascript/api/office/office.document)：`Document` 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Excel 文件。

下图说明了可能使用 Excel JavaScript API 或公共 API 的情况。

![Excel JS API 和公共 API 之间差异的图像](../images/excel-js-api-common-api.png)

## <a name="object-model"></a>对象模型

若要了解 Excel API，则必须了解工作簿的各个组件之间如何相互关联。

* 一个 **Workbook** 包含一个或多个 **Worksheet**。
* **Worksheet** 可通过 **Range** 对象访问单元格。
* **Range** 代表一组连续的单元格。
* **Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。
* **Worksheet** 包含单个工作表中存在的那些数据对象的集合。
* **Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。

### <a name="ranges"></a>Range

Range 是工作簿中的一组连续单元格。 加载项通常使用 A1 样式表示法（例如，对于 **B** 列和第 **3** 行中单个单元格，即 **B3** 或从 **C** 列至 **F** 列和第 **2** 行至第 **4** 行的单元格，即 **C2:F4**）来定义范围。

Range 具有三个核心属性：`values`、`formulas` 和 `format`。 这些属性获取或设置单元格值、要计算的公式以及单元格的视觉对象格式设置。

#### <a name="range-sample"></a>Range 示例

以下示例显示了如何创建销售记录。 此函数使用 `Range` 对象来设置值、公式和格式。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

此示例将在当前工作表中创建以下数据：

![显示值行、公式列和格式化标题的销售记录。](../images/excel-overview-range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Chart、Table 和其他数据对象

Excel JavaScript API 可以在 Excel 中创建和设置数据结构和可视化效果。 Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。

#### <a name="creating-a-table"></a>创建表

通过使用数据填充范围创建表。 会将格式设置和表控件（如筛选器）自动应用到该范围。

以下示例使用上一个示例中的范围创建了一个表。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

在包含之前数据的工作表上使用此示例代码将创建下表：

![使用之前的销售记录制成的表。](../images/excel-overview-table-sample.png)

#### <a name="creating-a-chart"></a>创建图表

创建图表以直观显示某个范围内的数据。 该 API 支持数十种图表类型，每种都可以根据需要进行自定义。

下面的示例为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方 100 像素处。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

在工作表上使用上一个表运行此示例将创建以下图表：

![一个柱形图，显示上一个销售记录中三个项目的数量。](../images/excel-overview-chart-sample.png)

## <a name="run-options"></a>运行选项

`Excel.run` 包含需要使用 [RunOptions](/javascript/api/excel/excel.runoptions) 对象的重载。 这包含一组影响函数运行时平台行为的属性。 目前，支持以下属性：

* `delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **true**，批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **false**，批处理请求会在用户处于单元格编辑模式时（导致无法访问用户的错误出现）自动失败。 未指定 `delayForCellEdit` 属性的默认行为等同于此属性为 **false**。

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="null-or-blank-property-values"></a>null 或空属性值

`null` 和空字符串在 Excel JavaScript API 中具有特殊含义。 它们用于表示空单元格、无格式或默认值。 本节详细介绍了在获取和设置属性时如何使用 `null` 和空字符串。

### <a name="null-input-in-2-d-array"></a>二维数组中的 null 输入

在 Excel 中，一个区域由一个二维数组表示，其中第一个维度是行，第二个维度是列。 若要仅为某个区域内的特定单元格设置值、数字格式或公式，请指定二维数组中这些单元格的值、数字格式或公式，并为二维数组中的所有其他单元格指定 `null`。

例如，要更新一个区域内某一个单元格的数字格式，并保留该区域内所有其他单元格的现有数字格式，可指定要更新的单元格的新数字格式，并为所有其他单元格指定 `null`。 下面的代码段为该区域内的第四个单元格设置了一个新的数字格式，并保留该区域内前三个单元格的数字格式不变。

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a>属性的 null 输入

`null` 不是单个属性的有效输入。例如，下面的代码片段无效，因为区域的 `values` 属性不能设置为 `null`。

```js
range.values = null;
```

同样，下面的代码片段也无效，因为 `null` 不是 `color` 属性的有效值。

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a>响应中的 null 属性值

如果指定区域内存在不同的值，诸如 `size` 和 `color` 等格式化属性将在响应中包含 `null` 值。 例如，如果你检索某个区域并加载其 `format.font.color` 属性：

* 如果区域中的所有单元格都具有相同的字体颜色，则 `range.format.font.color` 会指定该颜色。
* 如果该区域内存在多种字体颜色，则 `range.format.font.color` 为 `null`。

### <a name="blank-input-for-a-property"></a>属性的空白输入

如果为属性指定空白值（即两个引号之间没有空格 `''`），它会被解释为属性清除或重置指令。例如：

* 如果为区域的 `values` 属性指定空白值，此区域的内容会被清除。
* 如果为 `numberFormat` 属性指定一个空值，则数字格式会重置为 `General`。
* 如果为 `formula` 属性和 `formulaLocale` 属性指定一个空值，则公式值将被清除。

### <a name="blank-property-values-in-the-response"></a>响应中的空属性值

对于读取操作，响应中的空属性值（即两个引号之间没有空格 `''`）指示该单元格不包含任何数据或值。 在下面第一个示例中，区域中的第一个和最后一个单元格不包含任何数据。 在第二个示例中，区域中的前两个单元格不包含公式。

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="requirement-sets"></a>要求集

要求集是指各组已命名的 API 成员。 Office 加载项可以执行运行时检查或使用清单中指定的要求集确定 Office 应用程序是否支持加载项所需的 API。 要确定每个受支持平台上可用的具体要求集，请参阅 [Excel JavaScript API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md)。

### <a name="checking-for-requirement-set-support-at-runtime"></a>在运行时检查要求集支持

以下代码示例显示如何确定运行加载项的 Office 应用程序是否支持指定的 API 要求集。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>在清单中定义要求集支持

可以在加载项清单中使用[要求元素](../reference/manifest/requirements.md)指定加载项要求激活的最小要求集和/或 API 方法。 如果 Office 应用程序或平台不支持清单的 `Requirements` 元素中指定的要求集或 API 方法，该加载项不会在该应用程序或平台中运行，而且不会显示在“**我的加载项**”中显示的加载项列表中。

以下代码示例显示加载项清单中的 `Requirements` 元素，该元素指定应在支持 ExcelApi 要求集版本 1.3 或更高版本的所有 Office 客户端应用程序中加载该加载项。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> 为了让加载项适用于 Office 应用程序的所有平台（如 Excel 网页版、Windows 版 Excel 和 iPad 版 Excel），建议在运行时检查是否有要求支持，而不是在清单中定义要求集支持。

### <a name="requirement-sets-for-the-officejs-common-api"></a>Office.js 通用 API 的要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。

## <a name="handle-errors"></a>处理错误

当 API 错误出现时，API 返回包含代码和消息的 `error` 对象。 若要详细了解错误处理（包括 API 错误列表），请参阅[错误处理](excel-add-ins-error-handling.md)。

## <a name="see-also"></a>另请参阅

* [生成首个 Excel 加载项](../quickstarts/excel-quickstart-jquery.md)
* [Excel 加载项代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API 性能优化](../excel/performance.md)
* [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
* [常见的编码问题和意外的平台行为](../develop/common-coding-issues.md)
