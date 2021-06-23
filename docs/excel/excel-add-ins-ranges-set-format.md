---
title: 使用 JavaScript API 设置Excel格式
description: 了解如何使用 Excel JavaScript API 设置区域的格式。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a09d3b4d79584e186c0be37d4a30954c4d4d0086
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075724"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a><span data-ttu-id="9f230-103">使用 JavaScript API Excel区域格式</span><span class="sxs-lookup"><span data-stu-id="9f230-103">Set range format using the Excel JavaScript API</span></span>

<span data-ttu-id="9f230-104">本文提供的代码示例使用 JavaScript API 为区域单元格设置字体颜色、填充颜色和数字Excel格式。</span><span class="sxs-lookup"><span data-stu-id="9f230-104">This article provides code samples that set font color, fill color, and number format for cells in a range with the Excel JavaScript API.</span></span> <span data-ttu-id="9f230-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="9f230-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a><span data-ttu-id="9f230-106">设置字体颜色和填充颜色</span><span class="sxs-lookup"><span data-stu-id="9f230-106">Set font color and fill color</span></span>

<span data-ttu-id="9f230-107">下面的代码示例为区域 **B2:E2** 中的单元格设置字体颜色和填充颜色。</span><span class="sxs-lookup"><span data-stu-id="9f230-107">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a><span data-ttu-id="9f230-108">区域中设置字体颜色和填充颜色之前的数据</span><span class="sxs-lookup"><span data-stu-id="9f230-108">Data in range before font color and fill color are set</span></span>

![设置Excel之前的数据。](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a><span data-ttu-id="9f230-110">区域中设置字体颜色和填充颜色之后的数据</span><span class="sxs-lookup"><span data-stu-id="9f230-110">Data in range after font color and fill color are set</span></span>

![设置Excel格式后的数据。](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a><span data-ttu-id="9f230-112">设置数字格式</span><span class="sxs-lookup"><span data-stu-id="9f230-112">Set number format</span></span>

<span data-ttu-id="9f230-113">下面的代码示例为区域 **D3:E5** 中的单元格设置数字格式。</span><span class="sxs-lookup"><span data-stu-id="9f230-113">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

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

### <a name="data-in-range-before-number-format-is-set"></a><span data-ttu-id="9f230-114">区域中设置数字格式之前的数据</span><span class="sxs-lookup"><span data-stu-id="9f230-114">Data in range before number format is set</span></span>

![设置数字Excel之前的数据。](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a><span data-ttu-id="9f230-116">区域中设置数字格式之后的数据</span><span class="sxs-lookup"><span data-stu-id="9f230-116">Data in range after number format is set</span></span>

![设置数字Excel之后的数据。](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a><span data-ttu-id="9f230-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9f230-118">See also</span></span>

- [<span data-ttu-id="9f230-119">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="9f230-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9f230-120">使用 JavaScript API Excel单元格</span><span class="sxs-lookup"><span data-stu-id="9f230-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="9f230-121">使用 JavaScript API Excel和获取范围</span><span class="sxs-lookup"><span data-stu-id="9f230-121">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="9f230-122">使用 JavaScript API 设置和获取区域Excel文本或公式</span><span class="sxs-lookup"><span data-stu-id="9f230-122">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
