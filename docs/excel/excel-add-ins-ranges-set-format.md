---
title: 使用 Excel JavaScript API 设置区域的格式
description: 了解如何使用 Excel JavaScript API 设置区域的格式。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdd78ea69fc38cbefb9d240dbc61554891c73c21
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652794"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a><span data-ttu-id="0d2b2-103">使用 Excel JavaScript API 设置区域格式</span><span class="sxs-lookup"><span data-stu-id="0d2b2-103">Set range format using the Excel JavaScript API</span></span>

<span data-ttu-id="0d2b2-104">本文提供的代码示例使用 Excel JavaScript API 为区域单元格设置字体颜色、填充颜色和数字格式。</span><span class="sxs-lookup"><span data-stu-id="0d2b2-104">This article provides code samples that set font color, fill color, and number format for cells in a range with the Excel JavaScript API.</span></span> <span data-ttu-id="0d2b2-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="0d2b2-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a><span data-ttu-id="0d2b2-106">设置字体颜色和填充颜色</span><span class="sxs-lookup"><span data-stu-id="0d2b2-106">Set font color and fill color</span></span>

<span data-ttu-id="0d2b2-107">下面的代码示例为区域 **B2:E2** 中的单元格设置字体颜色和填充颜色。</span><span class="sxs-lookup"><span data-stu-id="0d2b2-107">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a><span data-ttu-id="0d2b2-108">区域中设置字体颜色和填充颜色之前的数据</span><span class="sxs-lookup"><span data-stu-id="0d2b2-108">Data in range before font color and fill color are set</span></span>

![Excel 中设置格式之前的数据](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a><span data-ttu-id="0d2b2-110">区域中设置字体颜色和填充颜色之后的数据</span><span class="sxs-lookup"><span data-stu-id="0d2b2-110">Data in range after font color and fill color are set</span></span>

![Excel 中设置格式之后的数据](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a><span data-ttu-id="0d2b2-112">设置数字格式</span><span class="sxs-lookup"><span data-stu-id="0d2b2-112">Set number format</span></span>

<span data-ttu-id="0d2b2-113">下面的代码示例为区域 **D3:E5** 中的单元格设置数字格式。</span><span class="sxs-lookup"><span data-stu-id="0d2b2-113">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

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

### <a name="data-in-range-before-number-format-is-set"></a><span data-ttu-id="0d2b2-114">区域中设置数字格式之前的数据</span><span class="sxs-lookup"><span data-stu-id="0d2b2-114">Data in range before number format is set</span></span>

![设置数字格式之前 Excel 中的数据](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a><span data-ttu-id="0d2b2-116">区域中设置数字格式之后的数据</span><span class="sxs-lookup"><span data-stu-id="0d2b2-116">Data in range after number format is set</span></span>

![设置数字格式后 Excel 中的数据](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a><span data-ttu-id="0d2b2-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0d2b2-118">See also</span></span>

- [<span data-ttu-id="0d2b2-119">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="0d2b2-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="0d2b2-120">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="0d2b2-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="0d2b2-121">使用 Excel JavaScript API 设置和获取区域</span><span class="sxs-lookup"><span data-stu-id="0d2b2-121">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="0d2b2-122">使用 Excel JavaScript API 设置和获取区域值、文本或公式</span><span class="sxs-lookup"><span data-stu-id="0d2b2-122">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
