---
title: 使用 JavaScript API 设置并Excel区域
description: 了解如何使用 javaScript API Excel JavaScript API 设置和获取Excel范围。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0bd4a4f4bcf40e7899ee429cdc631a43ba176077
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075773"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="371df-103">使用 JavaScript API Excel和获取范围</span><span class="sxs-lookup"><span data-stu-id="371df-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="371df-104">本文提供了使用 JavaScript API 设置和获取区域Excel示例。</span><span class="sxs-lookup"><span data-stu-id="371df-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="371df-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="371df-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="371df-106">设置所选区域</span><span class="sxs-lookup"><span data-stu-id="371df-106">Set the selected range</span></span>

<span data-ttu-id="371df-107">下面的代码示例选择活动工作表中的区域 **B2:E6**。</span><span class="sxs-lookup"><span data-stu-id="371df-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="371df-108">选定的区域 B2:E6</span><span class="sxs-lookup"><span data-stu-id="371df-108">Selected range B2:E6</span></span>

![选定区域Excel。](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="371df-110">获取所选区域</span><span class="sxs-lookup"><span data-stu-id="371df-110">Get the selected range</span></span>

<span data-ttu-id="371df-111">下面的代码示例获取所选区域、加载其 `address` 属性，然后向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="371df-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="371df-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="371df-112">See also</span></span>

- [<span data-ttu-id="371df-113">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="371df-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="371df-114">使用 JavaScript API Excel单元格</span><span class="sxs-lookup"><span data-stu-id="371df-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="371df-115">使用 JavaScript API 设置和获取区域Excel文本或公式</span><span class="sxs-lookup"><span data-stu-id="371df-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="371df-116">使用 JavaScript API Excel区域格式</span><span class="sxs-lookup"><span data-stu-id="371df-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
