---
title: 使用 Excel JavaScript API 设置和获取选定区域
description: 了解如何使用 Excel JavaScript API 设置和获取使用 Excel JavaScript API 的范围。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 06b6219924f0667ecef57d608cb417a76ef8031d
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652785"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="811eb-103">使用 Excel JavaScript API 设置和获取区域</span><span class="sxs-lookup"><span data-stu-id="811eb-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="811eb-104">本文提供使用 Excel JavaScript API 设置和获取区域的代码示例。</span><span class="sxs-lookup"><span data-stu-id="811eb-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="811eb-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="811eb-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="811eb-106">设置所选区域</span><span class="sxs-lookup"><span data-stu-id="811eb-106">Set the selected range</span></span>

<span data-ttu-id="811eb-107">下面的代码示例选择活动工作表中的区域 **B2:E6**。</span><span class="sxs-lookup"><span data-stu-id="811eb-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="811eb-108">选定的区域 B2:E6</span><span class="sxs-lookup"><span data-stu-id="811eb-108">Selected range B2:E6</span></span>

![Excel 中选定的区域](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="811eb-110">获取所选区域</span><span class="sxs-lookup"><span data-stu-id="811eb-110">Get the selected range</span></span>

<span data-ttu-id="811eb-111">下面的代码示例获取所选区域、加载其 `address` 属性，然后向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="811eb-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="811eb-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="811eb-112">See also</span></span>

- [<span data-ttu-id="811eb-113">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="811eb-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="811eb-114">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="811eb-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="811eb-115">使用 Excel JavaScript API 设置和获取区域值、文本或公式</span><span class="sxs-lookup"><span data-stu-id="811eb-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="811eb-116">使用 Excel JavaScript API 设置区域格式</span><span class="sxs-lookup"><span data-stu-id="811eb-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
