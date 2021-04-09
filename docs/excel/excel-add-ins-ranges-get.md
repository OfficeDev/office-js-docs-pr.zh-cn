---
title: 使用 Excel JavaScript API 获取区域
description: 了解如何使用 Excel JavaScript API 检索区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652800"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="33d8f-103">使用 Excel JavaScript API 获取区域</span><span class="sxs-lookup"><span data-stu-id="33d8f-103">Get a range using the Excel JavaScript API</span></span>

<span data-ttu-id="33d8f-104">本文提供的示例显示了使用 Excel JavaScript API 获取工作表内区域的不同方法。</span><span class="sxs-lookup"><span data-stu-id="33d8f-104">This article provides examples that show different ways to get a range within a worksheet using the Excel JavaScript API.</span></span> <span data-ttu-id="33d8f-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="33d8f-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a><span data-ttu-id="33d8f-106">按地址获取区域</span><span class="sxs-lookup"><span data-stu-id="33d8f-106">Get range by address</span></span>

<span data-ttu-id="33d8f-107">下面的代码示例从名为 **Sample** 的工作表获取地址 **为 B2：C5** 的范围，加载其 属性，然后向控制台写入 `address` 一条消息。</span><span class="sxs-lookup"><span data-stu-id="33d8f-107">The following code sample gets the range with address **B2:C5** from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="get-range-by-name"></a><span data-ttu-id="33d8f-108">按名称获取区域</span><span class="sxs-lookup"><span data-stu-id="33d8f-108">Get range by name</span></span>

<span data-ttu-id="33d8f-109">下面的代码示例从名为 Sample 的工作表获取名为 的范围，加载其 属性，然后 `MyRange` 向控制台 `address` 写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="33d8f-109">The following code sample gets the range named `MyRange` from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="get-used-range"></a><span data-ttu-id="33d8f-110">获取使用的区域</span><span class="sxs-lookup"><span data-stu-id="33d8f-110">Get used range</span></span>

<span data-ttu-id="33d8f-111">下面的代码示例从名为 **Sample** 的工作表获取已用区域，加载其 属性，然后 `address` 向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="33d8f-111">The following code sample gets the used range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span> <span data-ttu-id="33d8f-112">使用的区域是包含工作表中分配了值或格式的任意单元格的最小区域。</span><span class="sxs-lookup"><span data-stu-id="33d8f-112">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="33d8f-113">如果整个工作表为空， `getUsedRange()` 该方法将返回仅由左上单元格组成的区域。</span><span class="sxs-lookup"><span data-stu-id="33d8f-113">If the entire worksheet is blank, the `getUsedRange()` method returns a range that consists of only the top-left cell.</span></span>

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

## <a name="get-entire-range"></a><span data-ttu-id="33d8f-114">获取整个区域</span><span class="sxs-lookup"><span data-stu-id="33d8f-114">Get entire range</span></span>

<span data-ttu-id="33d8f-115">下面的代码示例从名为 **Sample** 的工作表获取整个工作表区域，加载其 属性，然后向控制台写入 `address` 一条消息。</span><span class="sxs-lookup"><span data-stu-id="33d8f-115">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="33d8f-116">另请参阅</span><span class="sxs-lookup"><span data-stu-id="33d8f-116">See also</span></span>

- [<span data-ttu-id="33d8f-117">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="33d8f-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="33d8f-118">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="33d8f-118">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="33d8f-119">使用 Excel JavaScript API 插入区域</span><span class="sxs-lookup"><span data-stu-id="33d8f-119">Insert a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-insert.md)
