---
title: 使用 Excel JavaScript API 读取或写入无限区域
description: 了解如何使用 Excel JavaScript API 读取或写入无限区域。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652786"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a><span data-ttu-id="2a895-103">使用 Excel JavaScript API 读取或写入无限区域</span><span class="sxs-lookup"><span data-stu-id="2a895-103">Read or write to an unbounded range using the Excel JavaScript API</span></span>

<span data-ttu-id="2a895-104">本文介绍如何使用 Excel JavaScript API 读取和写入无限区域。</span><span class="sxs-lookup"><span data-stu-id="2a895-104">This article describes how to read and write to an unbounded range with the Excel JavaScript API.</span></span> <span data-ttu-id="2a895-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="2a895-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

<span data-ttu-id="2a895-106">无限区域地址是指定整列或整行的范围地址。</span><span class="sxs-lookup"><span data-stu-id="2a895-106">An unbounded range address is a range address that specifies either entire columns or entire rows.</span></span> <span data-ttu-id="2a895-107">例如：</span><span class="sxs-lookup"><span data-stu-id="2a895-107">For example:</span></span>

- <span data-ttu-id="2a895-108">由整列组成的区域地址：</span><span class="sxs-lookup"><span data-stu-id="2a895-108">Range addresses comprised of entire columns:</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="2a895-109">由整行组成的区域地址：</span><span class="sxs-lookup"><span data-stu-id="2a895-109">Range addresses comprised of entire rows:</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a><span data-ttu-id="2a895-110">读取无限区域</span><span class="sxs-lookup"><span data-stu-id="2a895-110">Read an unbounded range</span></span>

<span data-ttu-id="2a895-p103">API 发出请求以检索无限区域时（例如，`getRange('C:C')`），该响应将包含单元格级别属性（如 `null`、`values`、`text` 和 `numberFormat`）的 `formula` 值。 其他区域属性（如 `address` 和 `cellCount`）将包含无限区域的有效值。</span><span class="sxs-lookup"><span data-stu-id="2a895-p103">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

## <a name="write-to-an-unbounded-range"></a><span data-ttu-id="2a895-113">写入一个无限区域</span><span class="sxs-lookup"><span data-stu-id="2a895-113">Write to an unbounded range</span></span>

<span data-ttu-id="2a895-114">由于输入请求过大，无法在无限区域上设置单元格级属性（如 、 和 `values` `numberFormat` `formula` ）。</span><span class="sxs-lookup"><span data-stu-id="2a895-114">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on an unbounded range because the input request is too large.</span></span> <span data-ttu-id="2a895-115">例如，下面的代码示例无效，因为它尝试指定 `values` 无限区域。</span><span class="sxs-lookup"><span data-stu-id="2a895-115">For example, the following code example is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="2a895-116">如果您尝试为无限区域设置单元格级别属性，API 将返回错误。</span><span class="sxs-lookup"><span data-stu-id="2a895-116">The API returns an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a><span data-ttu-id="2a895-117">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2a895-117">See also</span></span>

- [<span data-ttu-id="2a895-118">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="2a895-118">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="2a895-119">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="2a895-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="2a895-120">使用 Excel JavaScript API 读取或写入较大区域</span><span class="sxs-lookup"><span data-stu-id="2a895-120">Read or write to a large range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-large.md)
- [<span data-ttu-id="2a895-121"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="2a895-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
