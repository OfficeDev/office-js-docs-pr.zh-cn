---
title: 使用 Excel JavaScript API 查找字符串
description: 了解如何使用 Excel JavaScript API 查找范围内的字符串。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b649bb249cd24d7578bc4f8285e5d0a23d0e4cd
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652787"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="dcce4-103">使用 Excel JavaScript API 查找范围内的字符串</span><span class="sxs-lookup"><span data-stu-id="dcce4-103">Find a string within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="dcce4-104">本文提供了一个代码示例，该示例使用 Excel JavaScript API 查找范围内的字符串。</span><span class="sxs-lookup"><span data-stu-id="dcce4-104">This article provides a code sample that finds a string within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="dcce4-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="dcce4-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a><span data-ttu-id="dcce4-106">匹配范围内的字符串</span><span class="sxs-lookup"><span data-stu-id="dcce4-106">Match a string within a range</span></span>

<span data-ttu-id="dcce4-107">`Range` 对象具有 `find` 方法在区域内搜索指定字符串。</span><span class="sxs-lookup"><span data-stu-id="dcce4-107">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="dcce4-108">返回有匹配文本的第一个单元格区域。</span><span class="sxs-lookup"><span data-stu-id="dcce4-108">It returns the range of the first cell with matching text.</span></span>

<span data-ttu-id="dcce4-109">以下代码示例查找值等于字符串 **食品** 的第一个单元格，并将其地址记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="dcce4-109">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="dcce4-110">请注意，若指定的字符串不存在于区域中，`find` 将引发 `ItemNotFound` 错误。</span><span class="sxs-lookup"><span data-stu-id="dcce4-110">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="dcce4-111">若您预计到指定的字符串可能不存在区域中，则可使用 [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) 方法，以便您的代码可正常处理该情况。</span><span class="sxs-lookup"><span data-stu-id="dcce4-111">If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

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

<span data-ttu-id="dcce4-112">在表示一个单元格的区域调用 `find` 方法时，将在整个工作表进行搜索。</span><span class="sxs-lookup"><span data-stu-id="dcce4-112">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="dcce4-113">搜索开始于该单元格，并按照 `SearchCriteria.searchDirection` 指定的方向进行，如有需要在工作表结束的地方换行。</span><span class="sxs-lookup"><span data-stu-id="dcce4-113">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="dcce4-114">另请参阅</span><span class="sxs-lookup"><span data-stu-id="dcce4-114">See also</span></span>

- [<span data-ttu-id="dcce4-115">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="dcce4-115">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="dcce4-116">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="dcce4-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="dcce4-117">使用 Excel JavaScript API 查找区域内的特殊单元格</span><span class="sxs-lookup"><span data-stu-id="dcce4-117">Find special cells within a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-special-cells.md)
