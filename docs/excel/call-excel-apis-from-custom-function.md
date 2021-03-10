---
title: 从自定义函数调用 Excel JavaScript API
description: 了解可以从自定义函数调用的 Excel JavaScript API。
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613904"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a><span data-ttu-id="f9543-103">从自定义函数调用 Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="f9543-103">Call Excel JavaScript APIs from a custom function</span></span>

<span data-ttu-id="f9543-104">从自定义函数调用 Excel JavaScript API 以获取区域数据，并获取更多计算上下文。</span><span class="sxs-lookup"><span data-stu-id="f9543-104">Call Excel JavaScript APIs from your custom functions to get range data and obtain more context for your calculations.</span></span> <span data-ttu-id="f9543-105">通过自定义函数调用 Excel JavaScript API 在：</span><span class="sxs-lookup"><span data-stu-id="f9543-105">Calling Excel JavaScript APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="f9543-106">自定义函数需要在计算之前从 Excel 获取信息。</span><span class="sxs-lookup"><span data-stu-id="f9543-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="f9543-107">此信息可能包括文档属性、范围格式、自定义 XML 部件、工作簿名称或其他特定于 Excel 的信息。</span><span class="sxs-lookup"><span data-stu-id="f9543-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="f9543-108">自定义函数将在计算后设置返回值的单元格编号格式。</span><span class="sxs-lookup"><span data-stu-id="f9543-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f9543-109">若要从自定义函数调用 Excel JavaScript API，你需要使用共享的 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="f9543-109">To call Excel JavaScript APIs from your custom function, you'll need to use a shared JavaScript runtime.</span></span> <span data-ttu-id="f9543-110">查看 [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md) 以了解更多信息。</span><span class="sxs-lookup"><span data-stu-id="f9543-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="code-sample"></a><span data-ttu-id="f9543-111">代码示例</span><span class="sxs-lookup"><span data-stu-id="f9543-111">Code sample</span></span>

<span data-ttu-id="f9543-112">若要从自定义函数调用 Excel JavaScript API，首先需要上下文。</span><span class="sxs-lookup"><span data-stu-id="f9543-112">To call Excel JavaScript APIs from a custom function, you first need a context.</span></span> <span data-ttu-id="f9543-113">使用 [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) 对象获取上下文。</span><span class="sxs-lookup"><span data-stu-id="f9543-113">Use the [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) object to get a context.</span></span> <span data-ttu-id="f9543-114">然后使用上下文调用工作簿中所需的 API。</span><span class="sxs-lookup"><span data-stu-id="f9543-114">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="f9543-115">下面的代码示例演示如何用于从工作簿 `Excel.RequestContext` 中的单元格获取值。</span><span class="sxs-lookup"><span data-stu-id="f9543-115">The following code sample shows how to use `Excel.RequestContext` to get a value from a cell in the workbook.</span></span> <span data-ttu-id="f9543-116">在此示例中， `address` 参数将传递到 Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) 方法中，并且必须以字符串形式输入。</span><span class="sxs-lookup"><span data-stu-id="f9543-116">In this sample, the `address` parameter is passed into the Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) method and must be entered as a string.</span></span> <span data-ttu-id="f9543-117">例如，输入到 Excel UI 中的自定义函数必须遵循模式，其中要检索值的单元格 `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` 的地址。</span><span class="sxs-lookup"><span data-stu-id="f9543-117">For example, the custom function entered into the Excel UI must follow the pattern `=CONTOSO.GETRANGEVALUE("A1")`, where `"A1"` is the address of the cell from which to retrieve the value.</span></span>

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a><span data-ttu-id="f9543-118">通过自定义函数调用 Excel JavaScript API 的限制</span><span class="sxs-lookup"><span data-stu-id="f9543-118">Limitations of calling Excel JavaScript APIs through a custom function</span></span>

<span data-ttu-id="f9543-119">不要从更改 Excel 环境的自定义函数调用 Excel JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="f9543-119">Don't call Excel JavaScript APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="f9543-120">这意味着自定义函数不应执行下列任何操作：</span><span class="sxs-lookup"><span data-stu-id="f9543-120">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="f9543-121">在电子表格中插入、删除或设置单元格的格式。</span><span class="sxs-lookup"><span data-stu-id="f9543-121">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="f9543-122">更改另一个单元格的值。</span><span class="sxs-lookup"><span data-stu-id="f9543-122">Change another cell's value.</span></span>
- <span data-ttu-id="f9543-123">移动、重命名、删除或向工作簿添加工作表。</span><span class="sxs-lookup"><span data-stu-id="f9543-123">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="f9543-124">更改任何环境选项，如计算模式或屏幕视图。</span><span class="sxs-lookup"><span data-stu-id="f9543-124">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="f9543-125">向工作簿添加名称。</span><span class="sxs-lookup"><span data-stu-id="f9543-125">Add names to a workbook.</span></span>
- <span data-ttu-id="f9543-126">设置属性或执行大多数方法。</span><span class="sxs-lookup"><span data-stu-id="f9543-126">Set properties or execute most methods.</span></span>

<span data-ttu-id="f9543-127">更改 Excel 可能会导致性能不佳、时间不足和无限循环。</span><span class="sxs-lookup"><span data-stu-id="f9543-127">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="f9543-128">自定义函数计算不应在 Excel 重新计算时运行，因为它将导致不可预知的结果。</span><span class="sxs-lookup"><span data-stu-id="f9543-128">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="f9543-129">相反，请从功能区按钮或任务窗格的上下文中对 Excel 进行更改。</span><span class="sxs-lookup"><span data-stu-id="f9543-129">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f9543-130">后续步骤</span><span class="sxs-lookup"><span data-stu-id="f9543-130">Next steps</span></span>

- [<span data-ttu-id="f9543-131">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="f9543-131">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="f9543-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f9543-132">See also</span></span>

- [<span data-ttu-id="f9543-133">在 Excel 自定义函数和任务窗格教程之间共享数据和事件</span><span class="sxs-lookup"><span data-stu-id="f9543-133">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="f9543-134">将 Office 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="f9543-134">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
