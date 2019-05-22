---
ms.date: 05/09/2019
description: 了解如何在自定义函数中使用不同的参数, 例如 Excel 范围、可选参数、调用上下文等。
title: Excel 自定义函数的选项
localization_priority: Normal
ms.openlocfilehash: 7bf195bbae696274518966e2a24bd9819e9c3f4b
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337186"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="337e0-103">自定义函数参数选项</span><span class="sxs-lookup"><span data-stu-id="337e0-103">Custom functions parameter options</span></span>

<span data-ttu-id="337e0-104">自定义函数可通过多个不同的参数选项进行配置:</span><span class="sxs-lookup"><span data-stu-id="337e0-104">Custom functions are configurable with many different options for parameters:</span></span>
- [<span data-ttu-id="337e0-105">可选参数</span><span class="sxs-lookup"><span data-stu-id="337e0-105">Optional parameters</span></span>](#custom-functions-optional-parameters)
- [<span data-ttu-id="337e0-106">范围参数</span><span class="sxs-lookup"><span data-stu-id="337e0-106">Range parameters</span></span>](#range-parameters)
- [<span data-ttu-id="337e0-107">调用上下文参数</span><span class="sxs-lookup"><span data-stu-id="337e0-107">Invocation context parameter</span></span>](#invocation-parameter)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-optional-parameters"></a><span data-ttu-id="337e0-108">自定义函数可选参数</span><span class="sxs-lookup"><span data-stu-id="337e0-108">Custom functions optional parameters</span></span>

<span data-ttu-id="337e0-109">而常规参数是必需的, 而可选参数则不是。</span><span class="sxs-lookup"><span data-stu-id="337e0-109">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="337e0-110">当用户在 Excel 中调用函数时，可选参数将显示在括号中。</span><span class="sxs-lookup"><span data-stu-id="337e0-110">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="337e0-111">在下面的示例中, add 函数可以选择添加第三个数字。</span><span class="sxs-lookup"><span data-stu-id="337e0-111">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="337e0-112">在 Excel 中, `=CONTOSO.ADD(first, second, [third])`此函数显示为。</span><span class="sxs-lookup"><span data-stu-id="337e0-112">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third !== undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="337e0-113">定义包含一个或多个可选参数的函数时，应指定未定义可选参数时会发生什么情况。</span><span class="sxs-lookup"><span data-stu-id="337e0-113">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="337e0-114">在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。</span><span class="sxs-lookup"><span data-stu-id="337e0-114">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="337e0-115">如果未`zipCode`定义此参数, 则默认值将设置为`98052`。</span><span class="sxs-lookup"><span data-stu-id="337e0-115">If the `zipCode` parameter is undefined, the default value is set to `98052`.</span></span> <span data-ttu-id="337e0-116">如果未定义 `dayOfWeek` 参数，则会将其设置为星期三。</span><span class="sxs-lookup"><span data-stu-id="337e0-116">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code. If omitted, zipCode = 98052.
 * @param {string} dayOfWeek Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## <a name="range-parameters"></a><span data-ttu-id="337e0-117">范围参数</span><span class="sxs-lookup"><span data-stu-id="337e0-117">Range parameters</span></span>

<span data-ttu-id="337e0-118">您的自定义函数可能接受作为输入参数的单元格数据的范围。</span><span class="sxs-lookup"><span data-stu-id="337e0-118">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="337e0-119">函数还可以返回数据区域。</span><span class="sxs-lookup"><span data-stu-id="337e0-119">A function can also return a range of data.</span></span> <span data-ttu-id="337e0-120">Excel 将一个区域的单元格数据作为二维数组进行传递。</span><span class="sxs-lookup"><span data-stu-id="337e0-120">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="337e0-121">例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。</span><span class="sxs-lookup"><span data-stu-id="337e0-121">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="337e0-122">下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。</span><span class="sxs-lookup"><span data-stu-id="337e0-122">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="337e0-123">请注意, 在此函数的 JSON 元数据中, 该`type`参数的属性设置`matrix`为。</span><span class="sxs-lookup"><span data-stu-id="337e0-123">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## <a name="invocation-parameter"></a><span data-ttu-id="337e0-124">调用参数</span><span class="sxs-lookup"><span data-stu-id="337e0-124">Invocation parameter</span></span>

<span data-ttu-id="337e0-125">每个自定义函数自动传递`invocation`一个参数作为最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="337e0-125">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="337e0-126">此参数可用于检索其他上下文, 如调用单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="337e0-126">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="337e0-127">也可以用于向 Excel 发送信息, 例如用于[取消函数](custom-functions-web-reqs.md#stream-and-cancel-functions)的函数处理程序。</span><span class="sxs-lookup"><span data-stu-id="337e0-127">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#stream-and-cancel-functions).</span></span> <span data-ttu-id="337e0-128">即使不声明参数, 您的自定义函数也有此参数。</span><span class="sxs-lookup"><span data-stu-id="337e0-128">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="337e0-129">在 Excel 中, 用户不会看到此参数。</span><span class="sxs-lookup"><span data-stu-id="337e0-129">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="337e0-130">如果要在自定义`invocation`函数中使用, 则将其声明为最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="337e0-130">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="337e0-131">在下面的代码示例中, `invocation`将显式声明上下文以供参考。</span><span class="sxs-lookup"><span data-stu-id="337e0-131">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="337e0-132">参数允许您获取调用单元格的上下文, 这在某些方案中非常有用, 包括[查找调用自定义函数的单元格的地址](#addressing-cells-context-parameter)。</span><span class="sxs-lookup"><span data-stu-id="337e0-132">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="337e0-133">寻址单元格的上下文参数</span><span class="sxs-lookup"><span data-stu-id="337e0-133">Addressing cell's context parameter</span></span>

<span data-ttu-id="337e0-134">在某些情况下, 您需要获取调用自定义函数的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="337e0-134">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="337e0-135">这在以下情况下很有用:</span><span class="sxs-lookup"><span data-stu-id="337e0-135">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="337e0-136">格式区域: 将单元格的地址用作存储[OfficeRuntime](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data)中的信息的密钥。</span><span class="sxs-lookup"><span data-stu-id="337e0-136">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="337e0-137">然后，使用 Excel 中的 [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) 从 `OfficeRuntime.storage` 加载该键。</span><span class="sxs-lookup"><span data-stu-id="337e0-137">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="337e0-138">显示缓存值：如果脱机使用函数，将显示 `OfficeRuntime.storage` 中使用 `onCalculated` 存储的缓存值。</span><span class="sxs-lookup"><span data-stu-id="337e0-138">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="337e0-139">协调：使用单元格地址发现原始单元格，以帮助你在处理时进行协调。</span><span class="sxs-lookup"><span data-stu-id="337e0-139">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="337e0-140">若要在函数中请求寻址单元格的上下文, 您需要使用函数来查找单元格的地址, 如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="337e0-140">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="337e0-141">仅当`@requiresAddress`在函数的注释中对单元格地址进行了标记时, 才会公开该单元格地址的相关信息。</span><span class="sxs-lookup"><span data-stu-id="337e0-141">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
CustomFunctions.associate("GETADDRESS", getAddress);
```

<span data-ttu-id="337e0-142">默认情况下，从 `getAddress` 函数返回的值遵循以下格式：`SheetName!CellNumber`。</span><span class="sxs-lookup"><span data-stu-id="337e0-142">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="337e0-143">例如，如果名为“Expense”的工作表中的 B2 单元格调用了函数，则返回的值为 `Expenses!B2`。</span><span class="sxs-lookup"><span data-stu-id="337e0-143">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="337e0-144">后续步骤</span><span class="sxs-lookup"><span data-stu-id="337e0-144">Next steps</span></span>
<span data-ttu-id="337e0-145">了解如何[在自定义函数中保存状态](custom-functions-save-state.md), 或[在自定义函数中使用可变值](custom-functions-volatile.md)。</span><span class="sxs-lookup"><span data-stu-id="337e0-145">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="337e0-146">另请参阅</span><span class="sxs-lookup"><span data-stu-id="337e0-146">See also</span></span>

* [<span data-ttu-id="337e0-147">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="337e0-147">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="337e0-148">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="337e0-148">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="337e0-149">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="337e0-149">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="337e0-150">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="337e0-150">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="337e0-151">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="337e0-151">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="337e0-152">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="337e0-152">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
