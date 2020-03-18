---
ms.date: 07/15/2019
description: 了解如何在自定义函数中使用不同的参数，例如 Excel 范围、可选参数、调用上下文等。
title: Excel 自定义函数的选项
localization_priority: Normal
ms.openlocfilehash: 1b4097e1190c5d9dc284393d1321c8e2d6c1a8a4
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688667"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="781ca-103">自定义函数参数选项</span><span class="sxs-lookup"><span data-stu-id="781ca-103">Custom functions parameter options</span></span>

<span data-ttu-id="781ca-104">自定义函数可通过多个不同的参数选项进行配置。</span><span class="sxs-lookup"><span data-stu-id="781ca-104">Custom functions are configurable with many different options for parameters.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="781ca-105">可选参数</span><span class="sxs-lookup"><span data-stu-id="781ca-105">Optional parameters</span></span>

<span data-ttu-id="781ca-106">而常规参数是必需的，而可选参数则不是。</span><span class="sxs-lookup"><span data-stu-id="781ca-106">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="781ca-107">当用户在 Excel 中调用函数时，可选参数将显示在括号中。</span><span class="sxs-lookup"><span data-stu-id="781ca-107">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="781ca-108">在下面的示例中，add 函数可以选择添加第三个数字。</span><span class="sxs-lookup"><span data-stu-id="781ca-108">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="781ca-109">在 Excel 中， `=CONTOSO.ADD(first, second, [third])`此函数显示为。</span><span class="sxs-lookup"><span data-stu-id="781ca-109">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="781ca-110">JavaScript</span><span class="sxs-lookup"><span data-stu-id="781ca-110">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

#### <a name="typescript"></a>[<span data-ttu-id="781ca-111">TypeScript</span><span class="sxs-lookup"><span data-stu-id="781ca-111">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

---

> [!NOTE]
> <span data-ttu-id="781ca-112">如果没有为可选参数指定任何值，则 Excel 会为其分配`null`值。</span><span class="sxs-lookup"><span data-stu-id="781ca-112">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="781ca-113">这意味着 TypeScript 中的默认初始化参数不会按预期工作。</span><span class="sxs-lookup"><span data-stu-id="781ca-113">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="781ca-114">因此，请不要使用语法`function add(first:number, second:number, third=0):number` ，因为它不会`third`初始化为0。</span><span class="sxs-lookup"><span data-stu-id="781ca-114">Therefore, don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="781ca-115">而是使用上一示例中所示的 TypeScript 语法。</span><span class="sxs-lookup"><span data-stu-id="781ca-115">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="781ca-116">在定义包含一个或多个可选参数的函数时，应指定可选参数为 null 时所发生的情况。</span><span class="sxs-lookup"><span data-stu-id="781ca-116">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="781ca-117">在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-117">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="781ca-118">如果`zipCode`参数为 null，则默认值设置为`98052`。</span><span class="sxs-lookup"><span data-stu-id="781ca-118">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="781ca-119">如果`dayOfWeek`参数为 null，则将其设置为星期三。</span><span class="sxs-lookup"><span data-stu-id="781ca-119">If the `dayOfWeek` parameter is null, it is set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="781ca-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="781ca-120">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescript"></a>[<span data-ttu-id="781ca-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="781ca-121">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a><span data-ttu-id="781ca-122">范围参数</span><span class="sxs-lookup"><span data-stu-id="781ca-122">Range parameters</span></span>

<span data-ttu-id="781ca-123">您的自定义函数可能接受作为输入参数的单元格数据的范围。</span><span class="sxs-lookup"><span data-stu-id="781ca-123">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="781ca-124">函数还可以返回数据区域。</span><span class="sxs-lookup"><span data-stu-id="781ca-124">A function can also return a range of data.</span></span> <span data-ttu-id="781ca-125">Excel 将一个区域的单元格数据作为二维数组进行传递。</span><span class="sxs-lookup"><span data-stu-id="781ca-125">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="781ca-126">例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。</span><span class="sxs-lookup"><span data-stu-id="781ca-126">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="781ca-127">下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。</span><span class="sxs-lookup"><span data-stu-id="781ca-127">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="781ca-128">请注意，在此函数的 JSON 元数据中，该`type`参数的属性设置`matrix`为。</span><span class="sxs-lookup"><span data-stu-id="781ca-128">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a><span data-ttu-id="781ca-129">重复参数</span><span class="sxs-lookup"><span data-stu-id="781ca-129">Repeating parameters</span></span>

<span data-ttu-id="781ca-130">重复参数允许用户在函数中输入一系列可选的参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-130">A repeating parameter allows a user to enter a series of optional of arguments to a function.</span></span> <span data-ttu-id="781ca-131">调用函数时，将在参数的数组中提供值。</span><span class="sxs-lookup"><span data-stu-id="781ca-131">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="781ca-132">如果参数名称以数字结尾，则每个参数将递增该数，例如`ADD(number1, [number2], [number3],…)`。</span><span class="sxs-lookup"><span data-stu-id="781ca-132">If the parameter name ends with a number, each argument will increment the number, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="781ca-133">这与用于内置 Excel 函数的约定相匹配。</span><span class="sxs-lookup"><span data-stu-id="781ca-133">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="781ca-134">下面的函数汇总了数字、单元格地址和区域的总和（如果已输入）。</span><span class="sxs-lookup"><span data-stu-id="781ca-134">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

<span data-ttu-id="781ca-135">此函数显示`=CONTOSO.ADD([operands], [operands]...)`在 Excel 工作簿中。</span><span class="sxs-lookup"><span data-stu-id="781ca-135">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="781ca-136">重复单个值参数</span><span class="sxs-lookup"><span data-stu-id="781ca-136">Repeating single value parameter</span></span>

<span data-ttu-id="781ca-137">一个重复的单值参数允许传递多个单个值。</span><span class="sxs-lookup"><span data-stu-id="781ca-137">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="781ca-138">例如，用户可以输入 ADD （1，B2，3）。</span><span class="sxs-lookup"><span data-stu-id="781ca-138">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="781ca-139">下面的示例演示如何声明单个值参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-139">The following sample shows how to declare a single value parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a><span data-ttu-id="781ca-140">单个范围参数</span><span class="sxs-lookup"><span data-stu-id="781ca-140">Single range parameter</span></span>

<span data-ttu-id="781ca-141">从技术上讲，单个 range 参数并不是重复参数，但此处包含此参数，这是因为声明与重复参数非常相似。</span><span class="sxs-lookup"><span data-stu-id="781ca-141">A single range parameter is not technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="781ca-142">它向用户显示为 "添加" （A2： B3），其中单个范围是从 Excel 中传递的。</span><span class="sxs-lookup"><span data-stu-id="781ca-142">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="781ca-143">下面的示例展示了如何声明一个 range 参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-143">The following sample shows how to declare a single range parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a><span data-ttu-id="781ca-144">重复区域参数</span><span class="sxs-lookup"><span data-stu-id="781ca-144">Repeating range parameter</span></span>

<span data-ttu-id="781ca-145">重复区域参数允许传递多个区域或数字。</span><span class="sxs-lookup"><span data-stu-id="781ca-145">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="781ca-146">例如，用户可以输入 ADD （5，B2，C3，8，E5： E8）。</span><span class="sxs-lookup"><span data-stu-id="781ca-146">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="781ca-147">重复区域通常是使用类型为三维`number[][][]`矩阵的类型指定的。</span><span class="sxs-lookup"><span data-stu-id="781ca-147">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="781ca-148">有关示例，请参阅为重复参数列出的主示例（#repeating 参数）。</span><span class="sxs-lookup"><span data-stu-id="781ca-148">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="781ca-149">声明重复参数</span><span class="sxs-lookup"><span data-stu-id="781ca-149">Declaring repeating parameters</span></span>
<span data-ttu-id="781ca-150">在 Typescript 中，指示参数是多维的。</span><span class="sxs-lookup"><span data-stu-id="781ca-150">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="781ca-151">例如， `ADD(values: number[])`将指示一维数组， `ADD(values:number[][])`指示二维数组，依此类推。</span><span class="sxs-lookup"><span data-stu-id="781ca-151">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="781ca-152">在 JavaScript 中， `@param values {number[]}`对于二维数组使用一维`@param <name> {number[][]}`数组，对更多维度使用。</span><span class="sxs-lookup"><span data-stu-id="781ca-152">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="781ca-153">对于 "手动创作的 JSON"，请确保在 JSON `"repeating": true`文件中将参数指定为，并检查参数是否标记为`"dimensionality": matrix`。</span><span class="sxs-lookup"><span data-stu-id="781ca-153">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

>[!NOTE]
><span data-ttu-id="781ca-154">包含重复参数的函数将自动包含一个调用参数作为最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-154">Functions containing repeating parameters automatically contain an invocation parameter as the last parameter.</span></span> <span data-ttu-id="781ca-155">有关调用参数的详细信息，请参阅下一节。</span><span class="sxs-lookup"><span data-stu-id="781ca-155">For more information on invocation parameters, see the following section.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="781ca-156">调用参数</span><span class="sxs-lookup"><span data-stu-id="781ca-156">Invocation parameter</span></span>

<span data-ttu-id="781ca-157">每个自定义函数自动传递`invocation`一个参数作为最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-157">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="781ca-158">此参数可用于检索其他上下文，如调用单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="781ca-158">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="781ca-159">也可以用于向 Excel 发送信息，例如用于[取消函数](custom-functions-web-reqs.md#make-a-streaming-function)的函数处理程序。</span><span class="sxs-lookup"><span data-stu-id="781ca-159">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="781ca-160">即使不声明参数，您的自定义函数也有此参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-160">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="781ca-161">在 Excel 中，用户不会看到此参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-161">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="781ca-162">如果要在自定义`invocation`函数中使用，则将其声明为最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="781ca-162">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="781ca-163">在下面的代码示例中， `invocation`将显式声明上下文以供参考。</span><span class="sxs-lookup"><span data-stu-id="781ca-163">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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
```

<span data-ttu-id="781ca-164">参数允许您获取调用单元格的上下文，这在某些方案中非常有用，包括[查找调用自定义函数的单元格的地址](#addressing-cells-context-parameter)。</span><span class="sxs-lookup"><span data-stu-id="781ca-164">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="781ca-165">寻址单元格的上下文参数</span><span class="sxs-lookup"><span data-stu-id="781ca-165">Addressing cell's context parameter</span></span>

<span data-ttu-id="781ca-166">在某些情况下，您需要获取调用自定义函数的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="781ca-166">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="781ca-167">这在以下情况下很有用：</span><span class="sxs-lookup"><span data-stu-id="781ca-167">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="781ca-168">格式区域：将单元格的地址用作存储[OfficeRuntime](../excel/custom-functions-runtime.md#storing-and-accessing-data)中的信息的密钥。</span><span class="sxs-lookup"><span data-stu-id="781ca-168">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](../excel/custom-functions-runtime.md#storing-and-accessing-data).</span></span> <span data-ttu-id="781ca-169">然后，使用 Excel 中的 [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) 从 `OfficeRuntime.storage` 加载该键。</span><span class="sxs-lookup"><span data-stu-id="781ca-169">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="781ca-170">显示缓存值：如果脱机使用函数，将显示 `OfficeRuntime.storage` 中使用 `onCalculated` 存储的缓存值。</span><span class="sxs-lookup"><span data-stu-id="781ca-170">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="781ca-171">协调：使用单元格地址发现原始单元格，以帮助你在处理时进行协调。</span><span class="sxs-lookup"><span data-stu-id="781ca-171">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="781ca-172">若要在函数中请求寻址单元格的上下文，您需要使用函数来查找单元格的地址，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="781ca-172">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="781ca-173">仅当`@requiresAddress`在函数的注释中对单元格地址进行了标记时，才会公开该单元格地址的相关信息。</span><span class="sxs-lookup"><span data-stu-id="781ca-173">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

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
```

<span data-ttu-id="781ca-174">默认情况下，从 `getAddress` 函数返回的值遵循以下格式：`SheetName!CellNumber`。</span><span class="sxs-lookup"><span data-stu-id="781ca-174">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="781ca-175">例如，如果名为“Expense”的工作表中的 B2 单元格调用了函数，则返回的值为 `Expenses!B2`。</span><span class="sxs-lookup"><span data-stu-id="781ca-175">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="781ca-176">后续步骤</span><span class="sxs-lookup"><span data-stu-id="781ca-176">Next steps</span></span>

<span data-ttu-id="781ca-177">了解如何[在自定义函数中保存状态](custom-functions-save-state.md)，或[在自定义函数中使用可变值](custom-functions-volatile.md)。</span><span class="sxs-lookup"><span data-stu-id="781ca-177">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="781ca-178">另请参阅</span><span class="sxs-lookup"><span data-stu-id="781ca-178">See also</span></span>

* [<span data-ttu-id="781ca-179">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="781ca-179">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="781ca-180">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="781ca-180">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="781ca-181">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="781ca-181">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="781ca-182">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="781ca-182">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="781ca-183">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="781ca-183">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)