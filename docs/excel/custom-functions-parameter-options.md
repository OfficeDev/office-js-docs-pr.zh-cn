---
ms.date: 12/21/2020
description: 了解如何在自定义函数内使用不同的参数，如 Excel 范围、可选参数、调用上下文等。
title: Excel 自定义函数的选项
localization_priority: Normal
ms.openlocfilehash: 312046551236e96e67de6f63f3e3511aba6f50ce
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735527"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="f12b7-103">自定义函数参数选项</span><span class="sxs-lookup"><span data-stu-id="f12b7-103">Custom functions parameter options</span></span>

<span data-ttu-id="f12b7-104">自定义函数可配置许多不同的参数选项。</span><span class="sxs-lookup"><span data-stu-id="f12b7-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="f12b7-105">可选参数</span><span class="sxs-lookup"><span data-stu-id="f12b7-105">Optional parameters</span></span>

<span data-ttu-id="f12b7-106">当用户在 Excel 中调用函数时，可选参数将显示在括号中。</span><span class="sxs-lookup"><span data-stu-id="f12b7-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="f12b7-107">在下面的示例中，add 函数可以选择添加第三个数字。</span><span class="sxs-lookup"><span data-stu-id="f12b7-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="f12b7-108">此函数在 `=CONTOSO.ADD(first, second, [third])` Excel 中显示。</span><span class="sxs-lookup"><span data-stu-id="f12b7-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="f12b7-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f12b7-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="f12b7-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f12b7-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="f12b7-111">当未指定可选参数的值时，Excel 会为其分配值 `null` 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="f12b7-112">这意味着 TypeScript 中的默认初始化参数将不能正常工作。</span><span class="sxs-lookup"><span data-stu-id="f12b7-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="f12b7-113">请勿使用语法， `function add(first:number, second:number, third=0):number` 因为它不会初始化 `third` 为 0。</span><span class="sxs-lookup"><span data-stu-id="f12b7-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="f12b7-114">请改为使用 TypeScript 语法，如上一示例所示。</span><span class="sxs-lookup"><span data-stu-id="f12b7-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="f12b7-115">定义包含一个或多个可选参数的函数时，请指定可选参数为空时会发生什么情况。</span><span class="sxs-lookup"><span data-stu-id="f12b7-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="f12b7-116">在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。</span><span class="sxs-lookup"><span data-stu-id="f12b7-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="f12b7-117">如果 `zipCode` 参数为空，则默认值设置为 `98052` 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="f12b7-118">如果 `dayOfWeek` 参数为空，则设置为星期三。</span><span class="sxs-lookup"><span data-stu-id="f12b7-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="f12b7-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f12b7-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="f12b7-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f12b7-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="f12b7-121">Range 参数</span><span class="sxs-lookup"><span data-stu-id="f12b7-121">Range parameters</span></span>

<span data-ttu-id="f12b7-122">自定义函数可能会接受单元格数据区域作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="f12b7-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="f12b7-123">函数还可以返回一系列数据。</span><span class="sxs-lookup"><span data-stu-id="f12b7-123">A function can also return a range of data.</span></span> <span data-ttu-id="f12b7-124">Excel 将单元格数据区域作为二维数组传递。</span><span class="sxs-lookup"><span data-stu-id="f12b7-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="f12b7-125">例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。</span><span class="sxs-lookup"><span data-stu-id="f12b7-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="f12b7-126">以下函数接受参数，JSDOC 语法在此函数的 JSON 元数据中设置参数 `values` `number[][]` `dimensionality` `matrix` 的属性。</span><span class="sxs-lookup"><span data-stu-id="f12b7-126">The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function.</span></span> 

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

## <a name="repeating-parameters"></a><span data-ttu-id="f12b7-127">重复参数</span><span class="sxs-lookup"><span data-stu-id="f12b7-127">Repeating parameters</span></span>

<span data-ttu-id="f12b7-128">重复参数允许用户向函数输入一系列可选参数。</span><span class="sxs-lookup"><span data-stu-id="f12b7-128">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="f12b7-129">调用函数时，值在参数的数组中提供。</span><span class="sxs-lookup"><span data-stu-id="f12b7-129">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="f12b7-130">如果参数名称以数字结尾，则每个参数的编号将递增，如 `ADD(number1, [number2], [number3],…)` 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-130">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="f12b7-131">这符合用于内置 Excel 函数的约定。</span><span class="sxs-lookup"><span data-stu-id="f12b7-131">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="f12b7-132">以下函数对数字、单元格地址以及区域（如果输入）总计。</span><span class="sxs-lookup"><span data-stu-id="f12b7-132">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="f12b7-133">此函数 `=CONTOSO.ADD([operands], [operands]...)` 显示在 Excel 工作簿中。</span><span class="sxs-lookup"><span data-stu-id="f12b7-133">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="f12b7-134">重复单个值参数</span><span class="sxs-lookup"><span data-stu-id="f12b7-134">Repeating single value parameter</span></span>

<span data-ttu-id="f12b7-135">重复的单值参数允许传递多个单个值。</span><span class="sxs-lookup"><span data-stu-id="f12b7-135">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="f12b7-136">例如，用户可以输入 ADD (1，B2，3) 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-136">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="f12b7-137">以下示例演示如何声明单个值参数。</span><span class="sxs-lookup"><span data-stu-id="f12b7-137">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="f12b7-138">单个 range 参数</span><span class="sxs-lookup"><span data-stu-id="f12b7-138">Single range parameter</span></span>

<span data-ttu-id="f12b7-139">从技术上说，单个区域参数不是重复参数，但在此处包含，因为声明与重复参数非常相似。</span><span class="sxs-lookup"><span data-stu-id="f12b7-139">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="f12b7-140">对于从 Excel 传递单个区域 (A2：B3) 用户显示为 ADD。</span><span class="sxs-lookup"><span data-stu-id="f12b7-140">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="f12b7-141">以下示例演示如何声明单个 range 参数。</span><span class="sxs-lookup"><span data-stu-id="f12b7-141">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="f12b7-142">重复范围参数</span><span class="sxs-lookup"><span data-stu-id="f12b7-142">Repeating range parameter</span></span>

<span data-ttu-id="f12b7-143">重复范围参数允许传递多个范围或数字。</span><span class="sxs-lookup"><span data-stu-id="f12b7-143">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="f12b7-144">例如，用户可以输入 ADD (5，B2，C3，8，E5：E8) 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-144">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="f12b7-145">重复区域通常使用类型指定 `number[][][]` ，因为它们是三维矩阵。</span><span class="sxs-lookup"><span data-stu-id="f12b7-145">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="f12b7-146">有关示例，请参阅针对重复参数列出的主示例 (#repeating-parameters) 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-146">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="f12b7-147">声明重复参数</span><span class="sxs-lookup"><span data-stu-id="f12b7-147">Declaring repeating parameters</span></span>
<span data-ttu-id="f12b7-148">在 Typescript 中，指示参数是多维的。</span><span class="sxs-lookup"><span data-stu-id="f12b7-148">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="f12b7-149">例如，  `ADD(values: number[])` 表示一维数组， `ADD(values:number[][])` 表示二维数组，等等。</span><span class="sxs-lookup"><span data-stu-id="f12b7-149">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="f12b7-150">在 JavaScript 中，用于一维数组、二维数组等 `@param values {number[]}` `@param <name> {number[][]}` 用于更多维度。</span><span class="sxs-lookup"><span data-stu-id="f12b7-150">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="f12b7-151">对于手动创作的 JSON，请确保参数在 JSON 文件中指定，并检查参数 `"repeating": true` 是否标记为 `"dimensionality": matrix` 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-151">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="f12b7-152">调用参数</span><span class="sxs-lookup"><span data-stu-id="f12b7-152">Invocation parameter</span></span>

<span data-ttu-id="f12b7-153">每个自定义函数都会自动将一个参数作为最后一个输入参数传递，即使该参数 `invocation` 未显式声明。</span><span class="sxs-lookup"><span data-stu-id="f12b7-153">Every custom function is automatically passed an `invocation` argument as the last input parameter, even if it's not explicitly declared.</span></span> <span data-ttu-id="f12b7-154">此 `invocation` 参数对应于 [调用](/javascript/api/custom-functions-runtime/customfunctions.invocation) 对象。</span><span class="sxs-lookup"><span data-stu-id="f12b7-154">This `invocation` parameter corresponds to the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object.</span></span> <span data-ttu-id="f12b7-155">该对象可用于检索其他上下文，例如调用自定义函数的单元格 `Invocation` 的地址。</span><span class="sxs-lookup"><span data-stu-id="f12b7-155">The `Invocation` object can be used to retrieve additional context, such as the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f12b7-156">若要访问 `Invocation` 该对象，必须在自定义函数中声明 `invocation` 为最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="f12b7-156">To access the `Invocation` object, you must declare `invocation` as the last parameter in your custom function.</span></span> 

> [!NOTE]
> <span data-ttu-id="f12b7-157">此 `invocation` 参数不会显示为 Excel 中用户的自定义函数参数。</span><span class="sxs-lookup"><span data-stu-id="f12b7-157">The `invocation` parameter doesn't appear as a custom function argument for users in Excel.</span></span>

<span data-ttu-id="f12b7-158">以下示例演示如何使用参数返回调用自定义函数的单元格 `invocation` 的地址。</span><span class="sxs-lookup"><span data-stu-id="f12b7-158">The following sample shows how to use the `invocation` parameter to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f12b7-159">此示例使用 [对象的地址](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) `Invocation` 属性。</span><span class="sxs-lookup"><span data-stu-id="f12b7-159">This sample uses the [address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) property of the `Invocation` object.</span></span> <span data-ttu-id="f12b7-160">若要访问 `Invocation` 该对象，首先在 `CustomFunctions.Invocation` JSDoc 中声明为参数。</span><span class="sxs-lookup"><span data-stu-id="f12b7-160">To access the `Invocation` object, first declare `CustomFunctions.Invocation` as a parameter in your JSDoc.</span></span> <span data-ttu-id="f12b7-161">接下来， `@requiresAddress` 在 JSDoc 中声明以 `address` 访问对象 `Invocation` 的属性。</span><span class="sxs-lookup"><span data-stu-id="f12b7-161">Next, declare `@requiresAddress` in your JSDoc to access the `address` property of the `Invocation` object.</span></span> <span data-ttu-id="f12b7-162">最后，在函数中检索并返回 `address` 属性。</span><span class="sxs-lookup"><span data-stu-id="f12b7-162">Finally, within the function, retrieve and then return the `address` property.</span></span> 

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<span data-ttu-id="f12b7-163">在 Excel 中，调用对象属性的自定义函数将返回调用该函数的单元格中采用格式 `address` `Invocation` `SheetName!RelativeCellAddress` 的绝对地址。</span><span class="sxs-lookup"><span data-stu-id="f12b7-163">In Excel, a custom function calling the `address` property of the `Invocation` object will return the absolute address following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="f12b7-164">例如，如果输入参数位于单元格 F6 中 **名为"价格** "的工作表上，则返回的参数地址值将为 `Prices!F6` 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-164">For example, if the input parameter is located on a sheet called **Prices** in cell F6, the returned parameter address value will be `Prices!F6`.</span></span> 

<span data-ttu-id="f12b7-165">该 `invocation` 参数还可用于将信息发送到 Excel。</span><span class="sxs-lookup"><span data-stu-id="f12b7-165">The `invocation` parameter can also be used to send information to Excel.</span></span> <span data-ttu-id="f12b7-166">有关详细信息 [，请参阅"创建流](custom-functions-web-reqs.md#make-a-streaming-function) 式处理函数"。</span><span class="sxs-lookup"><span data-stu-id="f12b7-166">See [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function) to learn more.</span></span>

## <a name="detect-the-address-of-a-parameter"></a><span data-ttu-id="f12b7-167">检测参数的地址</span><span class="sxs-lookup"><span data-stu-id="f12b7-167">Detect the address of a parameter</span></span>

<span data-ttu-id="f12b7-168">结合调用 [参数，](#invocation-parameter)可以使用 [调用](/javascript/api/custom-functions-runtime/customfunctions.invocation) 对象检索自定义函数输入参数的地址。</span><span class="sxs-lookup"><span data-stu-id="f12b7-168">In combination with the [invocation parameter](#invocation-parameter), you can use the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object to retrieve the address of a custom function input parameter.</span></span> <span data-ttu-id="f12b7-169">调用时，对象的 [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) 属性允许函数返回所有输入 `Invocation` 参数的地址。</span><span class="sxs-lookup"><span data-stu-id="f12b7-169">When invoked, the [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) property of the `Invocation` object allows a function to return the addresses of all input parameters.</span></span> 

<span data-ttu-id="f12b7-170">这适用于输入数据类型可能有所不同的情况。</span><span class="sxs-lookup"><span data-stu-id="f12b7-170">This is useful in scenarios where input data types may vary.</span></span> <span data-ttu-id="f12b7-171">输入参数的地址可用于检查输入值的数量格式。</span><span class="sxs-lookup"><span data-stu-id="f12b7-171">The address of an input parameter can be used to check the number format of the input value.</span></span> <span data-ttu-id="f12b7-172">然后，如有必要，可以在输入之前调整数字格式。</span><span class="sxs-lookup"><span data-stu-id="f12b7-172">The number format can then be adjusted prior to input, if necessary.</span></span> <span data-ttu-id="f12b7-173">输入参数的地址还可用于检测输入值是否具有与后续计算相关的任何相关属性。</span><span class="sxs-lookup"><span data-stu-id="f12b7-173">The address of an input parameter can also be used to detect whether the input value has any related properties that may be relevant to subsequent calculations.</span></span> 

>[!IMPORTANT]
> <span data-ttu-id="f12b7-174">该属性 `parameterAddresses` 当前仅适用于 [手动创建的 JSON 元数据](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="f12b7-174">The `parameterAddresses` property currently only works with [manually-created JSON metadata](custom-functions-json.md).</span></span> <span data-ttu-id="f12b7-175">若要返回参数地址，对象必须将属性设置为 `options` `requiresParameterAddresses` ，并且 `true` `result` 该对象必须将属性设置为 `dimensionality` `matrix` 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-175">To return parameter addresses, the `options` object must have the `requiresParameterAddresses` property set to `true`, and the `result` object must have the `dimensionality` property set to `matrix`.</span></span>

<span data-ttu-id="f12b7-176">以下自定义函数采用三个输入参数，检索每个参数的对象属性， `parameterAddresses` `Invocation` 然后返回地址。</span><span class="sxs-lookup"><span data-stu-id="f12b7-176">The following custom function takes in three input parameters, retrieves the `parameterAddresses` property of the `Invocation` object for each parameter, and then returns the addresses.</span></span> 

```js
/**
 * Return the address of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<span data-ttu-id="f12b7-177">调用该属性的自定义函数运行时，将按照调用函数的单元格中的格式返回参数 `parameterAddresses` `SheetName!RelativeCellAddress` 地址。</span><span class="sxs-lookup"><span data-stu-id="f12b7-177">When a custom function calling the `parameterAddresses` property runs, the parameter address is returned following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="f12b7-178">例如，如果输入参数位于单元格 D8 中 **名为"成本** "的工作表上，则返回的参数地址值将为 `Costs!D8` 。</span><span class="sxs-lookup"><span data-stu-id="f12b7-178">For example, if the input parameter is located on a sheet called **Costs** in cell D8, the returned parameter address value will be `Costs!D8`.</span></span> <span data-ttu-id="f12b7-179">如果自定义函数具有多个参数并返回多个参数地址，则返回的地址将溢出多个单元格，从调用该函数的单元格垂直排列。</span><span class="sxs-lookup"><span data-stu-id="f12b7-179">If the custom function has multiple parameters and more than one parameter address is returned, the returned addresses will spill across multiple cells, descending vertically from the cell that invoked the function.</span></span> 

## <a name="next-steps"></a><span data-ttu-id="f12b7-180">后续步骤</span><span class="sxs-lookup"><span data-stu-id="f12b7-180">Next steps</span></span>

<span data-ttu-id="f12b7-181">了解如何在自定义 [函数中使用可变值](custom-functions-volatile.md)。</span><span class="sxs-lookup"><span data-stu-id="f12b7-181">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f12b7-182">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f12b7-182">See also</span></span>

* [<span data-ttu-id="f12b7-183">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="f12b7-183">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="f12b7-184">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="f12b7-184">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="f12b7-185">手动为自定义函数创建 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="f12b7-185">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f12b7-186">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="f12b7-186">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f12b7-187">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="f12b7-187">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
