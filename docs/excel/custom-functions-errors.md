---
ms.date: 09/21/2020
description: '处理和返回自定义函数中类似 #NULL! 自定义函数中。'
title: 处理并返回自定义函数中的错误
localization_priority: Normal
ms.openlocfilehash: 58c2ab432a4525f660e2d89735fd3add6e76fa7f
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175526"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a><span data-ttu-id="6cc41-104">处理并返回自定义函数中的错误</span><span class="sxs-lookup"><span data-stu-id="6cc41-104">Handle and return errors from your custom function</span></span>

<span data-ttu-id="6cc41-105">如果自定义函数运行时出现错误，则返回一个错误以通知用户。</span><span class="sxs-lookup"><span data-stu-id="6cc41-105">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="6cc41-106">如果您有特定参数要求（如仅正数），请测试参数并在它们不正确时引发错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-106">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="6cc41-107">还可以使用 `try`-`catch` 块来捕获自定义函数运行时发生的任何错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-107">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="6cc41-108">检测和引发错误</span><span class="sxs-lookup"><span data-stu-id="6cc41-108">Detect and throw an error</span></span>

<span data-ttu-id="6cc41-109">我们来看一种需要确保邮政编码参数格式正确的自定义函数能够正常工作的情况。</span><span class="sxs-lookup"><span data-stu-id="6cc41-109">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="6cc41-110">下面的自定义函数使用正则表达式来检查邮政编码。</span><span class="sxs-lookup"><span data-stu-id="6cc41-110">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="6cc41-111">如果邮政编码格式正确，则它将使用另一个函数查找城市并返回值。</span><span class="sxs-lookup"><span data-stu-id="6cc41-111">If the zip code format is correct, then it will look up the city using another function and return the value.</span></span> <span data-ttu-id="6cc41-112">如果格式无效，该函数将 `#VALUE!` 向单元格返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-112">If the format isn't valid, the function returns a `#VALUE!` error to the cell.</span></span>

```typescript
/**
* Gets a city name for the given U.S. zip code.
* @customfunction
* @param {string} zipCode
* @returns The city of the zip code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="6cc41-113">CustomFunctions.Error 对象</span><span class="sxs-lookup"><span data-stu-id="6cc41-113">The CustomFunctions.Error object</span></span>

<span data-ttu-id="6cc41-114">[Customfunctions.js](/javascript/api/custom-functions-runtime/customfunctions.error)对象用于将错误返回回单元格。</span><span class="sxs-lookup"><span data-stu-id="6cc41-114">The [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object is used to return an error back to the cell.</span></span> <span data-ttu-id="6cc41-115">创建对象时，通过选择下列枚举值之一来指定要使用的错误 `ErrorCode` 。</span><span class="sxs-lookup"><span data-stu-id="6cc41-115">When you create the object, specify which error you want to use by choosing one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="6cc41-116">ErrorCode 枚举值</span><span class="sxs-lookup"><span data-stu-id="6cc41-116">ErrorCode enum value</span></span>  |<span data-ttu-id="6cc41-117">Excel 单元格值</span><span class="sxs-lookup"><span data-stu-id="6cc41-117">Excel cell value</span></span>  |<span data-ttu-id="6cc41-118">含义</span><span class="sxs-lookup"><span data-stu-id="6cc41-118">Meaning</span></span>  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="6cc41-119">请注意，JavaScript 允许除以零，因此你需要仔细编写一个错误处理程序来检测这种情况。</span><span class="sxs-lookup"><span data-stu-id="6cc41-119">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidName`    | `#NAME?`  | <span data-ttu-id="6cc41-120">函数名称中有拼写错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-120">There is a typo in the function name.</span></span> <span data-ttu-id="6cc41-121">请注意，此错误被支持为自定义函数输入错误，而不是作为自定义函数输出错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-121">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span> | 
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="6cc41-122">公式中的数字有问题。</span><span class="sxs-lookup"><span data-stu-id="6cc41-122">There is a problem with a number in the formula.</span></span> |
|`invalidReference` | `#REF!` | <span data-ttu-id="6cc41-123">函数引用了无效的单元格。</span><span class="sxs-lookup"><span data-stu-id="6cc41-123">The function refers to an invalid cell.</span></span> <span data-ttu-id="6cc41-124">请注意，此错误被支持为自定义函数输入错误，而不是作为自定义函数输出错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-124">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span>|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="6cc41-125">公式中的值的类型错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-125">A value in the formula is of the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="6cc41-126">函数或服务不可用。</span><span class="sxs-lookup"><span data-stu-id="6cc41-126">The function or service isn't available.</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="6cc41-127">公式中的区域不相交。</span><span class="sxs-lookup"><span data-stu-id="6cc41-127">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="6cc41-128">下面的代码示例演示了如何创建并返回无效数字 (`#NUM!`) 错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-128">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="6cc41-129">`#VALUE!`和 `#N/A` 错误还支持自定义错误消息。</span><span class="sxs-lookup"><span data-stu-id="6cc41-129">The `#VALUE!` and `#N/A` errors also support custom error messages.</span></span> <span data-ttu-id="6cc41-130">自定义错误消息显示在错误指示器菜单中，该菜单通过将鼠标悬停在包含错误的每个单元格上的错误标志上方来访问。</span><span class="sxs-lookup"><span data-stu-id="6cc41-130">Custom error messages are displayed in the error indicator menu, which is accessed by hovering over the error flag on each cell with an error.</span></span> <span data-ttu-id="6cc41-131">下面的示例演示如何返回包含错误的自定义错误消息 `#VALUE!` 。</span><span class="sxs-lookup"><span data-stu-id="6cc41-131">The following example shows how to return a custom error message with the `#VALUE!` error.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="6cc41-132">使用 try-catch 块</span><span class="sxs-lookup"><span data-stu-id="6cc41-132">Use try-catch blocks</span></span>

<span data-ttu-id="6cc41-133">通常情况下，使用 `try` - `catch` 自定义函数中的块捕捉出现的任何潜在错误。</span><span class="sxs-lookup"><span data-stu-id="6cc41-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="6cc41-134">如果不在代码中处理异常，它们将返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="6cc41-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="6cc41-135">默认情况下，Excel 将返回 `#VALUE!` 未处理的错误或异常。</span><span class="sxs-lookup"><span data-stu-id="6cc41-135">By default, Excel returns `#VALUE!` for unhandled errors or exceptions.</span></span>

<span data-ttu-id="6cc41-136">在下面的代码示例中，自定义函数对 REST 服务执行 fetch 调用。</span><span class="sxs-lookup"><span data-stu-id="6cc41-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="6cc41-137">此调用有可能会失败，例如，如果 REST 服务返回错误或网络中断，就可能会失败。</span><span class="sxs-lookup"><span data-stu-id="6cc41-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="6cc41-138">如果发生这种情况，自定义函数将返回 `#N/A` 以指示 web 调用失败。</span><span class="sxs-lookup"><span data-stu-id="6cc41-138">If this happens, the custom function will return `#N/A` to indicate that the web call failed.</span></span>


```typescript
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + commentID;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    })
}
```

## <a name="next-steps"></a><span data-ttu-id="6cc41-139">后续步骤</span><span class="sxs-lookup"><span data-stu-id="6cc41-139">Next steps</span></span>

<span data-ttu-id="6cc41-140">了解如何[解决自定义函数中的问题](custom-functions-troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="6cc41-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6cc41-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6cc41-141">See also</span></span>

* [<span data-ttu-id="6cc41-142">自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="6cc41-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="6cc41-143">自定义函数要求</span><span class="sxs-lookup"><span data-stu-id="6cc41-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="6cc41-144">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="6cc41-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
