---
ms.date: 09/15/2020
description: '处理和返回自定义函数中类似 #NULL! 自定义函数中。'
title: 处理并返回自定义函数中的错误
localization_priority: Normal
ms.openlocfilehash: 5da68417aa52f1d14340c8c8a46f4943ffd2d223
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819530"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a><span data-ttu-id="7ce9a-104">处理并返回自定义函数中的错误</span><span class="sxs-lookup"><span data-stu-id="7ce9a-104">Handle and return errors from your custom function</span></span>

> [!NOTE]
> <span data-ttu-id="7ce9a-105">本文中所述的功能目前处于预览阶段，可能会发生更改。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-105">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="7ce9a-106">暂不支持在生产环境中使用。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-106">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="7ce9a-107">你将需要加入 [Office 预览体验成员](https://insider.office.com/join) 计划，以试用预览版功能。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-107">You will need to join the [Office Insider](https://insider.office.com/join) program to try the preview features.</span></span>  <span data-ttu-id="7ce9a-108">试用 "预览" 功能的一个不错的方法是使用 Microsoft 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-108">A good way to try out preview features is by using a Microsoft 365 subscription.</span></span> <span data-ttu-id="7ce9a-109">如果你还没有 Microsoft 365 订阅，则可以通过加入 [microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取免费的90天 renewable microsoft 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-109">If you don't already have a Microsoft 365 subscription, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="7ce9a-110">如果自定义函数运行时出现错误，则返回一个错误以通知用户。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-110">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="7ce9a-111">如果您有特定参数要求（如仅正数），请测试参数并在它们不正确时引发错误。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-111">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="7ce9a-112">还可以使用 `try`-`catch` 块来捕获自定义函数运行时发生的任何错误。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-112">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="7ce9a-113">检测和引发错误</span><span class="sxs-lookup"><span data-stu-id="7ce9a-113">Detect and throw an error</span></span>

<span data-ttu-id="7ce9a-114">我们来看一种需要确保邮政编码参数格式正确的自定义函数能够正常工作的情况。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-114">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="7ce9a-115">下面的自定义函数使用正则表达式来检查邮政编码。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-115">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="7ce9a-116">如果是正确的，它将使用另一个函数查找城市，并返回值。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-116">If it is correct, then it will look up the city using another function, and return the value.</span></span> <span data-ttu-id="7ce9a-117">如果不正确，则 `#VALUE!` 向单元格返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-117">If it isn't correct, it returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="7ce9a-118">CustomFunctions.Error 对象</span><span class="sxs-lookup"><span data-stu-id="7ce9a-118">The CustomFunctions.Error object</span></span>

<span data-ttu-id="7ce9a-119">`CustomFunctions.Error` 对象用于将错误返回单元格。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-119">The `CustomFunctions.Error` object is used to return an error back to the cell.</span></span> <span data-ttu-id="7ce9a-120">创建对象时，请使用以下 `ErrorCode` 枚举值之一指定要使用的错误。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-120">When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="7ce9a-121">ErrorCode 枚举值</span><span class="sxs-lookup"><span data-stu-id="7ce9a-121">ErrorCode enum value</span></span>  |<span data-ttu-id="7ce9a-122">Excel 单元格值</span><span class="sxs-lookup"><span data-stu-id="7ce9a-122">Excel cell value</span></span>  |<span data-ttu-id="7ce9a-123">含义</span><span class="sxs-lookup"><span data-stu-id="7ce9a-123">Meaning</span></span>  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="7ce9a-124">公式中使用的一个值为错误类型。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-124">A value used in the formula is the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="7ce9a-125">函数或服务不可用。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-125">The function or service isn't available.</span></span> |
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="7ce9a-126">请注意，JavaScript 允许除以零，因此你需要仔细编写一个错误处理程序来检测这种情况。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-126">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="7ce9a-127">公式中使用的数字有问题</span><span class="sxs-lookup"><span data-stu-id="7ce9a-127">There is a problem with the number used in the formula</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="7ce9a-128">公式中的区域不相交。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-128">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="7ce9a-129">下面的代码示例演示了如何创建并返回无效数字 (`#NUM!`) 错误。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-129">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="7ce9a-130">返回 `#VALUE!` 错误时，还可以添加当用户将鼠标悬停在单元格上方时将会弹出的自定义消息。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-130">When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell.</span></span> <span data-ttu-id="7ce9a-131">下面的示例演示了如何返回自定义错误消息。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-131">The following example shows how to return a custom error message.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="7ce9a-132">使用 try-catch 块</span><span class="sxs-lookup"><span data-stu-id="7ce9a-132">Use try-catch blocks</span></span>

<span data-ttu-id="7ce9a-133">通常情况下，使用 `try` - `catch` 自定义函数中的块捕捉出现的任何潜在错误。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="7ce9a-134">如果不在代码中处理异常，它们将返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="7ce9a-135">默认情况下，对于未处理的异常，Excel 返回 `#VALUE!`。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-135">By default, Excel returns `#VALUE!` for an unhandled exception.</span></span>

<span data-ttu-id="7ce9a-136">在下面的代码示例中，自定义函数对 REST 服务执行 fetch 调用。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="7ce9a-137">此调用有可能会失败，例如，如果 REST 服务返回错误或网络中断，就可能会失败。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="7ce9a-138">如果发生这种情况，自定义函数将返回 `#N/A` 以指示 Web 调用失败。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-138">If this happens, the custom function will return `#N/A` to indicate the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="7ce9a-139">后续步骤</span><span class="sxs-lookup"><span data-stu-id="7ce9a-139">Next steps</span></span>

<span data-ttu-id="7ce9a-140">了解如何[解决自定义函数中的问题](custom-functions-troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="7ce9a-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="7ce9a-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7ce9a-141">See also</span></span>

* [<span data-ttu-id="7ce9a-142">自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="7ce9a-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="7ce9a-143">自定义函数要求</span><span class="sxs-lookup"><span data-stu-id="7ce9a-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="7ce9a-144">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="7ce9a-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
