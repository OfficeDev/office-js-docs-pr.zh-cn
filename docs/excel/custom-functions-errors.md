---
ms.date: 11/04/2019
description: '处理和返回自定义函数中类似 #NULL! 的错误'
title: 处理和返回自定义函数中的错误（预览）
localization_priority: Priority
ms.openlocfilehash: 5c62b7ccfbc1f0b450e6f36a0fd32f76fe099716
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217069"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a><span data-ttu-id="7681d-104">处理和返回自定义函数中的错误（预览）</span><span class="sxs-lookup"><span data-stu-id="7681d-104">Handle and return errors from your custom function (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="7681d-105">本文中所述的功能目前处于预览阶段，可能会发生更改。</span><span class="sxs-lookup"><span data-stu-id="7681d-105">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="7681d-106">暂不支持在生产环境中使用。</span><span class="sxs-lookup"><span data-stu-id="7681d-106">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="7681d-107">若要试用预览功能，需[加入 Office 预览体验计划](https://insider.office.com/join)。</span><span class="sxs-lookup"><span data-stu-id="7681d-107">You will need to [Office Insider](https://insider.office.com/join) to try the preview features.</span></span>  <span data-ttu-id="7681d-108">试用预览版功能的好方法是使用 Office 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="7681d-108">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="7681d-109">如果你还没有 Office 365 订阅，可以通过加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获得 90 天免费的可续订 Office 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="7681d-109">If you don't already have an Office 365 account, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="7681d-110">如果自定义函数运行时出现错误，你需要返回一个错误以告知用户此情况。</span><span class="sxs-lookup"><span data-stu-id="7681d-110">If something goes wrong while your custom function runs, you will need to return an error to inform the user.</span></span> <span data-ttu-id="7681d-111">如果你有特定参数要求（例如仅限正数），则需要测试参数，如果不正确，需要引发错误。</span><span class="sxs-lookup"><span data-stu-id="7681d-111">If you have specific parameter requirements, such as only positive numbers, you will need to test the parameters and throw an error if they are not correct.</span></span> <span data-ttu-id="7681d-112">还可以使用 `try`-`catch` 块来捕获自定义函数运行时发生的任何错误。</span><span class="sxs-lookup"><span data-stu-id="7681d-112">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="7681d-113">检测和引发错误</span><span class="sxs-lookup"><span data-stu-id="7681d-113">Detect and throw an error</span></span>

<span data-ttu-id="7681d-114">假设你需要确保邮政编码参数的格式正确，使自定义函数能够正常工作。</span><span class="sxs-lookup"><span data-stu-id="7681d-114">Let’s look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="7681d-115">下面的自定义函数使用正则表达式来检查邮政编码。</span><span class="sxs-lookup"><span data-stu-id="7681d-115">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="7681d-116">如果正确，则查找城市（在另一个函数中），并返回值。</span><span class="sxs-lookup"><span data-stu-id="7681d-116">If it is correct, then it will look up the city (in another function) and return the value.</span></span> <span data-ttu-id="7681d-117">如果不正确，则会将 `#VALUE!` 错误返回到单元格。</span><span class="sxs-lookup"><span data-stu-id="7681d-117">If it is not correct, it returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="7681d-118">CustomFunctions.Error 对象</span><span class="sxs-lookup"><span data-stu-id="7681d-118">The CustomFunctions.Error object</span></span>

<span data-ttu-id="7681d-119">`CustomFunctions.Error` 对象用于将错误返回单元格。</span><span class="sxs-lookup"><span data-stu-id="7681d-119">The `CustomFunctions.Error` object is used to return an error back to the cell.</span></span> <span data-ttu-id="7681d-120">创建对象时，请使用以下 `ErrorCode` 枚举值之一指定要使用的错误。</span><span class="sxs-lookup"><span data-stu-id="7681d-120">When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="7681d-121">ErrorCode 枚举值</span><span class="sxs-lookup"><span data-stu-id="7681d-121">ErrorCode enum value</span></span>  |<span data-ttu-id="7681d-122">Excel 单元格值</span><span class="sxs-lookup"><span data-stu-id="7681d-122">Excel cell value</span></span>  |<span data-ttu-id="7681d-123">含义</span><span class="sxs-lookup"><span data-stu-id="7681d-123">Meaning</span></span>  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="7681d-124">公式中使用的一个值为错误类型。</span><span class="sxs-lookup"><span data-stu-id="7681d-124">A value used in the formula is the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="7681d-125">函数或服务不可用。</span><span class="sxs-lookup"><span data-stu-id="7681d-125">The function or service is not available.</span></span> |
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="7681d-126">请注意，JavaScript 允许除以零，因此你需要仔细编写一个错误处理程序来检测这种情况。</span><span class="sxs-lookup"><span data-stu-id="7681d-126">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="7681d-127">公式中使用的数字有问题</span><span class="sxs-lookup"><span data-stu-id="7681d-127">There is a problem with the number used in the formula</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="7681d-128">公式中的范围不相交。</span><span class="sxs-lookup"><span data-stu-id="7681d-128">The ranges in the formula do not intersect.</span></span> |

<span data-ttu-id="7681d-129">下面的代码示例演示了如何创建并返回无效数字 (`#NUM!`) 错误。</span><span class="sxs-lookup"><span data-stu-id="7681d-129">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="7681d-130">返回 `#VALUE!` 错误时，还可以添加当用户将鼠标悬停在单元格上方时将会弹出的自定义消息。</span><span class="sxs-lookup"><span data-stu-id="7681d-130">When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell.</span></span> <span data-ttu-id="7681d-131">下面的示例演示了如何返回自定义错误消息。</span><span class="sxs-lookup"><span data-stu-id="7681d-131">The following example shows how to return a custom error message.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, “The parameter can only contain lowercase characters.”);
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="7681d-132">使用 try-catch 块</span><span class="sxs-lookup"><span data-stu-id="7681d-132">Use try-catch blocks</span></span>

<span data-ttu-id="7681d-133">通常情况下，应在自定义函数中使用 `try`-`catch` 块来捕获发生的任何潜在错误。</span><span class="sxs-lookup"><span data-stu-id="7681d-133">In general, you should use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="7681d-134">如果不在代码中处理异常，它们将返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="7681d-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="7681d-135">默认情况下，对于未处理的异常，Excel 返回 `#VALUE!`。</span><span class="sxs-lookup"><span data-stu-id="7681d-135">By default, Excel returns `#VALUE!` for an unhandled exception.</span></span>

<span data-ttu-id="7681d-136">在下面的代码示例中，自定义函数对 REST 服务执行 fetch 调用。</span><span class="sxs-lookup"><span data-stu-id="7681d-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="7681d-137">此调用有可能会失败，例如，如果 REST 服务返回错误或网络中断，就可能会失败。</span><span class="sxs-lookup"><span data-stu-id="7681d-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="7681d-138">如果发生这种情况，自定义函数将返回 `#N/A` 以指示 Web 调用失败。</span><span class="sxs-lookup"><span data-stu-id="7681d-138">If this happens, the custom function will return `#N/A` to indicate the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="7681d-139">后续步骤</span><span class="sxs-lookup"><span data-stu-id="7681d-139">Next steps</span></span>

<span data-ttu-id="7681d-140">了解如何[解决自定义函数中的问题](custom-functions-troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="7681d-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="7681d-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7681d-141">See also</span></span>

* [<span data-ttu-id="7681d-142">自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="7681d-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="7681d-143">自定义函数要求</span><span class="sxs-lookup"><span data-stu-id="7681d-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="7681d-144">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="7681d-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
