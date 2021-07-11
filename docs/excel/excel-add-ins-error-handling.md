---
title: JavaScript API Excel错误处理
description: 了解如何Excel JavaScript API 错误处理逻辑，以考虑运行时错误。
ms.date: 01/15/2021
localization_priority: Normal
ms.openlocfilehash: 42ef52b5d20a2c2d1284f57c7b4026ff2c71ebdd
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349509"
---
# <a name="error-handling-with-the-excel-javascript-api"></a><span data-ttu-id="9cd0c-103">JavaScript API Excel错误处理</span><span class="sxs-lookup"><span data-stu-id="9cd0c-103">Error handling with the Excel JavaScript API</span></span>

<span data-ttu-id="9cd0c-p101">使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 鉴于 API 的异步特性，这样做非常关键。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="9cd0c-106">有关 JavaScript API 的方法和异步特性 `sync()` Excel，请参阅 Excel 外接程序中的[Office JavaScript 对象模型](excel-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="9cd0c-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="9cd0c-107">Best practices</span></span>

<span data-ttu-id="9cd0c-p102">通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。 建议在使用 Excel JavaScript API 生成加载项时使用相同模式。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

```js
Excel.run(function (context) {
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);
```

## <a name="api-errors"></a><span data-ttu-id="9cd0c-110">API 错误</span><span class="sxs-lookup"><span data-stu-id="9cd0c-110">API errors</span></span>

<span data-ttu-id="9cd0c-111">当 Excel JavaScript API 请求无法成功运行时，API 将返回一个包含以下属性的错误对象。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties.</span></span>

- <span data-ttu-id="9cd0c-p103">**代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。 例如，错误代码“InvalidReference”指示引用对于指定操作无效。 错误代码尚未本地化。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="9cd0c-115">**消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="9cd0c-116">错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定外接程序显示给最终用户的错误消息。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="9cd0c-117">**debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="9cd0c-118">如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息只会在服务器上可见。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="9cd0c-119">最终用户不会在加载项任务窗格或应用程序中的任何位置看到这些Office消息。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-119">End users will not see those error messages in the add-in task pane or anywhere in the Office application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="9cd0c-120">错误消息</span><span class="sxs-lookup"><span data-stu-id="9cd0c-120">Error Messages</span></span>

<span data-ttu-id="9cd0c-121">下表是 API 可能返回的错误列表。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="9cd0c-122">错误代码</span><span class="sxs-lookup"><span data-stu-id="9cd0c-122">Error code</span></span> | <span data-ttu-id="9cd0c-123">错误消息</span><span class="sxs-lookup"><span data-stu-id="9cd0c-123">Error message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="9cd0c-124">无法执行所请求的操作。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="9cd0c-125">已达到活动限制。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="9cd0c-126">请求的 API 不可用。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-126">The requested API is not available.</span></span>|
|`ApiNotFound`|<span data-ttu-id="9cd0c-127">找不到您尝试使用的 API。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-127">The API you are trying to use could not be found.</span></span> <span data-ttu-id="9cd0c-128">它在更高版本的 Excel 中Excel。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-128">It may be available in a newer version of Excel.</span></span> <span data-ttu-id="9cd0c-129">有关详细信息[，Excel JavaScript API](../reference/requirement-sets/excel-api-requirement-sets.md)要求集文章。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-129">See the [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) article for more information.</span></span>|
|`BadPassword`|<span data-ttu-id="9cd0c-130">你提供的密码不正确。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-130">The password you supplied is incorrect.</span></span>|
|`Conflict`|<span data-ttu-id="9cd0c-131">由于冲突，无法处理请求。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-131">Request could not be processed because of a conflict.</span></span>|
|`ContentLengthRequired`|<span data-ttu-id="9cd0c-132">`Content-length`HTTP 标头缺失。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-132">A `Content-length` HTTP header is missing.</span></span>|
|`GeneralException`|<span data-ttu-id="9cd0c-133">处理请求时出现内部错误。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-133">There was an internal error while processing the request.</span></span>|
|`InactiveWorkbook`|<span data-ttu-id="9cd0c-134">操作失败，因为多个工作簿已打开，并且此 API 调用的工作簿已失去焦点。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-134">The operation failed because multiple workbooks are open and the workbook being called by this API has lost focus.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="9cd0c-135">尝试的插入或删除操作导致冲突。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-135">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="9cd0c-136">自变量无效、缺少或格式不正确。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-136">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="9cd0c-137">由于之前的更新，此对象绑定不再有效。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-137">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="9cd0c-138">尝试的操作对于对象无效。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-138">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="9cd0c-139">此引用对于当前操作无效。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-139">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="9cd0c-140">无法处理此请求。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-140">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="9cd0c-141">当前选定内容对于此操作无效。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-141">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="9cd0c-142">所创建的资源已存在。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-142">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="9cd0c-143">所请求的资源不存在。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-143">The requested resource doesn't exist.</span></span>|
|`NonBlankCellOffSheet`|<span data-ttu-id="9cd0c-144">Microsoft Excel无法插入新单元格，因为它会将非空单元格推送到工作表末尾。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-144">Microsoft Excel can't insert new cells because it would push non-empty cells off the end of the worksheet.</span></span> <span data-ttu-id="9cd0c-145">这些非空单元格可能为空，但具有空值、某些格式或公式。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-145">These non-empty cells might appear empty but have blank values, some formatting, or a formula.</span></span> <span data-ttu-id="9cd0c-146">删除足够的行或列，为要插入的行或列提供空间，然后重试。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-146">Delete enough rows or columns to make room for what you want to insert and then try again.</span></span>|
|`NotImplemented`|<span data-ttu-id="9cd0c-147">所请求的功能未实现。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-147">The requested feature isn't implemented.</span></span>|
|`RangeExceedsLimit`|<span data-ttu-id="9cd0c-148">该范围中的单元格计数已超出支持的最大数。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-148">The cell count in the range has exceeded the maximum supported number.</span></span> <span data-ttu-id="9cd0c-149">有关详细信息[，请参阅Office加载项](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)的资源限制和性能优化一文。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-149">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>|
|`RequestAborted`|<span data-ttu-id="9cd0c-150">请求在运行时已中止。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-150">The request was aborted during run time.</span></span>|
|`RequestPayloadSizeLimitExceeded`|<span data-ttu-id="9cd0c-151">请求有效负载大小已超出限制。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-151">The request payload size has exceeded the limit.</span></span> <span data-ttu-id="9cd0c-152">有关详细信息[，请参阅Office加载项](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)的资源限制和性能优化一文。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-152">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span> <br><br><span data-ttu-id="9cd0c-153">此错误仅出现在Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-153">This error only occurs in Excel on the web.</span></span>|
|`ResponsePayloadSizeLimitExceeded`|<span data-ttu-id="9cd0c-154">响应有效负载大小已超出限制。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-154">The response payload size has exceeded the limit.</span></span> <span data-ttu-id="9cd0c-155">有关详细信息[，请参阅Office加载项](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)的资源限制和性能优化一文。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-155">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>  <br><br><span data-ttu-id="9cd0c-156">此错误仅出现在Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-156">This error only occurs in Excel on the web.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="9cd0c-157">服务不可用。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-157">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="9cd0c-158">所需的身份验证信息缺少或无效。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-158">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="9cd0c-159">不支持正在尝试的操作。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-159">The operation being attempted is not supported.</span></span>|
|`UnsupportedSheet`|<span data-ttu-id="9cd0c-160">此工作表类型不支持此操作，因为它是一个宏或图表工作表。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-160">This sheet type does not support this operation, since it is a Macro or Chart sheet.</span></span>|

> [!NOTE]
> <span data-ttu-id="9cd0c-161">上表列出了使用 JavaScript API 时Excel错误消息。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-161">The preceding table lists error messages you may encounter while using the Excel JavaScript API.</span></span> <span data-ttu-id="9cd0c-162">如果你使用通用 API 而不是特定于应用程序的 JavaScript API Excel，请参阅Office[通用 API](../reference/javascript-api-for-office-error-codes.md)错误代码以了解相关的错误消息。</span><span class="sxs-lookup"><span data-stu-id="9cd0c-162">If you are working with the Common API instead of the application-specific Excel JavaScript API, see [Office Common API error codes](../reference/javascript-api-for-office-error-codes.md) to learn about relevant error messages.</span></span>

## <a name="see-also"></a><span data-ttu-id="9cd0c-163">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9cd0c-163">See also</span></span>

- [<span data-ttu-id="9cd0c-164">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="9cd0c-164">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9cd0c-165">OfficeExtension.Error 对象（Excel JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="9cd0c-165">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [<span data-ttu-id="9cd0c-166">Office 常用 API 错误代码</span><span class="sxs-lookup"><span data-stu-id="9cd0c-166">Office Common API error codes</span></span>](../reference/javascript-api-for-office-error-codes.md)
