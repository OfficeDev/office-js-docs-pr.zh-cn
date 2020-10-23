---
title: 使用 Excel JavaScript API 处理错误
description: 了解有关 Excel JavaScript API 错误处理逻辑，以解决运行时错误。
ms.date: 10/22/2020
localization_priority: Normal
ms.openlocfilehash: a3b1bbfa7daba1b856bce35aa075d5b625bd9769
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740817"
---
# <a name="error-handling-with-the-excel-javascript-api"></a><span data-ttu-id="0aaa3-103">使用 Excel JavaScript API 处理错误</span><span class="sxs-lookup"><span data-stu-id="0aaa3-103">Error handling with the Excel JavaScript API</span></span>

<span data-ttu-id="0aaa3-p101">使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 鉴于 API 的异步特性，这样做非常关键。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="0aaa3-106">有关 `sync()` Excel JAVASCRIPT API 的方法和异步特性的详细信息，请参阅 [Office 外接程序中的 Excel JavaScript 对象模型](excel-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="0aaa3-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="0aaa3-107">Best practices</span></span>

<span data-ttu-id="0aaa3-p102">通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。 建议在使用 Excel JavaScript API 生成加载项时使用相同模式。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="0aaa3-110">API 错误</span><span class="sxs-lookup"><span data-stu-id="0aaa3-110">API errors</span></span>

<span data-ttu-id="0aaa3-111">当 Excel JavaScript API 请求无法成功运行时，API 将返回错误对象，其中包含以下属性：</span><span class="sxs-lookup"><span data-stu-id="0aaa3-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="0aaa3-p103">**代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。 例如，错误代码“InvalidReference”指示引用对于指定操作无效。 错误代码尚未本地化。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="0aaa3-115">**消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="0aaa3-116">错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定外接程序显示给最终用户的错误消息。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="0aaa3-117">**debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="0aaa3-118">如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息只会在服务器上可见。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="0aaa3-119">最终用户将不会在加载项任务窗格中或 Office 应用程序中的任何位置看到这些错误消息。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-119">End users will not see those error messages in the add-in task pane or anywhere in the Office application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="0aaa3-120">错误消息</span><span class="sxs-lookup"><span data-stu-id="0aaa3-120">Error Messages</span></span>

<span data-ttu-id="0aaa3-121">下表是 API 可能返回的错误列表。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="0aaa3-122">错误代码</span><span class="sxs-lookup"><span data-stu-id="0aaa3-122">Error code</span></span> | <span data-ttu-id="0aaa3-123">错误消息</span><span class="sxs-lookup"><span data-stu-id="0aaa3-123">Error message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="0aaa3-124">无法执行所请求的操作。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="0aaa3-125">已达到活动限制。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="0aaa3-126">请求的 API 不可用。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-126">The requested API is not available.</span></span>|
|`ApiNotFound`|<span data-ttu-id="0aaa3-127">找不到您尝试使用的 API。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-127">The API you are trying to use could not be found.</span></span> <span data-ttu-id="0aaa3-128">它可能在较新版本的 Excel 中可用。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-128">It may be available in a newer version of Excel.</span></span> <span data-ttu-id="0aaa3-129">有关详细信息，请参阅 [Excel JAVASCRIPT API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md) 一文。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-129">See the [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) article for more information.</span></span>|
|`BadPassword`|<span data-ttu-id="0aaa3-130">你提供的密码不正确。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-130">The password you supplied is incorrect.</span></span>|
|`Conflict`|<span data-ttu-id="0aaa3-131">由于冲突，无法处理请求。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-131">Request could not be processed because of a conflict.</span></span>|
|`ContentLengthRequired`|<span data-ttu-id="0aaa3-132">`Content-length`缺少 HTTP 标头。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-132">A `Content-length` HTTP header is missing.</span></span>|
|`GeneralException`|<span data-ttu-id="0aaa3-133">处理请求时出现内部错误。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-133">There was an internal error while processing the request.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="0aaa3-134">尝试的插入或删除操作导致冲突。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-134">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="0aaa3-135">自变量无效、缺少或格式不正确。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-135">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="0aaa3-136">由于之前的更新，此对象绑定不再有效。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-136">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="0aaa3-137">尝试的操作对于对象无效。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-137">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="0aaa3-138">此引用对于当前操作无效。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-138">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="0aaa3-139">无法处理此请求。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-139">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="0aaa3-140">当前选定内容对于此操作无效。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-140">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="0aaa3-141">所创建的资源已存在。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-141">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="0aaa3-142">所请求的资源不存在。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-142">The requested resource doesn't exist.</span></span>|
|`NonBlankCellOffSheet`|<span data-ttu-id="0aaa3-143">插入新单元格的请求无法完成，因为它会将非空单元格推送到工作表的末尾。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-143">The request to insert new cells can't be completed because it would push non-empty cells off the end of the worksheet.</span></span> <span data-ttu-id="0aaa3-144">这些非空单元格可能显示为空，但具有空值、部分格式或公式。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-144">These non-empty cells might appear empty but have blank values, some formatting, or a formula.</span></span> <span data-ttu-id="0aaa3-145">删除足够多的行或列，为要插入的内容留出空间，然后重试。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-145">Delete enough rows or columns to make room for what you want to insert and then try again.</span></span>|
|`NotImplemented`|<span data-ttu-id="0aaa3-146">所请求的功能未实现。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-146">The requested feature isn't implemented.</span></span>|
|`RangeExceedsLimit`|<span data-ttu-id="0aaa3-147">区域中的单元格计数已超过支持的最大数量。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-147">The cell count in the range has exceeded the maximum supported number.</span></span> <span data-ttu-id="0aaa3-148">有关详细信息，请参阅 [Office 外接程序的资源限制和性能优化一](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 文。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-148">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>|
|`RequestAborted`|<span data-ttu-id="0aaa3-149">请求在运行时已中止。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-149">The request was aborted during run time.</span></span>|
|`RequestPayloadSizeLimitExceeded`|<span data-ttu-id="0aaa3-150">请求负载大小已超出限制。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-150">The request payload size has exceeded the limit.</span></span> <span data-ttu-id="0aaa3-151">有关详细信息，请参阅 [Office 外接程序的资源限制和性能优化一](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 文。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-151">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span> <br><br><span data-ttu-id="0aaa3-152">此错误仅发生在 web 上的 Excel 中。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-152">This error only occurs in Excel on the web.</span></span>|
|`ResponsePayloadSizeLimitExceeded`|<span data-ttu-id="0aaa3-153">响应负载大小已超出限制。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-153">The response payload size has exceeded the limit.</span></span> <span data-ttu-id="0aaa3-154">有关详细信息，请参阅 [Office 外接程序的资源限制和性能优化一](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 文。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-154">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>  <br><br><span data-ttu-id="0aaa3-155">此错误仅发生在 web 上的 Excel 中。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-155">This error only occurs in Excel on the web.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="0aaa3-156">服务不可用。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-156">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="0aaa3-157">所需的身份验证信息缺少或无效。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-157">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="0aaa3-158">不支持正在尝试的操作。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-158">The operation being attempted is not supported.</span></span>|
|`UnsupportedSheet`|<span data-ttu-id="0aaa3-159">此工作表类型不支持此操作，因为它是宏或图表工作表。</span><span class="sxs-lookup"><span data-stu-id="0aaa3-159">This sheet type does not support this operation, since it is a Macro or Chart sheet.</span></span>|

## <a name="see-also"></a><span data-ttu-id="0aaa3-160">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0aaa3-160">See also</span></span>

- [<span data-ttu-id="0aaa3-161">Office 外接程序中的 Excel JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="0aaa3-161">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="0aaa3-162">OfficeExtension.Error 对象（Excel JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="0aaa3-162">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
