---
title: 错误处理
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e3732af26aeaa6129a4b98d6cbb8e3caf501141f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325106"
---
# <a name="error-handling"></a><span data-ttu-id="57319-102">错误处理</span><span class="sxs-lookup"><span data-stu-id="57319-102">Error handling</span></span>

<span data-ttu-id="57319-p101">使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 鉴于 API 的异步特性，这样做非常关键。</span><span class="sxs-lookup"><span data-stu-id="57319-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="57319-105">有关 Excel JavaScript API 的`sync()`方法和异步特性的详细信息，请参阅[使用 excel Javascript api 的基本编程概念](excel-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="57319-105">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="57319-106">最佳做法</span><span class="sxs-lookup"><span data-stu-id="57319-106">Best practices</span></span>

<span data-ttu-id="57319-p102">通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。 建议在使用 Excel JavaScript API 生成加载项时使用相同模式。</span><span class="sxs-lookup"><span data-stu-id="57319-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="57319-109">API 错误</span><span class="sxs-lookup"><span data-stu-id="57319-109">API errors</span></span>

<span data-ttu-id="57319-110">当 Excel JavaScript API 请求无法成功运行时，API 将返回错误对象，其中包含以下属性：</span><span class="sxs-lookup"><span data-stu-id="57319-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="57319-p103">**代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。 例如，错误代码“InvalidReference”指示引用对于指定操作无效。 错误代码尚未本地化。</span><span class="sxs-lookup"><span data-stu-id="57319-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="57319-114">**消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。</span><span class="sxs-lookup"><span data-stu-id="57319-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="57319-115">错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定外接程序显示给最终用户的错误消息。</span><span class="sxs-lookup"><span data-stu-id="57319-115">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="57319-116">**debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。</span><span class="sxs-lookup"><span data-stu-id="57319-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="57319-117">如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息只会在服务器上可见。</span><span class="sxs-lookup"><span data-stu-id="57319-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="57319-118">最终用户将不会在外接程序任务窗格中或主机应用程序的任何位置看到这些错误消息。</span><span class="sxs-lookup"><span data-stu-id="57319-118">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="57319-119">错误消息</span><span class="sxs-lookup"><span data-stu-id="57319-119">Error Messages</span></span>

<span data-ttu-id="57319-120">下表是 API 可能返回的错误列表。</span><span class="sxs-lookup"><span data-stu-id="57319-120">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="57319-121">error.code</span><span class="sxs-lookup"><span data-stu-id="57319-121">error.code</span></span> | <span data-ttu-id="57319-122">error.message</span><span class="sxs-lookup"><span data-stu-id="57319-122">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="57319-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="57319-123">InvalidArgument</span></span> |<span data-ttu-id="57319-124">自变量无效、缺少或格式不正确。</span><span class="sxs-lookup"><span data-stu-id="57319-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="57319-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="57319-125">InvalidRequest</span></span>  |<span data-ttu-id="57319-126">无法处理此请求。</span><span class="sxs-lookup"><span data-stu-id="57319-126">Cannot process the request.</span></span>|
|<span data-ttu-id="57319-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="57319-127">InvalidReference</span></span>|<span data-ttu-id="57319-128">此引用对于当前操作无效。</span><span class="sxs-lookup"><span data-stu-id="57319-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="57319-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="57319-129">InvalidBinding</span></span>  |<span data-ttu-id="57319-130">由于之前的更新，此对象绑定不再有效。</span><span class="sxs-lookup"><span data-stu-id="57319-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="57319-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="57319-131">InvalidSelection</span></span>|<span data-ttu-id="57319-132">当前选定内容对于此操作无效。</span><span class="sxs-lookup"><span data-stu-id="57319-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="57319-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="57319-133">Unauthenticated</span></span> |<span data-ttu-id="57319-134">所需的身份验证信息缺少或无效。</span><span class="sxs-lookup"><span data-stu-id="57319-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="57319-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="57319-135">AccessDenied</span></span> |<span data-ttu-id="57319-136">无法执行所请求的操作。</span><span class="sxs-lookup"><span data-stu-id="57319-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="57319-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="57319-137">ItemNotFound</span></span> |<span data-ttu-id="57319-138">所请求的资源不存在。</span><span class="sxs-lookup"><span data-stu-id="57319-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="57319-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="57319-139">ActivityLimitReached</span></span>|<span data-ttu-id="57319-140">已达到活动限制。</span><span class="sxs-lookup"><span data-stu-id="57319-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="57319-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="57319-141">GeneralException</span></span>|<span data-ttu-id="57319-142">处理请求时出现内部错误。</span><span class="sxs-lookup"><span data-stu-id="57319-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="57319-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="57319-143">NotImplemented</span></span>  |<span data-ttu-id="57319-144">所请求的功能未实现。</span><span class="sxs-lookup"><span data-stu-id="57319-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="57319-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="57319-145">ServiceNotAvailable</span></span>|<span data-ttu-id="57319-146">服务不可用。</span><span class="sxs-lookup"><span data-stu-id="57319-146">The service is unavailable.</span></span>|
|<span data-ttu-id="57319-147">Conflict</span><span class="sxs-lookup"><span data-stu-id="57319-147">Conflict</span></span>|<span data-ttu-id="57319-148">由于冲突，无法处理请求。</span><span class="sxs-lookup"><span data-stu-id="57319-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="57319-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="57319-149">ItemAlreadyExists</span></span>|<span data-ttu-id="57319-150">所创建的资源已存在。</span><span class="sxs-lookup"><span data-stu-id="57319-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="57319-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="57319-151">UnsupportedOperation</span></span>|<span data-ttu-id="57319-152">不支持正在尝试的操作。</span><span class="sxs-lookup"><span data-stu-id="57319-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="57319-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="57319-153">RequestAborted</span></span>|<span data-ttu-id="57319-154">请求在运行时已中止。</span><span class="sxs-lookup"><span data-stu-id="57319-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="57319-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="57319-155">ApiNotAvailable</span></span>|<span data-ttu-id="57319-156">请求的 API 不可用。</span><span class="sxs-lookup"><span data-stu-id="57319-156">The requested API is not available.</span></span>|
|<span data-ttu-id="57319-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="57319-157">InsertDeleteConflict</span></span>|<span data-ttu-id="57319-158">尝试的插入或删除操作导致冲突。</span><span class="sxs-lookup"><span data-stu-id="57319-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="57319-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="57319-159">InvalidOperation</span></span>|<span data-ttu-id="57319-160">尝试的操作对于对象无效。</span><span class="sxs-lookup"><span data-stu-id="57319-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="57319-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="57319-161">See also</span></span>

- [<span data-ttu-id="57319-162">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="57319-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="57319-163">OfficeExtension.Error 对象（Excel JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="57319-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error)
