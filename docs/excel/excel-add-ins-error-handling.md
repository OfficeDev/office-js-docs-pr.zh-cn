---
title: 错误处理
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 52d1c88fef0f4e3af075ed625f856b029353a963
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533712"
---
# <a name="error-handling"></a><span data-ttu-id="c7ece-102">错误处理</span><span class="sxs-lookup"><span data-stu-id="c7ece-102">Error handling</span></span>

<span data-ttu-id="c7ece-103">使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="c7ece-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="c7ece-104">鉴于 API 的异步特性，这样做非常关键。</span><span class="sxs-lookup"><span data-stu-id="c7ece-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="c7ece-105">若要详细了解 **sync()** 方法和 Excel JavaScript API 的异步特性，请参阅 [Excel JavaScript API 的基本编程概念](excel-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="c7ece-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="c7ece-106">最佳做法</span><span class="sxs-lookup"><span data-stu-id="c7ece-106">Best practices</span></span>

<span data-ttu-id="c7ece-107">通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。</span><span class="sxs-lookup"><span data-stu-id="c7ece-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="c7ece-108">建议在使用 Excel JavaScript API 生成加载项时使用相同模式。</span><span class="sxs-lookup"><span data-stu-id="c7ece-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="c7ece-109">API 错误</span><span class="sxs-lookup"><span data-stu-id="c7ece-109">API errors</span></span>

<span data-ttu-id="c7ece-110">当 Excel JavaScript API 请求无法成功运行时，API 将返回错误对象，其中包含以下属性：</span><span class="sxs-lookup"><span data-stu-id="c7ece-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="c7ece-111">**代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。</span><span class="sxs-lookup"><span data-stu-id="c7ece-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="c7ece-112">例如，错误代码“InvalidReference”指示引用对于指定操作无效。</span><span class="sxs-lookup"><span data-stu-id="c7ece-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="c7ece-113">错误代码尚未本地化。</span><span class="sxs-lookup"><span data-stu-id="c7ece-113">Error codes are not localized.</span></span>

- <span data-ttu-id="c7ece-114">**消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。</span><span class="sxs-lookup"><span data-stu-id="c7ece-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="c7ece-115">错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定外接程序显示给最终用户的错误消息。</span><span class="sxs-lookup"><span data-stu-id="c7ece-115">The error message is not intended for end-user consumption; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end-users.</span></span>

- <span data-ttu-id="c7ece-116">**debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。</span><span class="sxs-lookup"><span data-stu-id="c7ece-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="c7ece-117">如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息只会在服务器上可见。</span><span class="sxs-lookup"><span data-stu-id="c7ece-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="c7ece-118">最终用户将不会在外接程序任务窗格中或主机应用程序的任何位置看到这些错误消息。</span><span class="sxs-lookup"><span data-stu-id="c7ece-118">End-users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="c7ece-119">错误消息</span><span class="sxs-lookup"><span data-stu-id="c7ece-119">Error Messages</span></span>

<span data-ttu-id="c7ece-120">下表是 API 可能返回的错误列表。</span><span class="sxs-lookup"><span data-stu-id="c7ece-120">The following table defines a list of errors that the API may return.</span></span>

|<span data-ttu-id="c7ece-121">error.code</span><span class="sxs-lookup"><span data-stu-id="c7ece-121">ErrorCode</span></span> | <span data-ttu-id="c7ece-122">error.message</span><span class="sxs-lookup"><span data-stu-id="c7ece-122">ErrorMessage</span></span> |
|:----------|:--------------|
|<span data-ttu-id="c7ece-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="c7ece-123">InvalidArgument</span></span> |<span data-ttu-id="c7ece-124">参数无效或缺少或格式不正确。</span><span class="sxs-lookup"><span data-stu-id="c7ece-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="c7ece-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="c7ece-125">invalidRequest</span></span>  |<span data-ttu-id="c7ece-126">无法处理此请求。</span><span class="sxs-lookup"><span data-stu-id="c7ece-126">Cannot process the request.</span></span>|
|<span data-ttu-id="c7ece-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="c7ece-127">InvalidReference</span></span>|<span data-ttu-id="c7ece-128">此引用对于当前操作无效。</span><span class="sxs-lookup"><span data-stu-id="c7ece-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="c7ece-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="c7ece-129">InvalidBinding</span></span>  |<span data-ttu-id="c7ece-130">由于之前的更新，此对象绑定不再有效。</span><span class="sxs-lookup"><span data-stu-id="c7ece-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="c7ece-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c7ece-131">InvalidSelection</span></span>|<span data-ttu-id="c7ece-132">当前选定内容对于此操作无效。</span><span class="sxs-lookup"><span data-stu-id="c7ece-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="c7ece-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="c7ece-133">Unauthenticated</span></span> |<span data-ttu-id="c7ece-134">所需的身份验证信息缺少或无效。</span><span class="sxs-lookup"><span data-stu-id="c7ece-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="c7ece-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="c7ece-135">AccessDenied</span></span> |<span data-ttu-id="c7ece-136">无法执行所请求的操作。</span><span class="sxs-lookup"><span data-stu-id="c7ece-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="c7ece-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="c7ece-137">itemNotFound</span></span> |<span data-ttu-id="c7ece-138">所请求的资源不存在。</span><span class="sxs-lookup"><span data-stu-id="c7ece-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="c7ece-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="c7ece-139">activityLimitReached</span></span>|<span data-ttu-id="c7ece-140">已达到活动限制。</span><span class="sxs-lookup"><span data-stu-id="c7ece-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="c7ece-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7ece-141">generalException</span></span>|<span data-ttu-id="c7ece-142">处理请求时出现内部错误。</span><span class="sxs-lookup"><span data-stu-id="c7ece-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="c7ece-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="c7ece-143">NotImplemented</span></span>  |<span data-ttu-id="c7ece-144">所请求的功能未实现。</span><span class="sxs-lookup"><span data-stu-id="c7ece-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="c7ece-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="c7ece-145">serviceNotAvailable</span></span>|<span data-ttu-id="c7ece-146">服务不可用。</span><span class="sxs-lookup"><span data-stu-id="c7ece-146">The service is unavailable.</span></span>|
|<span data-ttu-id="c7ece-147">Conflict</span><span class="sxs-lookup"><span data-stu-id="c7ece-147">Conflict</span></span>|<span data-ttu-id="c7ece-148">由于冲突，无法处理请求。</span><span class="sxs-lookup"><span data-stu-id="c7ece-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="c7ece-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="c7ece-149">ItemAlreadyExists</span></span>|<span data-ttu-id="c7ece-150">所创建的资源已存在。</span><span class="sxs-lookup"><span data-stu-id="c7ece-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="c7ece-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="c7ece-151">UnsupportedOperation</span></span>|<span data-ttu-id="c7ece-152">不支持正在尝试的操作。</span><span class="sxs-lookup"><span data-stu-id="c7ece-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="c7ece-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="c7ece-153">RequestAborted</span></span>|<span data-ttu-id="c7ece-154">请求在运行时已中止。</span><span class="sxs-lookup"><span data-stu-id="c7ece-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="c7ece-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="c7ece-155">ApiNotAvailable</span></span>|<span data-ttu-id="c7ece-156">请求的 API 不可用。</span><span class="sxs-lookup"><span data-stu-id="c7ece-156">The requested API is not available.</span></span>|
|<span data-ttu-id="c7ece-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="c7ece-157">InsertDeleteConflict</span></span>|<span data-ttu-id="c7ece-158">尝试的插入或删除操作导致冲突。</span><span class="sxs-lookup"><span data-stu-id="c7ece-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="c7ece-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="c7ece-159">InvalidOperation</span></span>|<span data-ttu-id="c7ece-160">尝试的操作对于对象无效。</span><span class="sxs-lookup"><span data-stu-id="c7ece-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="c7ece-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c7ece-161">See also</span></span>

- [<span data-ttu-id="c7ece-162">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="c7ece-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c7ece-163">OfficeExtension.Error 对象（Excel JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="c7ece-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
