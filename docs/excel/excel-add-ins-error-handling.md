---
title: 错误处理
description: 了解有关 Excel JavaScript API 错误处理逻辑，以解决运行时错误。
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: cb06ab890fc10afd327bde0eb708ff97fabbcdca
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430553"
---
# <a name="error-handling"></a><span data-ttu-id="5a5a2-103">错误处理</span><span class="sxs-lookup"><span data-stu-id="5a5a2-103">Error handling</span></span>

<span data-ttu-id="5a5a2-p101">使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 鉴于 API 的异步特性，这样做非常关键。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="5a5a2-106">有关 `sync()` Excel JAVASCRIPT API 的方法和异步特性的详细信息，请参阅 [Office 外接程序中的 Excel JavaScript 对象模型](excel-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="5a5a2-107">最佳做法</span><span class="sxs-lookup"><span data-stu-id="5a5a2-107">Best practices</span></span>

<span data-ttu-id="5a5a2-p102">通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。 建议在使用 Excel JavaScript API 生成加载项时使用相同模式。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="5a5a2-110">API 错误</span><span class="sxs-lookup"><span data-stu-id="5a5a2-110">API errors</span></span>

<span data-ttu-id="5a5a2-111">当 Excel JavaScript API 请求无法成功运行时，API 将返回错误对象，其中包含以下属性：</span><span class="sxs-lookup"><span data-stu-id="5a5a2-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="5a5a2-p103">**代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。 例如，错误代码“InvalidReference”指示引用对于指定操作无效。 错误代码尚未本地化。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="5a5a2-115">**消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="5a5a2-116">错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定外接程序显示给最终用户的错误消息。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="5a5a2-117">**debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="5a5a2-118">如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息只会在服务器上可见。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="5a5a2-119">最终用户将不会在加载项任务窗格中或 Office 应用程序中的任何位置看到这些错误消息。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-119">End users will not see those error messages in the add-in task pane or anywhere in the Office application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="5a5a2-120">错误消息</span><span class="sxs-lookup"><span data-stu-id="5a5a2-120">Error Messages</span></span>

<span data-ttu-id="5a5a2-121">下表是 API 可能返回的错误列表。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="5a5a2-122">error.code</span><span class="sxs-lookup"><span data-stu-id="5a5a2-122">error.code</span></span> | <span data-ttu-id="5a5a2-123">error.message</span><span class="sxs-lookup"><span data-stu-id="5a5a2-123">error.message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="5a5a2-124">无法执行所请求的操作。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="5a5a2-125">已达到活动限制。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="5a5a2-126">请求的 API 不可用。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-126">The requested API is not available.</span></span>|
|`Conflict`|<span data-ttu-id="5a5a2-127">由于冲突，无法处理请求。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-127">Request could not be processed because of a conflict.</span></span>|
|`GeneralException`|<span data-ttu-id="5a5a2-128">处理请求时出现内部错误。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-128">There was an internal error while processing the request.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="5a5a2-129">尝试的插入或删除操作导致冲突。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-129">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="5a5a2-130">自变量无效、缺少或格式不正确。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-130">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="5a5a2-131">由于之前的更新，此对象绑定不再有效。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-131">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="5a5a2-132">尝试的操作对于对象无效。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-132">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="5a5a2-133">此引用对于当前操作无效。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-133">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="5a5a2-134">无法处理此请求。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-134">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="5a5a2-135">当前选定内容对于此操作无效。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-135">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="5a5a2-136">所创建的资源已存在。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-136">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="5a5a2-137">所请求的资源不存在。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-137">The requested resource doesn't exist.</span></span>|
|`NotImplemented`  |<span data-ttu-id="5a5a2-138">所请求的功能未实现。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-138">The requested feature isn't implemented.</span></span>|
|`RequestAborted`|<span data-ttu-id="5a5a2-139">请求在运行时已中止。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-139">The request was aborted during run time.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="5a5a2-140">服务不可用。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-140">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="5a5a2-141">所需的身份验证信息缺少或无效。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-141">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="5a5a2-142">不支持正在尝试的操作。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-142">The operation being attempted is not supported.</span></span>|
|`UnsupportedSheet`|<span data-ttu-id="5a5a2-143">此工作表类型不支持此操作，因为它是宏或图表工作表。</span><span class="sxs-lookup"><span data-stu-id="5a5a2-143">This sheet type does not support this operation, since it is a Macro or Chart sheet.</span></span>|

## <a name="see-also"></a><span data-ttu-id="5a5a2-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5a5a2-144">See also</span></span>

- [<span data-ttu-id="5a5a2-145">Office 外接程序中的 Excel JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="5a5a2-145">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="5a5a2-146">OfficeExtension.Error 对象（Excel JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="5a5a2-146">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
