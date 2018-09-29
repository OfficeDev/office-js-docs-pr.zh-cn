---
title: 错误处理
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 23a70b1d66befb971c3c1394eb9162c19f2ee176
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348084"
---
# <a name="error-handling"></a><span data-ttu-id="4a989-102">错误处理</span><span class="sxs-lookup"><span data-stu-id="4a989-102">Error handling</span></span>

<span data-ttu-id="4a989-103">使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="4a989-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="4a989-104">鉴于 API 的异步特性，这样做非常关键。</span><span class="sxs-lookup"><span data-stu-id="4a989-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="4a989-105">若要详细了解 **sync()** 方法和 Excel JavaScript API 异步特性，请参阅 [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="4a989-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="4a989-106">最佳做法</span><span class="sxs-lookup"><span data-stu-id="4a989-106">Best practices</span></span>

<span data-ttu-id="4a989-107">通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。</span><span class="sxs-lookup"><span data-stu-id="4a989-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="4a989-108">建议在使用 Excel JavaScript API 生成加载项时使用相同模式。</span><span class="sxs-lookup"><span data-stu-id="4a989-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="4a989-109">API 错误</span><span class="sxs-lookup"><span data-stu-id="4a989-109">API errors</span></span> 

<span data-ttu-id="4a989-110">当 Excel JavaScript API 请求无法成功运行时，API 将返回错误对象，其中包含以下属性：</span><span class="sxs-lookup"><span data-stu-id="4a989-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="4a989-111">**代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。</span><span class="sxs-lookup"><span data-stu-id="4a989-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="4a989-112">例如，错误代码“InvalidReference”指示引用对于指定操作无效。</span><span class="sxs-lookup"><span data-stu-id="4a989-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="4a989-113">错误代码尚未本地化。</span><span class="sxs-lookup"><span data-stu-id="4a989-113">Error codes are not localized.</span></span> 

- <span data-ttu-id="4a989-114">**消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。</span><span class="sxs-lookup"><span data-stu-id="4a989-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="4a989-115">错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定加载项显示给最终用户的错误消息。</span><span class="sxs-lookup"><span data-stu-id="4a989-115">The error message is not intended for end-user consumption; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end-users.</span></span>

- <span data-ttu-id="4a989-116">**debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误的根本原因。</span><span class="sxs-lookup"><span data-stu-id="4a989-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="4a989-117">如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息只会在服务器上可见。</span><span class="sxs-lookup"><span data-stu-id="4a989-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="4a989-118">最终用户不会在加载项任务窗格或主机应用的其他任何位置看到这些错误消息。</span><span class="sxs-lookup"><span data-stu-id="4a989-118">End-users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="see-also"></a><span data-ttu-id="4a989-119">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4a989-119">See also</span></span>

- [<span data-ttu-id="4a989-120">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="4a989-120">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4a989-121">OfficeExtension.Error 对象 (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="4a989-121">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
