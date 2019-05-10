---
ms.date: 05/03/2019
description: 处理 Excel 自定义函数中的错误。
title: 在 Excel 中处理自定义函数时出错
localization_priority: Priority
ms.openlocfilehash: 188ece6c77bc2cafad6f22448fb698e0c0370ef8
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628156"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="d1125-103">自定义函数中的错误处理</span><span class="sxs-lookup"><span data-stu-id="d1125-103">Error handling within custom functions</span></span>

<span data-ttu-id="d1125-104">在生成定义自定义函数的加载项时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="d1125-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="d1125-105">自定义函数的错误处理与 [Excel JavaScript API 的错误处理](excel-add-ins-error-handling.md)大致相同。</span><span class="sxs-lookup"><span data-stu-id="d1125-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="d1125-106">在以下代码示例中，`.catch` 将处理之前发生在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="d1125-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="next-steps"></a><span data-ttu-id="d1125-107">后续步骤</span><span class="sxs-lookup"><span data-stu-id="d1125-107">Next steps</span></span>
<span data-ttu-id="d1125-108">了解如何[解决自定义函数中的问题](custom-functions-troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="d1125-108">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d1125-109">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d1125-109">See also</span></span>

* [<span data-ttu-id="d1125-110">自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="d1125-110">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="d1125-111">自定义函数要求</span><span class="sxs-lookup"><span data-stu-id="d1125-111">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="d1125-112">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="d1125-112">Create custom functions in Excel</span></span>](custom-functions-overview.md)
