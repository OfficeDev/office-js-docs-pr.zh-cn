---
ms.date: 02/08/2019
description: 处理 Excel 自定义函数中的错误。
title: Excel 中自定义函数的错误处理（预览）
localization_priority: Priority
ms.openlocfilehash: 170da03331663d6779bed7bf0bf5a9b75b908b3f
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632693"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="bf5a5-103">自定义函数中的错误处理</span><span class="sxs-lookup"><span data-stu-id="bf5a5-103">Error handling within custom functions</span></span>

<span data-ttu-id="bf5a5-104">在生成定义自定义函数的加载项时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="bf5a5-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="bf5a5-105">自定义函数的错误处理与 [Excel JavaScript API 的错误处理](excel-add-ins-error-handling.md)大致相同。</span><span class="sxs-lookup"><span data-stu-id="bf5a5-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="bf5a5-106">在以下代码示例中，`.catch` 将处理之前发生在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="bf5a5-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="see-also"></a><span data-ttu-id="bf5a5-107">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bf5a5-107">See also</span></span>

* [<span data-ttu-id="bf5a5-108">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="bf5a5-108">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="bf5a5-109">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="bf5a5-109">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="bf5a5-110">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="bf5a5-110">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="bf5a5-111">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="bf5a5-111">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="bf5a5-112">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="bf5a5-112">Custom functions changelog</span></span>](custom-functions-changelog.md)
