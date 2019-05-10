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
# <a name="error-handling-within-custom-functions"></a>自定义函数中的错误处理

在生成定义自定义函数的加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 自定义函数的错误处理与 [Excel JavaScript API 的错误处理](excel-add-ins-error-handling.md)大致相同。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

在以下代码示例中，`.catch` 将处理之前发生在代码中的任何错误。

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

## <a name="next-steps"></a>后续步骤
了解如何[解决自定义函数中的问题](custom-functions-troubleshooting.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数调试](custom-functions-debugging.md)
* [自定义函数要求](custom-functions-requirements.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
