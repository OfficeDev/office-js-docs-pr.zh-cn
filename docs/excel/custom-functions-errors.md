---
ms.date: 05/06/2020
description: '处理和返回自定义函数中类似 #NULL! 的错误'
title: 处理和返回自定义函数中的错误（预览）
localization_priority: Normal
ms.openlocfilehash: 5598db045920a5fb419ed91435c3e81275dfb3a4
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275901"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a>处理和返回自定义函数中的错误（预览）

> [!NOTE]
> 本文中所述的功能目前处于预览阶段，可能会发生更改。 暂不支持在生产环境中使用。 你将需要加入[Office 预览体验成员](https://insider.office.com/join)计划，以试用预览版功能。  试用预览版功能的好方法是使用 Office 365 订阅。 如果你还没有 Office 365 订阅，可以通过加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获得 90 天免费的可续订 Office 365 订阅。

如果自定义函数运行时出现错误，则返回一个错误以通知用户。 如果您有特定参数要求（如仅正数），请测试参数并在它们不正确时引发错误。 还可以使用 `try`-`catch` 块来捕获自定义函数运行时发生的任何错误。

## <a name="detect-and-throw-an-error"></a>检测和引发错误

我们来看一种需要确保邮政编码参数格式正确的自定义函数能够正常工作的情况。 下面的自定义函数使用正则表达式来检查邮政编码。 如果是正确的，它将使用另一个函数查找城市，并返回值。 如果不正确，则 `#VALUE!` 向单元格返回一个错误。

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

## <a name="the-customfunctionserror-object"></a>CustomFunctions.Error 对象

`CustomFunctions.Error` 对象用于将错误返回单元格。 创建对象时，请使用以下 `ErrorCode` 枚举值之一指定要使用的错误。


|ErrorCode 枚举值  |Excel 单元格值  |含义  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | 公式中使用的一个值为错误类型。 |
|`notAvailable`   | `#N/A`    | 函数或服务不可用。 |
|`divisionByZero` | `#DIV/0`  | 请注意，JavaScript 允许除以零，因此你需要仔细编写一个错误处理程序来检测这种情况。 |
|`invalidNumber`  | `#NUM!`   | 公式中使用的数字有问题 |
|`nullReference`  | `#NULL!`  | 公式中的区域不相交。 |

下面的代码示例演示了如何创建并返回无效数字 (`#NUM!`) 错误。

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

返回 `#VALUE!` 错误时，还可以添加当用户将鼠标悬停在单元格上方时将会弹出的自定义消息。 下面的示例演示了如何返回自定义错误消息。

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>使用 try-catch 块

通常情况下，使用 `try` - `catch` 自定义函数中的块捕捉出现的任何潜在错误。 如果不在代码中处理异常，它们将返回到 Excel。 默认情况下，对于未处理的异常，Excel 返回 `#VALUE!`。

在下面的代码示例中，自定义函数对 REST 服务执行 fetch 调用。 此调用有可能会失败，例如，如果 REST 服务返回错误或网络中断，就可能会失败。 如果发生这种情况，自定义函数将返回 `#N/A` 以指示 Web 调用失败。


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

## <a name="next-steps"></a>后续步骤

了解如何[解决自定义函数中的问题](custom-functions-troubleshooting.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数调试](custom-functions-debugging.md)
* [自定义函数要求](custom-functions-requirement-sets.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
