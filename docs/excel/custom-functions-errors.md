---
ms.date: 09/23/2020
description: '处理和返回自定义函数中类似 #NULL! 自定义函数中。'
title: 处理并返回自定义函数中的错误
localization_priority: Normal
ms.openlocfilehash: b3d3b325649a0775d3375c9f5285bba7cde0aa16
ms.sourcegitcommit: 09e1d8ff14b3c09a3eb11c91432c224a539181a4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2020
ms.locfileid: "48268542"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>处理并返回自定义函数中的错误

如果自定义函数运行时出现错误，则返回一个错误以通知用户。 如果您有特定参数要求（如仅正数），请测试参数并在它们不正确时引发错误。 还可以使用 `try`-`catch` 块来捕获自定义函数运行时发生的任何错误。

## <a name="detect-and-throw-an-error"></a>检测和引发错误

我们来看一种需要确保邮政编码参数格式正确的自定义函数能够正常工作的情况。 下面的自定义函数使用正则表达式来检查邮政编码。 如果邮政编码格式正确，则它将使用另一个函数查找城市并返回值。 如果格式无效，该函数将 `#VALUE!` 向单元格返回一个错误。

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

[Customfunctions.js](/javascript/api/custom-functions-runtime/customfunctions.error)对象用于将错误返回回单元格。 创建对象时，通过选择下列枚举值之一来指定要使用的错误 `ErrorCode` 。


|ErrorCode 枚举值  |Excel 单元格值  |Description  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | 函数试图被零除。 |
|`invalidName`    | `#NAME?`  | 函数名称中有拼写错误。 请注意，此错误被支持为自定义函数输入错误，而不是作为自定义函数输出错误。 | 
|`invalidNumber`  | `#NUM!`   | 公式中的数字有问题。 |
|`invalidReference` | `#REF!` | 函数引用了无效的单元格。 请注意，此错误被支持为自定义函数输入错误，而不是作为自定义函数输出错误。|
|`invalidValue`   | `#VALUE!` | 公式中的值的类型错误。 |
|`notAvailable`   | `#N/A`    | 函数或服务不可用。 |
|`nullReference`  | `#NULL!`  | 公式中的区域不相交。 |

下面的代码示例演示了如何创建并返回无效数字 (`#NUM!`) 错误。

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

`#VALUE!`和 `#N/A` 错误还支持自定义错误消息。 自定义错误消息显示在错误指示器菜单中，该菜单通过将鼠标悬停在包含错误的每个单元格上的错误标志上方来访问。 下面的示例演示如何返回包含错误的自定义错误消息 `#VALUE!` 。

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>使用 try-catch 块

通常情况下，使用 `try` - `catch` 自定义函数中的块捕捉出现的任何潜在错误。 如果不在代码中处理异常，它们将返回到 Excel。 默认情况下，Excel 将返回 `#VALUE!` 未处理的错误或异常。

在下面的代码示例中，自定义函数对 REST 服务执行 fetch 调用。 此调用有可能会失败，例如，如果 REST 服务返回错误或网络中断，就可能会失败。 如果发生这种情况，自定义函数将返回 `#N/A` 以指示 web 调用失败。


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
