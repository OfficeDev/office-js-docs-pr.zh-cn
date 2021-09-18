---
title: 处理和返回自定义函数中的错误
description: '处理和返回自定义函数中类似 #NULL! 来自自定义函数。'
ms.date: 08/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: a2f93059f9082bc5a53c07159c9356a41cf16729
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/18/2021
ms.locfileid: "59443543"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>处理和返回自定义函数中的错误

如果自定义函数运行时出错，则返回错误以通知用户。 如果你有特定的参数要求（例如仅正数），请测试参数，如果它们不正确，则引发错误。 您还可以使用 块捕获自定义函数运行时出现 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) 的任何错误。

## <a name="detect-and-throw-an-error"></a>检测和引发错误

让我们看一下需要确保邮政编码参数格式正确的情况，自定义函数才能正常工作。 下面的自定义函数使用正则表达式来检查邮政编码。 如果邮政编码格式正确，它将使用另一个函数查找城市并返回值。 如果格式无效，函数会向 `#VALUE!` 单元格返回错误。

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

[CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error)对象用于将错误返回给单元格。 创建对象时，通过选择以下枚举值之一来指定 `ErrorCode` 要使用哪个错误。

|ErrorCode 枚举值  |Excel 单元格值  |说明  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | 函数试图除以零。 |
|`invalidName`    | `#NAME?`  | 函数名称有一个拼写错误。 请注意，此错误作为自定义函数输入错误受到支持，但不作为自定义函数输出错误。 |
|`invalidNumber`  | `#NUM!`   | 公式中的数字存在问题。 |
|`invalidReference` | `#REF!` | 函数引用无效的单元格。 请注意，此错误作为自定义函数输入错误受到支持，但不作为自定义函数输出错误。|
|`invalidValue`   | `#VALUE!` | 公式中的值的类型错误。 |
|`notAvailable`   | `#N/A`    | 函数或服务不可用。 |
|`nullReference`  | `#NULL!`  | 公式中的区域不相交。 |

下面的代码示例演示了如何创建并返回无效数字 (`#NUM!`) 错误。

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

和 `#VALUE!` `#N/A` 错误还支持自定义错误消息。 自定义错误消息显示在错误指示器菜单中，通过将鼠标悬停在出错的每个单元格上的错误标志上来访问该菜单。 以下示例演示如何返回包含错误的自定义 `#VALUE!` 错误消息。

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### <a name="handle-errors-when-working-with-dynamic-arrays"></a>使用动态数组时处理错误

除了返回单个错误之外，自定义函数还可以输出包含错误的动态数组。 例如，自定义函数可以输出数组 `[1],[#NUM!],[3]` 。 下面的代码示例演示如何向自定义函数中输入三个参数，将其中一个输入参数替换为错误，然后返回一个二维数组，并返回处理每个输入参数 `#NUM!` 的结果。

```js
/**
* Returns the #NUM! error as part of a 2-dimensional array.
* @customfunction
* @param {number} first First parameter.
* @param {number} second Second parameter.
* @param {number} third Third parameter.
* @returns {number[][]} Three results, as a 2-dimensional array.
*/
function returnInvalidNumberError(first, second, third) {
  // Use the `CustomFunctions.Error` object to retrieve an invalid number error.
  var error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  var firstResult = first;
  var secondResult =  error;
  var thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### <a name="errors-as-custom-function-inputs"></a>作为自定义函数输入的错误

即使输入区域包含错误，自定义函数也可以计算。 例如，自定义函数可以将 **区域 A2：A7** 作为输入，即使 **A6：A7** 包含错误。

若要处理包含错误的输入，自定义函数必须将 JSON 元数据属性 `allowErrorForDataTypeAny` 设置为 `true` 。 有关详细信息 [，请参阅手动为自定义函数创建 JSON](custom-functions-json.md#metadata-reference) 元数据。

> [!IMPORTANT]
> `allowErrorForDataTypeAny`属性只能与手动创建的[JSON 元数据一同使用](custom-functions-json.md)。 此属性不能用于自动生成的 JSON 元数据进程。

## <a name="use-trycatch-blocks"></a>使用 `try...catch` 块

通常，使用 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) 自定义函数中的块来捕获发生的任何潜在错误。 如果不处理代码中的异常，它们将被返回到Excel。 默认情况下，Excel `#VALUE!` 错误或异常返回。

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
* [自定义函数要求集](../reference/requirement-sets/custom-functions-requirement-sets.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
