---
ms.date: 04/03/2019
description: 使用 JSDOC 标记动态创建自定义函数 JSON 元数据。
title: 创建自定义函数的 JSON 元数据（预览）
localization_priority: Priority
ms.openlocfilehash: c6d89684da2d0773ccfb1763e5e3e426e647523b
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478951"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a>创建自定义函数的 JSON 元数据（预览）

在 JavaScript 或 TypeScript 中写入 Excel 自定义函数时，使用 JSDoc 标记提供有关自定义函数的额外信息。 然后在生成时使用 JSDoc 标记创建 [JSON 元数据文件](custom-functions-json.md)。 使用 JSDoc 标记使您免除手动编辑 JSON 元数据文件的工作。

为 JavaScript 或 TypeScript 函数添加代码注释中的 `@customfunction` 标记以将其标记为自定义函数。

可以使用 JavaScript 中的 [@param](#param) 标记或从 TypeScript 中的[函数类型](http://www.typescriptlang.org/docs/handbook/functions.html)提供函数参数类型。 有关详细信息，请参阅 [@param](#param) 标记和[类型](#Types)部分。

## <a name="jsdoc-tags"></a>JSDoc 标记
Excel 自定义函数支持以下 JSDoc 标记：
* [@cancelable](#cancelable)
* [@customfunction](#customfunction) id name
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ name description
* [@requiresAddress](#requiresAddress)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

表示自定义函数希望在取消函数时执行操作。

最后一个函数参数的类型必须是 `CustomFunctions.CancelableInvocation`。 该函数可以将函数分配给 `oncanceled` 属性来表示在取消函数时要执行的操作。

如果最后一个函数参数的类型为 `CustomFunctions.CancelableInvocation`，则即使标记不存在，也会被视为 `@cancelable`。

函数不能同时具有 `@cancelable` 和 `@streaming` 标记。

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

语法：@customfunction _id_ _name_

指定此标记以将 JavaScript/TypeScript 函数视为 Excel 自定义函数。

需要此标记才能创建自定义函数的元数据。

还应调用 `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a>id 

id 用作存储在文档中的自定义函数的固定标识符。 不得更改。

* 如果未提供 id，JavaScript/TypeScript 函数名称将转换为大写形式，并删除不允许使用的字符。
* id 对于所有自定义函数必须是唯一的。
* 允许使用的字符仅限于：A-Z、a-z、0-9 和句点 (.)。

#### <a name="name"></a>name

提供自定义函数的显示名称。 

* 如果未提供名称，则 id 还会用作名称。
* 允许使用的字符：字母 [Unicode 字母字符](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、句点 (.) 和下划线 (\_)。
* 必须以字母开头。
* 最大长度为 128 个字符。

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

语法：@helpurl _url_

提供的 _url_ 显示在 Excel 中。

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a>JavaScript

JavaScript 语法：@param {type} name _description_

* `{type}` 应在大括号内指定类型信息。 有关可能使用的类型的详细信息，请参阅[类型](##types)。 可选：如果未指定，则使用类型 `any`。
* `name` 指定 @param 标记应用于哪个参数。 必需。
* `description` 为函数参数提供显示在 Excel 中的说明。 可选。

若要将自定义函数参数表示为可选，请执行以下操作：
* 为参数名称加上方括号。 例如：`@param {string} [text] Optional text`。

#### <a name="typescript"></a>TypeScript

TypeScript 语法：@param name _description_

* `name` 指定 @param 标记适用于哪个参数。 必需。
* `description` 为函数参数提供显示在 Excel 中的说明。 可选。

有关可能使用的函数参数类型的详细信息，请参阅[类型](##types)。

若要将自定义函数参数表示为可选，请执行以下操作之一：
* 使用可选参数。 例如： `function f(text?: string)`
* 为该参数提供默认值。 例如： `function f(text: string = "abc")`

有关 @param 的详细说明，请参阅：[JSDoc](http://usejsdoc.org/tags-param.html)

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

表示应提供计算函数所在的单元格的地址。 

最后一个函数参数的类型必须是 `CustomFunctions.Invocation` 或派生类型。 调用函数时，`address` 属性将包含地址。

---
### <a name="returns"></a>@returns
<a id="returns"/>

语法：@returns {_type_}

提供返回值的类型。

如果省略 `{type}`，则将使用 TypeScript 类型信息。 如果没有类型信息，则类型将为 `any`。

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

用于表示自定义函数是一个流式处理函数。 

最后一个参数的类型应为 `CustomFunctions.StreamingInvocation<ResultType>`。
该函数应返回 `void`。

流式处理函数不直接返回值，而是应该使用最后一个参数调用 `setResult(result: ResultType)`。

由流式处理函数引发的异常将被忽略。 `setResult()` 可能称为“错误”，以指示错误结果。

流式处理函数不能标记为 [@volatile](#volatile)。

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

可变函数是其结果不能假定为即使不采用任何参数或参数未发生更改也始终保持不变的函数。 Excel 在每次完成计算后，都会重新计算包含可变函数和所有依赖项的单元格。 因此，过于依赖可变函数会使重新计算时间变慢，请谨慎使用。

流式处理函数不能为可变函数。

---

## <a name="types"></a>类型

通过指定参数类型，Excel 会在调用函数之前将值转换为该类型。 如果类型为 `any`，则不会执行任何转换。

### <a name="value-types"></a>值类型

可以使用以下类型之一表示单个值：`boolean`、`number`、`string`。

### <a name="matrix-type"></a>矩阵类型

使用二维数组类型将参数或返回值变为值的矩阵。 例如，类型 `number[][]` 表示数字的矩阵。 `string[][]` 表示字符串的矩阵。 

### <a name="error-type"></a>错误类型

非流式处理函数可以通过返回错误类型来指示错误。

流式处理函数可以通过使用错误类型调用 setResult() 来指示错误。

### <a name="promise"></a>Promise

函数可以返回 Promise，将在解析 promise 后提供值。 如果 promise 被拒绝，则会出现错误。

### <a name="other-types"></a>其他类型

任何其他类型都将被视为错误。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
* [自定义函数调试](custom-functions-debugging.md)
