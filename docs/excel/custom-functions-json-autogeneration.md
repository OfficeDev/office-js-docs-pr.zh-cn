---
ms.date: 06/27/2019
description: 使用 JSDoc 标记动态创建自定义函数 JSON 元数据。
title: 为自定义函数自动生成 JSON 元数据
localization_priority: Priority
ms.openlocfilehash: 1230e1bfdeead306531a218373c2756b29fa4abe
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454655"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>为自定义函数自动生成 JSON 元数据

在 JavaScript 或 TypeScript 中写入 Excel 自定义函数时，使用 JSDoc 标记提供有关自定义函数的额外信息。 然后在生成时使用 JSDoc 标记创建 [JSON 元数据文件](custom-functions-json.md)。 使用 JSDoc 标记使您免除手动编辑 JSON 元数据文件的工作。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

为 JavaScript 或 TypeScript 函数添加代码注释中的 `@customfunction` 标记以将其标记为自定义函数。

可以使用 JavaScript 中的 [@param](#param) 标记或从 TypeScript 中的[函数类型](https://www.typescriptlang.org/docs/handbook/functions.html)提供函数参数类型。 有关详细信息，请参阅 [@param](#param) 标记和[类型](#types)部分。

### <a name="adding-a-description-to-a-function"></a>为函数添加说明

当用户需要帮助来了解自定义函数的功能时，将向用户显示用作帮助文本的说明。 说明不需要任何特定标记。 只需在 JSDoc 注释中输入简短的文本说明即可。 一般来说，说明位于 JSDoc 注释部分的开头，但无论位于何处，它都有用。

若要查看内置函数说明的示例，请打开 Excel，转到“**公式**”选项卡，然后选择“**插入函数**”。 然后，你可以浏览所有函数说明，还可以查看列出的自定义函数。

在以下示例中，短语“计算球体的体积” 就是自定义函数的相关说明。

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


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

下面的示例显示了声明自定义函数的最简单方法。

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

`id` 是自定义函数的固定标识符。

* 如果未提供 `id`，请将 JavaScript/TypeScript 函数名称转换为大写并删除禁用字符。
* `id` 对于所有自定义函数必须是唯一的。
* 允许使用的字符限为：A-Z、a-z、0-9、下划线 (\_) 和句点 (.)。

在下面的示例中，增量是函数的 `id` 和 `name`。

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a>name

提供自定义函数的显示`name`。

* 如果未提供名称，则 id 还会用作名称。
* 允许使用的字符：字母 [Unicode 字母字符](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、句点 (.) 和下划线 (\_)。
* 必须以字母开头。
* 最大长度为 128 个字符。

在下面的示例中，INC 是函数的 `id`，并且 `increment` 是 `name`。

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a>说明

说明不需要任何特定标记。 通过在 JSDoc 注释中添加一个短语来描述函数的功能，为自定义函数添加说明。 默认情况下，JSDoc 注释部分中未标记的任何文本都是该函数的说明。 当 Excel 中的用户进入该函数时，将向其显示相关说明。 在以下示例中，短语“对两个数字求和的函数”是 id 属性为 `ADD` 的自定义函数的相关说明。

在下面的示例中，ADD 是函数的 `id` 和 `name`，并且提供了说明。

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

语法：@helpurl _url_

提供的 _url_ 显示在 Excel 中。

在下面的示例中，`helpurl` 是 www.contoso.com/weatherhelp。

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a>JavaScript

JavaScript 语法：@param {type} name _description_

* `{type}` 应在大括号内指定类型信息。 有关可能使用的类型的详细信息，请参阅[类型](#types)部分。 可选：如果未指定，则使用类型 `any`。
* `name` 指定 @param 标记适用于哪个参数。 必需。
* `description` 为函数参数提供显示在 Excel 中的说明。 可选。

若要将自定义函数参数表示为可选，请执行以下操作：
* 为参数名称加上方括号。 例如：`@param {string} [text] Optional text`。

> [!NOTE]
> 可选参数的默认值为 `null`。

下面的示例显示了 ADD 函数，该函数将两个或三个数字相加，第三个数字作为可选参数。

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a>TypeScript

TypeScript 语法：@param name _description_

* `name` 指定 @param 标记适用于哪个参数。 必需。
* `description` 为函数参数提供显示在 Excel 中的说明。 可选。

有关可能使用的函数参数类型的详细信息，请参阅[类型](#types)部分。

若要将自定义函数参数表示为可选，请执行以下操作之一：
* 使用可选参数。 例如：`function f(text?: string)`
* 为该参数提供默认值。 例如：`function f(text: string = "abc")`

有关 @param 的详细说明，请参阅：[JSDoc](https://jsdoc.app/tags-param.html)

> [!NOTE]
> 可选参数的默认值为 `null`。

下面的示例显示了将两个数字相加的 `add` 函数。

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

表示应提供计算函数所在的单元格的地址。

最后一个函数参数的类型必须是 `CustomFunctions.Invocation` 或派生类型。 调用函数时，`address` 属性将包含地址。 有关使用 `@requiresAddress` 标记的函数示例，请参阅[寻址单元格的上下文参数](./custom-functions-parameter-options.md#addressing-cells-context-parameter)。

---
### <a name="returns"></a>@returns
<a id="returns"/>

语法：@returns {_type_}

提供返回值的类型。

如果省略 `{type}`，则将使用 TypeScript 类型信息。 如果没有类型信息，则类型将为 `any`。

下面的示例显示了使用 `@returns` 标记的 `add` 函数。

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

用于表示自定义函数是一个流式处理函数。 

最后一个参数的类型应为 `CustomFunctions.StreamingInvocation<ResultType>`。
该函数应返回 `void`。

流式处理函数不直接返回值，而是应该使用最后一个参数调用 `setResult(result: ResultType)`。

由流式处理函数引发的异常将被忽略。 `setResult()` 可能称为“错误”，以指示错误结果。 有关流式处理函数的示例和更多信息，请参阅[生成流式处理函数](./custom-functions-web-reqs.md#make-a-streaming-function)。

流式处理函数不能标记为 [@volatile](#volatile)。

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

可变函数是指其结果不断变化的函数，即使不采用任何参数或参数未发生更改都是如此。 Excel 在每次完成计算后，都会重新计算包含可变函数和所有依赖项的单元格。 因此，过于依赖可变函数会使重新计算时间变慢，请谨慎使用。

流式处理函数不能为可变函数。

以下函数是可变函数并使用 `@volatile` 标记。

```js
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a>类型

通过指定参数类型，Excel 会在调用函数之前将值转换为该类型。 如果类型为 `any`，则不会执行任何转换。

### <a name="value-types"></a>值类型

可以使用以下类型之一表示单个值：`boolean`、`number`、`string`。

### <a name="matrix-type"></a>矩阵类型

使用二维数组类型将参数或返回值变为值的矩阵。 例如，类型 `number[][]` 表示数字的矩阵。 `string[][]` 表示字符串的矩阵。 

### <a name="error-type"></a>错误类型

非流式处理函数可以通过返回错误类型来指示错误。

流式处理函数可以通过使用错误类型调用 `setResult()` 来指示错误。

### <a name="promise"></a>Promise

函数可以返回 Promise，将在解析 promise 后提供值。 如果 promise 被拒绝，则会出现错误。

### <a name="other-types"></a>其他类型

任何其他类型都将被视为错误。

## <a name="next-steps"></a>后续步骤
了解[自定义函数的命名约定](custom-functions-naming.md)。 或者，了解如何[本地化函数](custom-functions-localize.md)，这需要你[手动编写 JSON 文件](custom-functions-json.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
