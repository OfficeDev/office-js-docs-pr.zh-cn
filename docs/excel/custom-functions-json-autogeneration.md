---
ms.date: 07/08/2021
description: 使用 JSDoc 标记动态创建自定义函数 JSON 元数据。'
title: 为自定义函数自动生成 JSON 元数据
ms.localizationpriority: medium
ms.openlocfilehash: c14196da1d686beb32b31589187e57b2b6f2616d
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744551"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>为自定义函数自动生成 JSON 元数据

在使用 JavaScript 或 TypeScript 编写 Excel 自定义函数时，使用 [JSDoc 标记](https://jsdoc.app/)提供有关自定义函数的额外信息。 然后在生成时使用 JSDoc 标记创建 JSON 元数据文件。 使用 JSDoc 标记可让你无需手动编辑 [JSON 元数据文件](custom-functions-json.md)。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

为 JavaScript 或 TypeScript 函数添加代码注释中的 `@customfunction` 标记以将其标记为自定义函数。

可以使用 JavaScript 中的 [@param](#param) 标记或从 TypeScript 中的[函数类型](https://www.typescriptlang.org/docs/handbook/functions.html)提供函数参数类型。 有关详细信息，请参阅 [@param](#param) 标记和[类型](#types)部分。

### <a name="add-a-description-to-a-function"></a>向函数添加说明

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

以下 JSDoc 标记在自定义Excel中受支持。

* [@cancelable](#cancelable)
* [@customfunction](#customfunction) id name
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ name description
* [@requiresAddress](#requiresAddress)
* [@requiresParameterAddresses](#requiresParameterAddresses)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
<a id="cancelable"></a>
### <a name="cancelable"></a>@cancelable

指示自定义函数在函数取消时执行一个操作。

最后一个函数参数的类型必须是 `CustomFunctions.CancelableInvocation`。 函数可以将函数分配给 `oncanceled` 属性，以在函数取消时表示结果。

如果最后一个函数参数的类型为 `CustomFunctions.CancelableInvocation`，则即使标记不存在，也会被视为 `@cancelable`。

函数不能同时具有 `@cancelable` 和 `@streaming` 标记。

<a id="customfunction"></a>

### <a name="customfunction"></a>@customfunction

语法：@customfunction _id_ _name_

此标记指示 JavaScript/TypeScript 函数是一个Excel函数。 需要为自定义函数创建元数据。

下面显示了此标记的示例。

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

标识 `id` 自定义函数。

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

在用户进入函数Excel，用户会显示一个说明，并指定函数的功能。 说明不需要任何特定标记。 通过在 JSDoc 注释中添加一个短语来描述函数的功能，为自定义函数添加说明。 默认情况下，JSDoc 注释部分中未标记的任何文本都是该函数的说明。

在以下示例中，短语“对两个数字求和的函数”是 id 属性为 `ADD` 的自定义函数的相关说明。

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>
### <a name="helpurl"></a>@helpurl

语法：@helpurl _url_

提供的 _url_ 显示在 Excel 中。

在下面的示例中，为 `helpurl` `www.contoso.com/weatherhelp`。

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>
### <a name="param"></a>@param

#### <a name="javascript"></a>JavaScript

JavaScript 语法：@param {type} name _description_

* `{type}` 指定大括号中的类型信息。 有关可能使用的类型的详细信息，请参阅[类型](#types)部分。 如果未指定类型，则使用默认 `any` 类型。
* `name` 指定该标记@param的参数。 这是必需的。
* `description` 为函数参数提供显示在 Excel 中的说明。 可选。

若要将自定义函数参数表示为可选参数，请用方括号将参数名称括起。 例如，`@param {string} [text] Optional text`。

> [!NOTE]
> 可选参数的默认值为 `null`。

以下示例显示添加两个或三个数字的 ADD 函数，第三个数字作为可选参数。

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

* `name` 指定该标记@param的参数。 这是必需的。
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

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a>@requiresAddress

表示应提供计算函数所在的单元格的地址。

最后一个函数参数必须为 类型 `CustomFunctions.Invocation` 或派生类型，以使用 `@requiresAddress`。 调用函数时，`address` 属性将包含地址。

以下示例演示如何将 参数`invocation``@requiresAddress`与 结合使用以返回调用自定义函数的单元格的地址。 有关详细信息 [，请参阅调用](custom-functions-parameter-options.md#invocation-parameter) 参数。

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<a id="requiresParameterAddresses"></a>
### <a name="requiresparameteraddresses"></a>@requiresParameterAddresses

指示函数应返回输入参数的地址。 

最后一个函数参数必须为 类型 `CustomFunctions.Invocation` 或派生类型，以使用  `@requiresParameterAddresses`。 JSDoc 注释还必须包含一 `@returns` 个标记，该标记指定返回值是矩阵，如 `@returns {string[][]}` 或 `@returns {number[][]}`。 有关 [其他信息，](#matrix-type) 请参阅矩阵类型。 

调用 函数时， `parameterAddresses` 属性将包含输入参数的地址。

以下示例演示如何将 参数 `invocation` 与 结合使用 `@requiresParameterAddresses` 以返回三个输入参数的地址。 有关详细信息 [，请参阅检测参数](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) 的地址。 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array.
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<a id="returns"></a>
### <a name="returns"></a>@returns

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

<a id="streaming"></a>
### <a name="streaming"></a>@streaming

用于表示自定义函数是一个流式处理函数。 

最后一个参数的类型为 `CustomFunctions.StreamingInvocation<ResultType>`。
函数返回 `void`。

流式处理函数不会直接返回值，而是使用最后一 `setResult(result: ResultType)` 个参数调用。

由流式处理函数引发的异常将被忽略。 `setResult()` 可能称为“错误”，以指示错误结果。 有关流式处理函数的示例和更多信息，请参阅[生成流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)。

流式处理函数不能标记为 [@volatile](#volatile)。

<a id="volatile"></a>
### <a name="volatile"></a>@volatile

可变函数是指其结果不断变化的函数，即使不采用任何参数或参数未发生更改都是如此。 Excel 在每次完成计算后，都会重新计算包含可变函数和所有依赖项的单元格。 因此，过于依赖可变函数会使重新计算时间变慢，请谨慎使用。

流式处理函数不能为可变函数。

以下函数是可变函数并使用 `@volatile` 标记。

```js
/**
 * Simulates rolling a 6-sided die.
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

使用二维数组类型将参数或返回值变为值的矩阵。 例如，类型指示 `number[][]` 数字矩阵， `string[][]` 并指示字符串矩阵。

### <a name="error-type"></a>错误类型

非流式处理函数可以通过返回错误类型来指示错误。

流式处理函数可以通过使用错误类型调用 `setResult()` 来指示错误。

### <a name="promise"></a>Promise

自定义函数可以返回一个承诺，该承诺在承诺实现时提供值。 如果承诺被拒绝，则自定义函数将引发错误。

### <a name="other-types"></a>其他类型

任何其他类型都将被视为错误。

## <a name="next-steps"></a>后续步骤

了解[自定义函数的命名约定](custom-functions-naming.md)。 或者，了解如何[本地化函数](custom-functions-localize.md)，这需要你[手动编写 JSON 文件](custom-functions-json.md)。

## <a name="see-also"></a>另请参阅

* [手动为自定义函数创建 JSON 元数据](custom-functions-json.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
