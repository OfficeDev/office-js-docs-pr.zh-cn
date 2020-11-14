---
ms.date: 11/06/2020
description: 了解如何在自定义函数中使用不同的参数，例如 Excel 范围、可选参数、调用上下文等。
title: Excel 自定义函数的选项
localization_priority: Normal
ms.openlocfilehash: 0a803a4d41354530584b25d2bf9df944af430909
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071618"
---
# <a name="custom-functions-parameter-options"></a>自定义函数参数选项

可以使用许多不同的参数选项配置自定义函数。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>可选参数

当用户在 Excel 中调用函数时，可选参数将显示在括号中。 在下面的示例中，add 函数可以选择添加第三个数字。 在 Excel 中，此函数显示为 `=CONTOSO.ADD(first, second, [third])` 。

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

---

> [!NOTE]
> 如果没有为可选参数指定任何值，则 Excel 会为其分配值 `null` 。 这意味着 TypeScript 中的默认初始化参数不会按预期工作。 请勿使用语法， `function add(first:number, second:number, third=0):number` 因为它不会初始化 `third` 为0。 而是使用上一示例中所示的 TypeScript 语法。

在定义包含一个或多个可选参数的函数时，请指定可选参数为 null 时将发生的情况。 在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。 如果 `zipCode` 参数为 null，则默认值设置为 `98052` 。 如果 `dayOfWeek` 参数为 null，则将其设置为星期三。

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a>范围参数

您的自定义函数可能接受作为输入参数的单元格数据的范围。 函数还可以返回数据区域。 Excel 将一个区域的单元格数据作为二维数组进行传递。

例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。 下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。 请注意，在此函数的 JSON 元数据中，该参数的 `type` 属性设置为 `matrix` 。

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a>重复参数

重复参数允许用户输入函数的一系列可选参数。 调用函数时，将在参数的数组中提供值。 如果参数名称以数字结尾，则每个参数的数目都将以增量方式增加，例如 `ADD(number1, [number2], [number3],…)` 。 这与用于内置 Excel 函数的约定相匹配。

下面的函数汇总了数字、单元格地址和区域的总和（如果已输入）。

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

此函数显示 `=CONTOSO.ADD([operands], [operands]...)` 在 Excel 工作簿中。

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>重复单个值参数

一个重复的单值参数允许传递多个单个值。 例如，用户可以输入 ADD (1，B2，3) 。 下面的示例演示如何声明单个值参数。

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a>单个范围参数

从技术上讲，单个 range 参数不是重复参数，但此处包含此参数，这是因为声明与重复参数非常相似。 在从 Excel 中传递单个范围的情况下，会向用户显示 "添加 (A2： B3) "。 下面的示例展示了如何声明一个 range 参数。

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a>重复区域参数

重复区域参数允许传递多个区域或数字。 例如，用户可以输入 ADD (5，B2，C3，8，E5： E8) 。 重复区域通常是使用类型为三维矩阵的类型指定的 `number[][][]` 。 有关示例，请参阅为重复参数列出的主要示例 ( # # 重复参数) 。


### <a name="declaring-repeating-parameters"></a>声明重复参数
在 Typescript 中，指示参数是多维的。 例如，  `ADD(values: number[])` 将指示一维数组， `ADD(values:number[][])` 指示二维数组，依此类推。

在 JavaScript 中，对于二维数组使用 `@param values {number[]}` 一维数组，对 `@param <name> {number[][]}` 更多维度使用。

对于 "手动创作的 JSON"，请确保 `"repeating": true` 在 json 文件中将参数指定为，并检查参数是否标记为 `"dimensionality": matrix` 。

## <a name="invocation-parameter"></a>调用参数

每个自定义函数自动传递一个 `invocation` 参数作为最后一个参数。 此参数可用于检索其他上下文，如调用单元格的地址。 也可以用于向 Excel 发送信息，例如用于 [取消函数](custom-functions-web-reqs.md#make-a-streaming-function)的函数处理程序。 即使不声明参数，您的自定义函数也有此参数。 在 Excel 中，用户不会看到此参数。 如果要 `invocation` 在自定义函数中使用，则将其声明为最后一个参数。

在下面的代码示例中，将 `invocation` 显式声明上下文以供参考。

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
```

## <a name="next-steps"></a>后续步骤

了解如何 [在自定义函数中使用可变值](custom-functions-volatile.md)。

## <a name="see-also"></a>另请参阅

* [使用自定义函数接收和处理数据](custom-functions-web-reqs.md)
* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [手动创建自定义函数的 JSON 元数据](custom-functions-json.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
