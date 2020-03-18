---
ms.date: 07/15/2019
description: 了解如何在自定义函数中使用不同的参数，例如 Excel 范围、可选参数、调用上下文等。
title: Excel 自定义函数的选项
localization_priority: Normal
ms.openlocfilehash: 66e873117b82ed7258b5965a6e964f4b9e01df21
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719481"
---
# <a name="custom-functions-parameter-options"></a>自定义函数参数选项

自定义函数可通过多个不同的参数选项进行配置。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>可选参数

而常规参数是必需的，而可选参数则不是。 当用户在 Excel 中调用函数时，可选参数将显示在括号中。 在下面的示例中，add 函数可以选择添加第三个数字。 在 Excel 中， `=CONTOSO.ADD(first, second, [third])`此函数显示为。

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
> 如果没有为可选参数指定任何值，则 Excel 会为其分配`null`值。 这意味着 TypeScript 中的默认初始化参数不会按预期工作。 因此，请不要使用语法`function add(first:number, second:number, third=0):number` ，因为它不会`third`初始化为0。 而是使用上一示例中所示的 TypeScript 语法。

在定义包含一个或多个可选参数的函数时，应指定可选参数为 null 时所发生的情况。 在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。 如果`zipCode`参数为 null，则默认值设置为`98052`。 如果`dayOfWeek`参数为 null，则将其设置为星期三。

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

例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。 下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。 请注意，在此函数的 JSON 元数据中，该`type`参数的属性设置`matrix`为。

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

重复参数允许用户在函数中输入一系列可选的参数。 调用函数时，将在参数的数组中提供值。 如果参数名称以数字结尾，则每个参数将递增该数，例如`ADD(number1, [number2], [number3],…)`。 这与用于内置 Excel 函数的约定相匹配。

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

此函数显示`=CONTOSO.ADD([operands], [operands]...)`在 Excel 工作簿中。

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>重复单个值参数

一个重复的单值参数允许传递多个单个值。 例如，用户可以输入 ADD （1，B2，3）。 下面的示例演示如何声明单个值参数。

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

从技术上讲，单个 range 参数并不是重复参数，但此处包含此参数，这是因为声明与重复参数非常相似。 它向用户显示为 "添加" （A2： B3），其中单个范围是从 Excel 中传递的。 下面的示例展示了如何声明一个 range 参数。

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

重复区域参数允许传递多个区域或数字。 例如，用户可以输入 ADD （5，B2，C3，8，E5： E8）。 重复区域通常是使用类型为三维`number[][][]`矩阵的类型指定的。 有关示例，请参阅为重复参数列出的主示例（#repeating 参数）。


### <a name="declaring-repeating-parameters"></a>声明重复参数
在 Typescript 中，指示参数是多维的。 例如， `ADD(values: number[])`将指示一维数组， `ADD(values:number[][])`指示二维数组，依此类推。

在 JavaScript 中， `@param values {number[]}`对于二维数组使用一维`@param <name> {number[][]}`数组，对更多维度使用。

对于 "手动创作的 JSON"，请确保在 JSON `"repeating": true`文件中将参数指定为，并检查参数是否标记为`"dimensionality": matrix`。

>[!NOTE]
>包含重复参数的函数将自动包含一个调用参数作为最后一个参数。 有关调用参数的详细信息，请参阅下一节。

## <a name="invocation-parameter"></a>调用参数

每个自定义函数自动传递`invocation`一个参数作为最后一个参数。 此参数可用于检索其他上下文，如调用单元格的地址。 也可以用于向 Excel 发送信息，例如用于[取消函数](custom-functions-web-reqs.md#make-a-streaming-function)的函数处理程序。 即使不声明参数，您的自定义函数也有此参数。 在 Excel 中，用户不会看到此参数。 如果要在自定义`invocation`函数中使用，则将其声明为最后一个参数。

在下面的代码示例中， `invocation`将显式声明上下文以供参考。

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

参数允许您获取调用单元格的上下文，这在某些方案中非常有用，包括[查找调用自定义函数的单元格的地址](#addressing-cells-context-parameter)。

### <a name="addressing-cells-context-parameter"></a>寻址单元格的上下文参数

在某些情况下，您需要获取调用自定义函数的单元格的地址。 这在以下情况下很有用：

- 格式区域：将单元格的地址用作存储[OfficeRuntime](../excel/custom-functions-runtime.md#storing-and-accessing-data)中的信息的密钥。 然后，使用 Excel 中的 [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) 从 `OfficeRuntime.storage` 加载该键。
- 显示缓存值：如果脱机使用函数，将显示 `OfficeRuntime.storage` 中使用 `onCalculated` 存储的缓存值。
- 协调：使用单元格地址发现原始单元格，以帮助你在处理时进行协调。

若要在函数中请求寻址单元格的上下文，您需要使用函数来查找单元格的地址，如以下示例所示。 仅当`@requiresAddress`在函数的注释中对单元格地址进行了标记时，才会公开该单元格地址的相关信息。

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
```

默认情况下，从 `getAddress` 函数返回的值遵循以下格式：`SheetName!CellNumber`。 例如，如果名为“Expense”的工作表中的 B2 单元格调用了函数，则返回的值为 `Expenses!B2`。

## <a name="next-steps"></a>后续步骤

了解如何[在自定义函数中保存状态](custom-functions-save-state.md)，或[在自定义函数中使用可变值](custom-functions-volatile.md)。

## <a name="see-also"></a>另请参阅

* [使用自定义函数接收和处理数据](custom-functions-web-reqs.md)
* [自定义函数元数据](custom-functions-json.md)
* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)