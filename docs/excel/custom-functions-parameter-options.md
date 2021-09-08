---
ms.date: 03/08/2021
description: 了解如何在自定义函数内使用不同的参数，如Excel范围、可选参数、调用上下文等。
title: 自定义Excel选项
localization_priority: Normal
ms.openlocfilehash: a168853eeb6a81cf3d0054cb3628b609ec283af7
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938755"
---
# <a name="custom-functions-parameter-options"></a>自定义函数参数选项

自定义函数可配置许多不同的参数选项。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>可选参数

当用户在 Excel 中调用函数时，可选参数将显示在括号中。 在下面的示例中，add 函数可以选择添加第三个数字。 此函数显示为 `=CONTOSO.ADD(first, second, [third])` Excel。

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
> 如果没有为可选参数指定值，Excel会为其分配值 `null` 。 这意味着 TypeScript 中的默认初始化参数将不能正常工作。 请勿使用语法 `function add(first:number, second:number, third=0):number` ，因为它不会初始化 `third` 为 0。 请改为使用 TypeScript 语法，如上一示例所示。

定义包含一个或多个可选参数的函数时，请指定可选参数为 null 时会发生什么情况。 在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。 如果 `zipCode` 参数为 null，则默认值设置为 `98052` 。 如果 `dayOfWeek` 参数为 null，则设置为星期三。

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

## <a name="range-parameters"></a>Range 参数

自定义函数可以接受单元格数据区域作为输入参数。 函数还可以返回数据区域。 Excel将单元格数据区域作为二维数组传递。

例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。 以下函数接受 参数，JSDOC 语法在此函数的 JSON 元数据中将参数 `values` `number[][]` 的 `dimensionality` `matrix` 属性设置至 。 

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

重复参数允许用户向函数输入一系列可选参数。 调用 函数时，值在参数的数组中提供。 如果参数名称以数字结尾，则每个参数的编号将递增，例如 `ADD(number1, [number2], [number3],…)` 。 这符合用于内置函数的Excel。

以下函数对数字、单元格地址以及区域（如果输入）总计进行计算。

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

此函数 `=CONTOSO.ADD([operands], [operands]...)` 显示在工作簿Excel中。

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>重复单个值参数

重复的单值参数允许传递多个单个值。 例如，用户可以输入 ADD (1，B2，3) 。 以下示例演示如何声明单个值参数。

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

### <a name="single-range-parameter"></a>单个区域参数

从技术上说，单个范围参数不是重复参数，但此处包含此参数，因为声明与重复参数非常相似。 对于用户，它显示为 ADD (A2：B3) 其中从单个区域传递Excel。 以下示例演示如何声明单个 range 参数。

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

### <a name="repeating-range-parameter"></a>重复范围参数

重复的 range 参数允许传递多个范围或数字。 例如，用户可以输入 ADD (5，B2，C3，8，E5：E8) 。 重复区域通常使用类型指定 `number[][][]` ，因为它们是三维矩阵。 有关示例，请参阅针对重复参数 [列出的主示例](#repeating-parameters)。


### <a name="declaring-repeating-parameters"></a>声明重复参数
在 Typescript 中，指示参数是多维的。 例如，  `ADD(values: number[])` 表示一维数组，表示二维数组 `ADD(values:number[][])` ，等等。

在 JavaScript 中，对一维数组、二维数组等使用更多 `@param values {number[]}` `@param <name> {number[][]}` 维度。

对于手动创作的 JSON，请确保你的参数在 JSON 文件中指定，并检查你的参数 `"repeating": true` 是否标记为 `"dimensionality": matrix` 。

## <a name="invocation-parameter"></a>调用参数

每个自定义函数都会自动将参数作为最后一个输入参数传递，即使该参数 `invocation` 未显式声明。 此 `invocation` 参数对应于 [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) 对象。 对象可用于检索其他上下文，例如调用自定义函数的单元格 `Invocation` 的地址。 若要访问 `Invocation` 对象，必须在 `invocation` 自定义函数中声明为最后一个参数。 

> [!NOTE]
> 对于用户，此参数不会显示为自定义函数 `invocation` Excel。

以下示例演示如何使用 参数返回调用自定义函数 `invocation` 的单元格的地址。 此示例使用 [对象的 address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) `Invocation` 属性。 若要访问 `Invocation` 对象，请首先 `CustomFunctions.Invocation` 在 JSDoc 中声明为参数。 接下来， `@requiresAddress` 在 JSDoc 中声明以 `address` 访问 对象的 `Invocation` 属性。 最后，在 函数中检索并返回 `address` 属性。 

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

在Excel中，调用对象属性的自定义函数将返回遵循调用函数的单元格中的格式 `address` `Invocation` `SheetName!RelativeCellAddress` 的绝对地址。 例如，如果输入参数位于单元格 F6 中 **名为"价格** "的工作表上，则返回的参数地址值将为 `Prices!F6` 。 

`invocation`参数还可用于向用户发送Excel。 有关详细信息 [，请参阅制作流](custom-functions-web-reqs.md#make-a-streaming-function) 式处理函数。

## <a name="detect-the-address-of-a-parameter"></a>检测参数的地址

与 [调用参数结合使用](#invocation-parameter)时，可以使用 [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) 对象检索自定义函数输入参数的地址。 调用时，对象的 [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) 属性允许函数返回所有输入 `Invocation` 参数的地址。 

在输入数据类型可能不同的情况下，这非常有用。 输入参数的地址可用于检查输入值的数量格式。 然后，如有必要，可以在输入之前调整数字格式。 输入参数的地址还可用于检测输入值是否具有与后续计算相关的任何相关属性。 

>[!NOTE]
> 如果你使用手动创建的[JSON](custom-functions-json.md)元数据来返回参数地址，而不是 Yo Office 生成器，则对象必须将 属性设置为 ，并且对象必须将 属性设置为 `options` `requiresParameterAddresses` `true` `result` `dimensionality` `matrix` 。

以下自定义函数采用三个输入参数，检索每个参数的对象属性， `parameterAddresses` `Invocation` 然后返回地址。 

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

调用属性的自定义函数运行时，参数地址将按照调用函数的 `parameterAddresses` `SheetName!RelativeCellAddress` 单元格中的格式返回。 例如，如果输入参数位于单元格 D8 中 **名为"成本** "的工作表上，则返回的参数地址值将为 `Costs!D8` 。 如果自定义函数具有多个参数并返回多个参数地址，则返回的地址将溢出多个单元格，从调用该函数的单元格垂直下降。 

## <a name="next-steps"></a>后续步骤

了解如何在自定义 [函数中使用可变值](custom-functions-volatile.md)。

## <a name="see-also"></a>另请参阅

* [使用自定义函数接收和处理数据](custom-functions-web-reqs.md)
* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [手动为自定义函数创建 JSON 元数据](custom-functions-json.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
