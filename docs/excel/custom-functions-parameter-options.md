---
title: Excel 自定义函数的选项
description: 了解如何在自定义函数中使用不同的参数，例如 Excel 范围、可选参数、调用上下文等。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: de86afc60d7d0b81820bd742e989e0ee7dd6970c
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958571"
---
# <a name="custom-functions-parameter-options"></a>自定义函数参数选项

可使用许多不同的参数选项配置自定义函数。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>可选参数

当用户在 Excel 中调用函数时，可选参数将显示在括号中。 在以下示例中，添加函数可以选择添加第三个数字。 此函数在 Excel 中显示。`=CONTOSO.ADD(first, second, [third])`

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
> 如果未为可选参数指定任何值，Excel 将为其分配值 `null`。 这意味着 TypeScript 中的默认初始化参数将无法按预期工作。 不要使用语 `function add(first:number, second:number, third=0):number` 法，因为它不会初始化 `third` 为 0。 请改用 TypeScript 语法，如上一示例所示。

定义包含一个或多个可选参数的函数时，指定可选参数为 null 时会发生什么情况。 在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。 `zipCode`如果参数为 null，则默认值设置为 `98052`。 `dayOfWeek`如果参数为 null，则设置为星期三。

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

自定义函数可以接受一系列单元格数据作为输入参数。 函数还可以返回一系列数据。 Excel 将以二维数组的形式传递一系列单元格数据。

例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。 以下函数接受参数`values`，JSDOC 语`number[][]`法将参数的属性`matrix`设置为此函数的 `dimensionality` JSON 元数据中。

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
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

重复参数允许用户向函数输入一系列可选参数。 调用函数时，这些值在参数的数组中提供。 如果参数名称以数字结尾，则每个参数的数目将递增，例如 `ADD(number1, [number2], [number3],…)`。 这符合用于内置 Excel 函数的约定。

以下函数对输入的数字、单元格地址以及范围的总数求和。

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

![正在输入到 Excel 工作表单元格中的 ADD 自定义函数](../images/operands.png)

### <a name="repeating-single-value-parameter"></a>重复单值参数

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

### <a name="single-range-parameter"></a>单范围参数

单个范围参数在技术上不是重复参数，但在此处包含，因为声明非常类似于重复参数。 在用户看来，它显示为 ADD (A2：B3) ，其中单个范围从 Excel 传递。 以下示例演示如何声明单个范围参数。

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

重复范围参数允许传递多个范围或数字。 例如，用户可以输入 ADD (5，B2，C3，8，E5：E8) 。 重复范围通常与类型 `number[][][]` 一起指定，因为它们是三维矩阵。 有关示例，请参阅为 [重复参数](#repeating-parameters)列出的主要示例。

### <a name="declaring-repeating-parameters"></a>声明重复参数

在 Typescript 中，指示参数是多维的。 例如，  `ADD(values: number[])` 将指示一维数组、 `ADD(values:number[][])` 指示二维数组等。

在 JavaScript 中，用于 `@param values {number[]}` 一维数组、 `@param <name> {number[][]}` 二维数组等，用于更多维度。

对于手写的 JSON，请确保参数在 `"repeating": true` JSON 文件中指定，并检查参数是否标记为 `"dimensionality": matrix`。

## <a name="invocation-parameter"></a>调用参数

每个自定义函数都会自动将参数作为最后一个 `invocation` 输入参数传递，即使该参数未显式声明。 此 `invocation` 参数对应于 [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) 对象。 该 `Invocation` 对象可用于检索其他上下文，例如调用自定义函数的单元格的地址。 若要访问对象 `Invocation` ，必须声明 `invocation` 为自定义函数中的最后一个参数。

> [!NOTE]
> 该 `invocation` 参数不会显示为 Excel 中用户的自定义函数参数。

以下示例演示如何使用 `invocation` 参数返回调用自定义函数的单元格的地址。 此示例使用对象的`Invocation`[地址](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-address-member)属性。 若要访问对象 `Invocation` ，请先在 JSDoc 中声明 `CustomFunctions.Invocation` 为参数。 接下来，在 JSDoc 中声明 `@requiresAddress` 以访问 `address` 对象的 `Invocation` 属性。 最后，在函数中检索并返回属性 `address` 。

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
  const address = invocation.address;
  return address;
}
```

在 Excel 中，调用 `address` 对象属性的 `Invocation` 自定义函数将按照调用该函数的单元格中的格式 `SheetName!RelativeCellAddress` 返回绝对地址。 例如，如果输入参数位于单元格 F6 中名为“ **价格** ”的工作表上，则返回的参数地址值将为 `Prices!F6`。

该 `invocation` 参数还可用于将信息发送到 Excel。 有关详细信息，请参阅 [“创建流式处理”函](custom-functions-web-reqs.md#make-a-streaming-function) 数。

## <a name="detect-the-address-of-a-parameter"></a>检测参数的地址

结合 [调用参数](#invocation-parameter)，可以使用 [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) 对象检索自定义函数输入参数的地址。 调用时，对象的 `Invocation` [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-parameteraddresses-member) 属性允许函数返回所有输入参数的地址。

在输入数据类型可能有所不同的情况下，这很有用。 输入参数的地址可用于检查输入值的数字格式。 然后，可以根据需要在输入之前调整数字格式。 输入参数的地址还可用于检测输入值是否具有与后续计算相关的任何相关属性。

>[!NOTE]
> 如果使用 [手动创建的 JSON 元数据](custom-functions-json.md) 来返回参数地址，而不是 [Office 外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md)， `options` 则该对象必须将 `requiresParameterAddresses` 属性设置 `true`为，并且 `result` 该对象必须将 `dimensionality` 属性设置为 `matrix`。

以下自定义函数采用三个输入参数，检索 `parameterAddresses` 每个参数的 `Invocation` 对象属性，然后返回地址。

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
  const addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

调用属性的 `parameterAddresses` 自定义函数运行时，参数地址将按照调用该函数的单元格中的格式 `SheetName!RelativeCellAddress` 返回。 例如，如果输入参数位于单元格 D8 中名为 **Costs** 的工作表上，则返回的参数地址值将为 `Costs!D8`。 如果自定义函数具有多个参数，并且返回了多个参数地址，则返回的地址将溢出多个单元格，从调用该函数的单元格垂直下降。

## <a name="next-steps"></a>后续步骤

了解如何 [在自定义函数中使用易失性值](custom-functions-volatile.md)。

## <a name="see-also"></a>另请参阅

- [使用自定义函数接收和处理数据](custom-functions-web-reqs.md)
- [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
- [为自定义函数手动创建 JSON 元数据](custom-functions-json.md)
- [在 Excel 中创建自定义函数](custom-functions-overview.md)
- [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
