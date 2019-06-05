---
ms.date: 05/30/2019
description: 了解如何在自定义函数中使用不同的参数, 例如 Excel 范围、可选参数、调用上下文等。
title: Excel 自定义函数的选项
localization_priority: Normal
ms.openlocfilehash: 7bc907157810ce88330fe41b21ca6ff115525491
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706055"
---
# <a name="custom-functions-parameter-options"></a>自定义函数参数选项

自定义函数可通过多个不同的参数选项进行配置:
- [可选参数](#custom-functions-optional-parameters)
- [范围参数](#range-parameters)
- [调用上下文参数](#invocation-parameter)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-optional-parameters"></a>自定义函数可选参数

而常规参数是必需的, 而可选参数则不是。 当用户在 Excel 中调用函数时，可选参数将显示在括号中。 在下面的示例中, add 函数可以选择添加第三个数字。 在 Excel 中, `=CONTOSO.ADD(first, second, [third])`此函数显示为。

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third !== undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

定义包含一个或多个可选参数的函数时，应指定未定义可选参数时会发生什么情况。 在以下示例中，`zipCode` 和 `dayOfWeek` 都是 `getWeatherReport` 函数的可选参数。 如果未`zipCode`定义此参数, 则默认值将设置为`98052`。 如果未定义 `dayOfWeek` 参数，则会将其设置为星期三。

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code. If omitted, zipCode = 98052.
 * @param {string} dayOfWeek Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## <a name="range-parameters"></a>范围参数

您的自定义函数可能接受作为输入参数的单元格数据的范围。 函数还可以返回数据区域。 Excel 将一个区域的单元格数据作为二维数组进行传递。

例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。 下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。 请注意, 在此函数的 JSON 元数据中, 该`type`参数的属性设置`matrix`为。

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## <a name="invocation-parameter"></a>调用参数

每个自定义函数自动传递`invocation`一个参数作为最后一个参数。 此参数可用于检索其他上下文, 如调用单元格的地址。 也可以用于向 Excel 发送信息, 例如用于[取消函数](custom-functions-web-reqs.md#make-a-streaming-function)的函数处理程序。 即使不声明参数, 您的自定义函数也有此参数。 在 Excel 中, 用户不会看到此参数。 如果要在自定义`invocation`函数中使用, 则将其声明为最后一个参数。

在下面的代码示例中, `invocation`将显式声明上下文以供参考。

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
CustomFunctions.associate("ADD", add);
```

参数允许您获取调用单元格的上下文, 这在某些方案中非常有用, 包括[查找调用自定义函数的单元格的地址](#addressing-cells-context-parameter)。

### <a name="addressing-cells-context-parameter"></a>寻址单元格的上下文参数

在某些情况下, 您需要获取调用自定义函数的单元格的地址。 这在以下情况下很有用:

- 格式区域: 将单元格的地址用作存储[OfficeRuntime](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data)中的信息的密钥。 然后，使用 Excel 中的 [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) 从 `OfficeRuntime.storage` 加载该键。
- 显示缓存值：如果脱机使用函数，将显示 `OfficeRuntime.storage` 中使用 `onCalculated` 存储的缓存值。
- 协调：使用单元格地址发现原始单元格，以帮助你在处理时进行协调。

若要在函数中请求寻址单元格的上下文, 您需要使用函数来查找单元格的地址, 如以下示例所示。 仅当`@requiresAddress`在函数的注释中对单元格地址进行了标记时, 才会公开该单元格地址的相关信息。

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
CustomFunctions.associate("GETADDRESS", getAddress);
```

默认情况下，从 `getAddress` 函数返回的值遵循以下格式：`SheetName!CellNumber`。 例如，如果名为“Expense”的工作表中的 B2 单元格调用了函数，则返回的值为 `Expenses!B2`。

## <a name="next-steps"></a>后续步骤
了解如何[在自定义函数中保存状态](custom-functions-save-state.md), 或[在自定义函数中使用可变值](custom-functions-volatile.md)。

## <a name="see-also"></a>另请参阅

* [使用自定义函数接收和处理数据](custom-functions-web-reqs.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数元数据](custom-functions-json.md)
* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
