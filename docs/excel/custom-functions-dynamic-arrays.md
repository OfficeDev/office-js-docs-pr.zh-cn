---
ms.date: 05/11/2020
description: 从 Office Excel 外接程序中的自定义函数返回多个结果。
title: 从自定义函数返回多个结果
localization_priority: Normal
ms.openlocfilehash: 23ca1b038d73a93e6f96167cbdc23d79ccbfe622
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275971"
---
# <a name="return-multiple-results-from-your-custom-function"></a>从自定义函数返回多个结果

您可以从自定义函数返回多个结果，这些结果将返回到相邻的单元格。 此行为称为 "spilling"。 当您的自定义函数返回结果数组时，它被称为动态数组公式。 有关 Excel 中动态数组公式的详细信息，请参阅[动态数组和溢出的数组行为](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)。

下图显示了函数如何 `SORT` 扩散到相邻的单元格中。 您的自定义函数还可以返回如下所示的多个结果。

![将多个结果显示为多个单元格的 "排序" 函数的屏幕截图。](../images/dynamic-array-spill.png)

若要创建一个动态数组公式的自定义函数，它必须返回一个二维值数组。 如果结果溢出到已有值的相邻单元格，则公式将显示 `#SPILL!` 错误。

下面的示例演示如何返回泼溅的动态数组。

```javascript
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}
```

下面的示例演示如何返回一个靠右的动态数组。 

```javascript
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}
```

下面的示例演示如何返回一个动态数组，该数组同时扩散到右侧和右侧。

```javascript
/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}
```

## <a name="see-also"></a>另请参阅

- [动态数组和溢出的数组行为](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Excel 自定义函数的选项](custom-functions-parameter-options.md)