---
title: Excel 外接程序中的空值和 null 值
description: 了解如何在 Excel 对象模型方法和属性中使用空值。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 3f38569f7342bb88c52ce424db426bfa7939be5e
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409384"
---
# <a name="blank-and-null-values-in-excel-add-ins"></a>Excel 外接程序中的空值和 null 值

`null` 和空字符串在 Excel JavaScript API 中具有特殊含义。 它们用于表示空单元格、无格式或默认值。 本节详细介绍了在获取和设置属性时如何使用 `null` 和空字符串。

## <a name="null-input-in-2-d-array"></a>二维数组中的 null 输入

在 Excel 中，一个区域由一个二维数组表示，其中第一个维度是行，第二个维度是列。 若要仅为某个区域内的特定单元格设置值、数字格式或公式，请指定二维数组中这些单元格的值、数字格式或公式，并为二维数组中的所有其他单元格指定 `null`。

例如，要更新一个区域内某一个单元格的数字格式，并保留该区域内所有其他单元格的现有数字格式，可指定要更新的单元格的新数字格式，并为所有其他单元格指定 `null`。 下面的代码段为该区域内的第四个单元格设置了一个新的数字格式，并保留该区域内前三个单元格的数字格式不变。

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a>属性的 null 输入

`null` 不是单个属性的有效输入。例如，下面的代码片段无效，因为区域的 `values` 属性不能设置为 `null`。

```js
range.values = null; // This is not a valid snippet. 
```

同样，下面的代码片段也无效，因为 `null` 不是 `color` 属性的有效值。

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a>响应中的 null 属性值

如果指定区域内存在不同的值，诸如 `size` 和 `color` 等格式化属性将在响应中包含 `null` 值。 例如，如果你检索某个区域并加载其 `format.font.color` 属性：

* 如果区域中的所有单元格都具有相同的字体颜色，则 `range.format.font.color` 会指定该颜色。
* 如果该区域内存在多种字体颜色，则 `range.format.font.color` 为 `null`。

## <a name="blank-input-for-a-property"></a>属性的空白输入

如果为属性指定空白值（即两个引号之间没有空格 `''`），它会被解释为属性清除或重置指令。例如：

* 如果为区域的 `values` 属性指定空白值，此区域的内容会被清除。
* 如果为 `numberFormat` 属性指定一个空值，则数字格式会重置为 `General`。
* 如果为 `formula` 属性和 `formulaLocale` 属性指定一个空值，则公式值将被清除。

## <a name="blank-property-values-in-the-response"></a>响应中的空属性值

对于读取操作，响应中的空属性值（即两个引号之间没有空格 `''`）指示该单元格不包含任何数据或值。 在下面第一个示例中，区域中的第一个和最后一个单元格不包含任何数据。 在第二个示例中，区域中的前两个单元格不包含公式。

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```
