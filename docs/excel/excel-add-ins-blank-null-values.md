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
# <a name="blank-and-null-values-in-excel-add-ins"></a><span data-ttu-id="b46f3-103">Excel 外接程序中的空值和 null 值</span><span class="sxs-lookup"><span data-stu-id="b46f3-103">Blank and null values in Excel add-ins</span></span>

<span data-ttu-id="b46f3-104">`null` 和空字符串在 Excel JavaScript API 中具有特殊含义。</span><span class="sxs-lookup"><span data-stu-id="b46f3-104">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="b46f3-105">它们用于表示空单元格、无格式或默认值。</span><span class="sxs-lookup"><span data-stu-id="b46f3-105">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="b46f3-106">本节详细介绍了在获取和设置属性时如何使用 `null` 和空字符串。</span><span class="sxs-lookup"><span data-stu-id="b46f3-106">This section details the use of `null` and empty string when getting and setting properties.</span></span>

## <a name="null-input-in-2-d-array"></a><span data-ttu-id="b46f3-107">二维数组中的 null 输入</span><span class="sxs-lookup"><span data-stu-id="b46f3-107">null input in 2-D Array</span></span>

<span data-ttu-id="b46f3-p102">在 Excel 中，一个区域由一个二维数组表示，其中第一个维度是行，第二个维度是列。 若要仅为某个区域内的特定单元格设置值、数字格式或公式，请指定二维数组中这些单元格的值、数字格式或公式，并为二维数组中的所有其他单元格指定 `null`。</span><span class="sxs-lookup"><span data-stu-id="b46f3-p102">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="b46f3-p103">例如，要更新一个区域内某一个单元格的数字格式，并保留该区域内所有其他单元格的现有数字格式，可指定要更新的单元格的新数字格式，并为所有其他单元格指定 `null`。 下面的代码段为该区域内的第四个单元格设置了一个新的数字格式，并保留该区域内前三个单元格的数字格式不变。</span><span class="sxs-lookup"><span data-stu-id="b46f3-p103">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a><span data-ttu-id="b46f3-112">属性的 null 输入</span><span class="sxs-lookup"><span data-stu-id="b46f3-112">null input for a property</span></span>

<span data-ttu-id="b46f3-p104">`null` 不是单个属性的有效输入。例如，下面的代码片段无效，因为区域的 `values` 属性不能设置为 `null`。</span><span class="sxs-lookup"><span data-stu-id="b46f3-p104">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null; // This is not a valid snippet. 
```

<span data-ttu-id="b46f3-115">同样，下面的代码片段也无效，因为 `null` 不是 `color` 属性的有效值。</span><span class="sxs-lookup"><span data-stu-id="b46f3-115">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a><span data-ttu-id="b46f3-116">响应中的 null 属性值</span><span class="sxs-lookup"><span data-stu-id="b46f3-116">null property values in the response</span></span>

<span data-ttu-id="b46f3-p105">如果指定区域内存在不同的值，诸如 `size` 和 `color` 等格式化属性将在响应中包含 `null` 值。 例如，如果你检索某个区域并加载其 `format.font.color` 属性：</span><span class="sxs-lookup"><span data-stu-id="b46f3-p105">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="b46f3-119">如果区域中的所有单元格都具有相同的字体颜色，则 `range.format.font.color` 会指定该颜色。</span><span class="sxs-lookup"><span data-stu-id="b46f3-119">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="b46f3-120">如果该区域内存在多种字体颜色，则 `range.format.font.color` 为 `null`。</span><span class="sxs-lookup"><span data-stu-id="b46f3-120">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

## <a name="blank-input-for-a-property"></a><span data-ttu-id="b46f3-121">属性的空白输入</span><span class="sxs-lookup"><span data-stu-id="b46f3-121">Blank input for a property</span></span>

<span data-ttu-id="b46f3-p106">如果为属性指定空白值（即两个引号之间没有空格 `''`），它会被解释为属性清除或重置指令。例如：</span><span class="sxs-lookup"><span data-stu-id="b46f3-p106">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="b46f3-124">如果为区域的 `values` 属性指定空白值，此区域的内容会被清除。</span><span class="sxs-lookup"><span data-stu-id="b46f3-124">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="b46f3-125">如果为 `numberFormat` 属性指定一个空值，则数字格式会重置为 `General`。</span><span class="sxs-lookup"><span data-stu-id="b46f3-125">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="b46f3-126">如果为 `formula` 属性和 `formulaLocale` 属性指定一个空值，则公式值将被清除。</span><span class="sxs-lookup"><span data-stu-id="b46f3-126">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

## <a name="blank-property-values-in-the-response"></a><span data-ttu-id="b46f3-127">响应中的空属性值</span><span class="sxs-lookup"><span data-stu-id="b46f3-127">Blank property values in the response</span></span>

<span data-ttu-id="b46f3-p107">对于读取操作，响应中的空属性值（即两个引号之间没有空格 `''`）指示该单元格不包含任何数据或值。 在下面第一个示例中，区域中的第一个和最后一个单元格不包含任何数据。 在第二个示例中，区域中的前两个单元格不包含公式。</span><span class="sxs-lookup"><span data-stu-id="b46f3-p107">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```
