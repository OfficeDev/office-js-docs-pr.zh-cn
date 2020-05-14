---
ms.date: 12/18/2019
description: 从 Office Excel 外接程序中的自定义函数返回多个结果。
title: 从自定义函数返回多个结果
localization_priority: Normal
ms.openlocfilehash: 753755b481ab3db0de711c80ef082aedc82177ae
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217835"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="c0dd9-103">从自定义函数返回多个结果</span><span class="sxs-lookup"><span data-stu-id="c0dd9-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="c0dd9-104">您可以从自定义函数返回多个结果，这些结果将返回到相邻的单元格。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="c0dd9-105">此行为称为 "spilling"。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-105">This behavior is called spilling.</span></span> <span data-ttu-id="c0dd9-106">当您的自定义函数返回结果数组时，它被称为动态数组公式。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-106">When your custom function returns an array of results, it is known as a dynamic array formula.</span></span> <span data-ttu-id="c0dd9-107">有关 Excel 中动态数组公式的详细信息，请参阅[动态数组和溢出的数组行为](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="c0dd9-108">下图显示了函数如何 `SORT` 扩散到相邻的单元格中。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-108">The following image shows how the `SORT` function spills down into neighboring cells.</span></span> <span data-ttu-id="c0dd9-109">您的自定义函数还可以返回如下所示的多个结果。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-109">Your custom function can also return multiple results like this.</span></span>

![将多个结果显示为多个单元格的 "排序" 函数的屏幕截图。](../images/dynamic-array-spill.png)

<span data-ttu-id="c0dd9-111">若要创建一个动态数组公式的自定义函数，它必须返回一个二维值数组。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="c0dd9-112">如果结果溢出到已有值的相邻单元格，则公式将显示 `#SPILL!` 错误。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-112">If the results spill into neighboring cells that already have values, the formula will display a `#SPILL!` error.</span></span>

<span data-ttu-id="c0dd9-113">下面的示例演示如何返回泼溅的动态数组。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-113">The following example shows how to return a dynamic array that spills down.</span></span>

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

<span data-ttu-id="c0dd9-114">下面的示例演示如何返回一个靠右的动态数组。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-114">The following example shows how to return a dynamic array that spills right.</span></span> 

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

<span data-ttu-id="c0dd9-115">下面的示例演示如何返回一个动态数组，该数组同时扩散到右侧和右侧。</span><span class="sxs-lookup"><span data-stu-id="c0dd9-115">The following example shows how to return a dynamic array that spills both down and right.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="c0dd9-116">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c0dd9-116">See also</span></span>

- [<span data-ttu-id="c0dd9-117">动态数组和溢出的数组行为</span><span class="sxs-lookup"><span data-stu-id="c0dd9-117">Dynamic arrays and spilled array behavior</span></span>](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="c0dd9-118">Excel 自定义函数的选项</span><span class="sxs-lookup"><span data-stu-id="c0dd9-118">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)