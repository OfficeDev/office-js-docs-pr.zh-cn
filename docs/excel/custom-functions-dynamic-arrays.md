---
ms.date: 12/18/2019
description: 从 Office Excel 外接程序中的自定义函数返回多个结果。
title: 从自定义函数返回多个结果
localization_priority: Normal
ms.openlocfilehash: 687ffcd66cff16d92fec372a778fe94bad7b38d5
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2020
ms.locfileid: "40970370"
---
# <a name="return-multiple-results-from-your-custom-function"></a><span data-ttu-id="3e989-103">从自定义函数返回多个结果</span><span class="sxs-lookup"><span data-stu-id="3e989-103">Return multiple results from your custom function</span></span>

<span data-ttu-id="3e989-104">您可以从自定义函数返回多个结果，这些结果将返回到相邻的单元格。</span><span class="sxs-lookup"><span data-stu-id="3e989-104">You can return multiple results from your custom function which will be returned to neighboring cells.</span></span> <span data-ttu-id="3e989-105">此行为称为 "spilling"。</span><span class="sxs-lookup"><span data-stu-id="3e989-105">This behavior is called spilling.</span></span> <span data-ttu-id="3e989-106">当您的自定义函数返回结果数组时，它被称为动态数组公式。</span><span class="sxs-lookup"><span data-stu-id="3e989-106">When your custom function returns an array of results, it is known as a dynamic array formula.</span></span> <span data-ttu-id="3e989-107">有关 Excel 中动态数组公式的详细信息，请参阅[动态数组和溢出的数组行为](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)。</span><span class="sxs-lookup"><span data-stu-id="3e989-107">For more information on dynamic array formulas in Excel, see [Dynamic arrays and spilled array behavior](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span>

<span data-ttu-id="3e989-108">下图显示了**SORT**函数如何扩散到相邻的单元格中。</span><span class="sxs-lookup"><span data-stu-id="3e989-108">The following image shows how the **SORT** function spills down into neighboring cells.</span></span> <span data-ttu-id="3e989-109">您的自定义函数还可以返回如下所示的多个结果。</span><span class="sxs-lookup"><span data-stu-id="3e989-109">Your custom function can also return multiple results like this.</span></span>

![将多个结果显示为多个单元格的排序函数的屏幕截图。](../images/dynamic-array-spill.png)

<span data-ttu-id="3e989-111">若要创建一个动态数组公式的自定义函数，它必须返回一个二维值数组。</span><span class="sxs-lookup"><span data-stu-id="3e989-111">To create a custom function that is a dynamic array formula, it must return a two-dimensional array of values.</span></span> <span data-ttu-id="3e989-112">如果结果溢出到已有值的相邻单元格，则公式将显示 **#SPILL！**</span><span class="sxs-lookup"><span data-stu-id="3e989-112">If the results spill into neighboring cells that already have values, the formula will display a **#SPILL!**</span></span> <span data-ttu-id="3e989-113">误差.</span><span class="sxs-lookup"><span data-stu-id="3e989-113">error.</span></span> 

<span data-ttu-id="3e989-114">下面的示例演示如何返回泼溅的动态数组。</span><span class="sxs-lookup"><span data-stu-id="3e989-114">The following example shows how to return a dynamic array that spills down.</span></span>

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

<span data-ttu-id="3e989-115">下面的示例演示如何返回一个靠右的动态数组。</span><span class="sxs-lookup"><span data-stu-id="3e989-115">The following example shows how to return a dynamic array that spills right.</span></span> 

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

<span data-ttu-id="3e989-116">下面的示例演示如何返回一个动态数组，该数组同时扩散到右侧和右侧。</span><span class="sxs-lookup"><span data-stu-id="3e989-116">The following example shows how to return a dynamic array that spills both down and right.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="3e989-117">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3e989-117">See also</span></span>

- [<span data-ttu-id="3e989-118">动态数组和溢出的数组行为</span><span class="sxs-lookup"><span data-stu-id="3e989-118">Dynamic arrays and spilled array behavior</span></span>](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
- [<span data-ttu-id="3e989-119">Excel 自定义函数的选项</span><span class="sxs-lookup"><span data-stu-id="3e989-119">Options for Excel custom functions</span></span>](custom-functions-parameter-options.md)