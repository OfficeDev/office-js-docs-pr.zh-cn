---
ms.date: 01/14/2020
description: 了解如何实现易失性和脱机流式处理自定义函数。
title: 函数中的可变值
localization_priority: Normal
ms.openlocfilehash: 617599a2687696a96240c4f162f9b02788a215f4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717206"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="b1610-103">函数中的可变值</span><span class="sxs-lookup"><span data-stu-id="b1610-103">Volatile values in functions</span></span>

<span data-ttu-id="b1610-104">可变函数是值在每次计算单元格时更改的函数。</span><span class="sxs-lookup"><span data-stu-id="b1610-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="b1610-105">即使函数的所有参数都不变，该值也可以更改。</span><span class="sxs-lookup"><span data-stu-id="b1610-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="b1610-106">每当 Excel 重新计算时，这些函数即会重新计算。</span><span class="sxs-lookup"><span data-stu-id="b1610-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="b1610-107">例如，假设某个单元格调用函数 `NOW`。</span><span class="sxs-lookup"><span data-stu-id="b1610-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="b1610-108">每当调用 `NOW` 时，它将自动返回当前的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="b1610-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="b1610-109">Excel 包含多个内置可变函数，例如 `RAND` 和 `TODAY`。</span><span class="sxs-lookup"><span data-stu-id="b1610-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="b1610-110">有关 Excel 可变函数的完整列表，请参阅[可变函数和非可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)。</span><span class="sxs-lookup"><span data-stu-id="b1610-110">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="b1610-111">利用自定义函数，您可以创建自己的可变函数，这在处理日期、时间、随机编号和建模时可能很有用。</span><span class="sxs-lookup"><span data-stu-id="b1610-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="b1610-112">例如， [Monte Carlo 模拟](https://en.wikipedia.org/wiki/Monte_Carlo_method)要求生成随机输入以确定最佳解决方案。</span><span class="sxs-lookup"><span data-stu-id="b1610-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="b1610-113">如果选择自动生成 JSON 文件，则使用 JSDoc 注释标记`@volatile`声明一个可变函数。</span><span class="sxs-lookup"><span data-stu-id="b1610-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDoc comment tag `@volatile`.</span></span> <span data-ttu-id="b1610-114">有关自动生成的详细信息，请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="b1610-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="b1610-115">可变自定义函数的示例如下所示，模拟掷出六个侧骰子的情况。</span><span class="sxs-lookup"><span data-stu-id="b1610-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

![显示自定义函数的 gif，该函数返回随机值以模拟掷出的六边骰子](../images/six-sided-die.gif)

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided() {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a><span data-ttu-id="b1610-117">后续步骤</span><span class="sxs-lookup"><span data-stu-id="b1610-117">Next steps</span></span>
<span data-ttu-id="b1610-118">了解如何[在自定义函数中保存状态](custom-functions-save-state.md)。</span><span class="sxs-lookup"><span data-stu-id="b1610-118">Learn how to [save state in your custom functions](custom-functions-save-state.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b1610-119">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b1610-119">See also</span></span>

* [<span data-ttu-id="b1610-120">自定义函数参数选项</span><span class="sxs-lookup"><span data-stu-id="b1610-120">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="b1610-121">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="b1610-121">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="b1610-122">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="b1610-122">Create custom functions in Excel</span></span>](custom-functions-overview.md)
