---
ms.date: 01/14/2020
description: 了解如何实现可变和脱机流式处理自定义函数。
title: 函数中的可变值
localization_priority: Normal
ms.openlocfilehash: f441ef4fb7f90add5318546e3ccf4cc8bc60a8cf
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075885"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="e2fad-103">函数中的可变值</span><span class="sxs-lookup"><span data-stu-id="e2fad-103">Volatile values in functions</span></span>

<span data-ttu-id="e2fad-104">可变函数是每次计算单元格时值更改的函数。</span><span class="sxs-lookup"><span data-stu-id="e2fad-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="e2fad-105">即使函数的参数都未更改，值也可以更改。</span><span class="sxs-lookup"><span data-stu-id="e2fad-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="e2fad-106">每当 Excel 重新计算时，这些函数即会重新计算。</span><span class="sxs-lookup"><span data-stu-id="e2fad-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="e2fad-107">例如，假设某个单元格调用函数 `NOW`。</span><span class="sxs-lookup"><span data-stu-id="e2fad-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="e2fad-108">每当调用 `NOW` 时，它将自动返回当前的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="e2fad-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="e2fad-109">Excel 包含多个内置可变函数，例如 `RAND` 和 `TODAY`。</span><span class="sxs-lookup"><span data-stu-id="e2fad-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="e2fad-110">有关 Excel 可变函数的完整列表，请参阅[可变函数和非可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)。</span><span class="sxs-lookup"><span data-stu-id="e2fad-110">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="e2fad-111">自定义函数允许您创建自己的可变函数，在处理日期、时间、随机数字和建模时，这些函数可能很有用。</span><span class="sxs-lookup"><span data-stu-id="e2fad-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="e2fad-112">例如， [为确定最佳解决方案，将要求](https://en.wikipedia.org/wiki/Monte_Carlo_method) 生成随机输入。</span><span class="sxs-lookup"><span data-stu-id="e2fad-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="e2fad-113">如果选择自动生成 JSON 文件，请声明具有 JSDoc 注释标记的可变函数 `@volatile` 。</span><span class="sxs-lookup"><span data-stu-id="e2fad-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDoc comment tag `@volatile`.</span></span> <span data-ttu-id="e2fad-114">有关自动生成的信息，请参阅自动生成 [自定义函数的 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="e2fad-114">From more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="e2fad-115">以下是一个可变自定义函数的示例，该函数模拟滚动六面切纸。</span><span class="sxs-lookup"><span data-stu-id="e2fad-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

![显示返回随机值的自定义函数的 GIF，用于模拟滚动六面切纸。](../images/six-sided-die.gif)

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

## <a name="next-steps"></a><span data-ttu-id="e2fad-117">后续步骤</span><span class="sxs-lookup"><span data-stu-id="e2fad-117">Next steps</span></span>
* <span data-ttu-id="e2fad-118">了解自定义 [函数参数选项](custom-functions-parameter-options.md)。</span><span class="sxs-lookup"><span data-stu-id="e2fad-118">Learn about [custom functions parameter options](custom-functions-parameter-options.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e2fad-119">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e2fad-119">See also</span></span>

* [<span data-ttu-id="e2fad-120">手动为自定义函数创建 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="e2fad-120">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="e2fad-121">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="e2fad-121">Create custom functions in Excel</span></span>](custom-functions-overview.md)
