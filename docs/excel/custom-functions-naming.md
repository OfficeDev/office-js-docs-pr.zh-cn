---
ms.date: 07/10/2019
description: 了解 Excel 自定义函数名称的要求并避免出现常见命名缺陷。
title: Excel 中自定义函数的命名准则
localization_priority: Normal
ms.openlocfilehash: 79d0bfb069fe5abefeb6d0e88428d0728f3869e3
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771531"
---
# <a name="naming-guidelines"></a><span data-ttu-id="29a6a-103">命名准则</span><span class="sxs-lookup"><span data-stu-id="29a6a-103">Naming guidelines</span></span>

<span data-ttu-id="29a6a-104">自定义函数由 JSON 元数据文件中的**id**和**name**属性标识。</span><span class="sxs-lookup"><span data-stu-id="29a6a-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span>

- <span data-ttu-id="29a6a-105">函数`id`用于唯一标识 JavaScript 代码中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="29a6a-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span>
- <span data-ttu-id="29a6a-106">函数`name`用作在 Excel 中向用户显示的显示名称。</span><span class="sxs-lookup"><span data-stu-id="29a6a-106">The function `name` is used as the display name that appears to a user in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="29a6a-107">函数`name`可以与函数`id`不同, 例如出于本地化目的。</span><span class="sxs-lookup"><span data-stu-id="29a6a-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="29a6a-108">通常情况下, 如果没有`name`明显的原因, 函数应`id`保持与的相同。</span><span class="sxs-lookup"><span data-stu-id="29a6a-108">In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="29a6a-109">函数的`name`并`id`共享一些常见要求:</span><span class="sxs-lookup"><span data-stu-id="29a6a-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="29a6a-110">函数`id`可能只使用字符 A 到 Z、从零到九、下划线和句点。</span><span class="sxs-lookup"><span data-stu-id="29a6a-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="29a6a-111">函数`name`可能使用任何 Unicode 字母字符、下划线和句点。</span><span class="sxs-lookup"><span data-stu-id="29a6a-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="29a6a-112">这两`name`个`id`函数都必须以字母开头, 并且最小限制为三个字符。</span><span class="sxs-lookup"><span data-stu-id="29a6a-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="29a6a-113">Excel 使用大写字母作为内置函数名称 (例如`SUM`)。</span><span class="sxs-lookup"><span data-stu-id="29a6a-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="29a6a-114">因此, 请考虑将大写字母用作自定义函数`name`和`id`最佳实践。</span><span class="sxs-lookup"><span data-stu-id="29a6a-114">Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="29a6a-115">函数的`name`名称不应与以下相同:</span><span class="sxs-lookup"><span data-stu-id="29a6a-115">A function's `name` shouldn't be named the same as:</span></span>

- <span data-ttu-id="29a6a-116">A1 到 XFD1048576 之间的任何单元格, 或从 R1C1 到 R1048576C16384 之间的任何单元格。</span><span class="sxs-lookup"><span data-stu-id="29a6a-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="29a6a-117">任何 Excel 4.0 宏函数 (例如`RUN`, `ECHO`)。</span><span class="sxs-lookup"><span data-stu-id="29a6a-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="29a6a-118">有关这些函数的完整列表, 请参阅[此 Excel 宏函数参考文档](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)。</span><span class="sxs-lookup"><span data-stu-id="29a6a-118">For a full list of these functions, see [this Excel Macro Functions Reference document](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="29a6a-119">命名冲突</span><span class="sxs-lookup"><span data-stu-id="29a6a-119">Naming conflicts</span></span>

<span data-ttu-id="29a6a-120">如果您的`name`函数与已存在的外`name`接程序中的函数相同, 则 **#REF!**</span><span class="sxs-lookup"><span data-stu-id="29a6a-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="29a6a-121">错误将出现在工作簿中。</span><span class="sxs-lookup"><span data-stu-id="29a6a-121">error will appear in your workbook.</span></span>

<span data-ttu-id="29a6a-122">若要修复命名冲突, 请更改`name`外接程序中的, 然后再次尝试该函数。</span><span class="sxs-lookup"><span data-stu-id="29a6a-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="29a6a-123">此外, 还可以使用冲突的名称卸载加载项。</span><span class="sxs-lookup"><span data-stu-id="29a6a-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="29a6a-124">或者, 如果要在不同的环境中测试外接程序, 请尝试使用不同的命名空间来区分您的函数`NAMESPACE_NAMEOFFUNCTION`(如)。</span><span class="sxs-lookup"><span data-stu-id="29a6a-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="29a6a-125">最佳做法</span><span class="sxs-lookup"><span data-stu-id="29a6a-125">Best practices</span></span>

- <span data-ttu-id="29a6a-126">请考虑向函数中添加多个参数, 而不是使用相同或相似的名称创建多个函数。</span><span class="sxs-lookup"><span data-stu-id="29a6a-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="29a6a-127">函数名称应指示函数的操作, 例如 ( `=GETZIPCODE`而不是) `ZIPCODE`。</span><span class="sxs-lookup"><span data-stu-id="29a6a-127">Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.</span></span>
- <span data-ttu-id="29a6a-128">避免函数名称中不明确的缩写。</span><span class="sxs-lookup"><span data-stu-id="29a6a-128">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="29a6a-129">清晰度比简洁性更重要。</span><span class="sxs-lookup"><span data-stu-id="29a6a-129">Clarity is more important than brevity.</span></span> <span data-ttu-id="29a6a-130">选择一个名称 ( `=INCREASETIME`而不`=INC`是)。</span><span class="sxs-lookup"><span data-stu-id="29a6a-130">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="29a6a-131">对执行类似操作的函数始终使用相同的动作。</span><span class="sxs-lookup"><span data-stu-id="29a6a-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="29a6a-132">`=DELETEZIPCODE`例如, 使用`=DELETEADDRESS`和, 而不是`=DELETEZIPCODE`和`=REMOVEADDRESS`。</span><span class="sxs-lookup"><span data-stu-id="29a6a-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>
- <span data-ttu-id="29a6a-133">在命名流式处理函数时, 请考虑在函数的说明中添加对该效果的注释或`STREAM`添加到函数名称的末尾。</span><span class="sxs-lookup"><span data-stu-id="29a6a-133">When naming a streaming function, consider adding a note to that effect in the description of the function or adding `STREAM` to the end of the function's name.</span></span>

## <a name="localizing-function-names"></a><span data-ttu-id="29a6a-134">对函数名称进行本地化</span><span class="sxs-lookup"><span data-stu-id="29a6a-134">Localizing function names</span></span>

<span data-ttu-id="29a6a-135">您可以使用单独的 JSON 文件本地化不同语言的函数名称, 并在外接程序清单文件中重写值。</span><span class="sxs-lookup"><span data-stu-id="29a6a-135">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="29a6a-136">作为一种最佳做法, 应避免在另`id`一`name`种语言中为函数提供内置 Excel 函数, 因为这可能会与本地化函数发生冲突。</span><span class="sxs-lookup"><span data-stu-id="29a6a-136">As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="29a6a-137">有关本地化的完整信息, 请参阅[本地化自定义函数](custom-functions-localize.md)</span><span class="sxs-lookup"><span data-stu-id="29a6a-137">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="29a6a-138">后续步骤</span><span class="sxs-lookup"><span data-stu-id="29a6a-138">Next steps</span></span>
<span data-ttu-id="29a6a-139">了解[错误处理最佳实践](custom-functions-errors.md)。</span><span class="sxs-lookup"><span data-stu-id="29a6a-139">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="29a6a-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="29a6a-140">See also</span></span>

* [<span data-ttu-id="29a6a-141">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="29a6a-141">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="29a6a-142">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="29a6a-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="29a6a-143">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="29a6a-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
