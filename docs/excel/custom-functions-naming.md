---
ms.date: 05/03/2019
description: 了解 Excel 自定义函数名称的要求并避免出现常见命名缺陷。
title: Excel 中自定义函数的命名准则
localization_priority: Normal
ms.openlocfilehash: 3abe04eebfa703666b70ecbde1c68ab0c942003c
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628044"
---
# <a name="naming-guidelines"></a><span data-ttu-id="5c27e-103">命名准则</span><span class="sxs-lookup"><span data-stu-id="5c27e-103">Naming guidelines</span></span>

<span data-ttu-id="5c27e-104">自定义函数由 JSON 元数据文件中的**id**和**name**属性标识。</span><span class="sxs-lookup"><span data-stu-id="5c27e-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

- <span data-ttu-id="5c27e-105">函数`id`用于唯一标识 JavaScript 代码中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="5c27e-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span> 
- <span data-ttu-id="5c27e-106">函数`name`用作在 Excel 中向用户显示的显示名称。</span><span class="sxs-lookup"><span data-stu-id="5c27e-106">The function `name` is used as the display name that appears to a user in Excel.</span></span> 

<span data-ttu-id="5c27e-107">函数`name`可以与函数`id`不同, 例如出于本地化目的。</span><span class="sxs-lookup"><span data-stu-id="5c27e-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="5c27e-108">通常情况下, 如果没有`name`明显的原因, 函数应`id`保持与的相同。</span><span class="sxs-lookup"><span data-stu-id="5c27e-108">In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="5c27e-109">函数的`name`并`id`共享一些常见要求:</span><span class="sxs-lookup"><span data-stu-id="5c27e-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="5c27e-110">函数`id`可能只使用字符 A 到 Z、从零到九、下划线和句点。</span><span class="sxs-lookup"><span data-stu-id="5c27e-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="5c27e-111">函数`name`可能使用任何 Unicode 字母字符、下划线和句点。</span><span class="sxs-lookup"><span data-stu-id="5c27e-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="5c27e-112">这两`name`个`id`函数都必须以字母开头, 并且最小限制为三个字符。</span><span class="sxs-lookup"><span data-stu-id="5c27e-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="5c27e-113">Excel 使用大写字母作为内置函数名称 (例如`SUM`)。</span><span class="sxs-lookup"><span data-stu-id="5c27e-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="5c27e-114">因此, 请考虑将大写字母用作自定义函数`name`和`id`最佳实践。</span><span class="sxs-lookup"><span data-stu-id="5c27e-114">Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="5c27e-115">函数的`name`名称不应与以下相同:</span><span class="sxs-lookup"><span data-stu-id="5c27e-115">A function's `name` shouldn't be named the same as:</span></span>

- <span data-ttu-id="5c27e-116">A1 到 XFD1048576 之间的任何单元格, 或从 R1C1 到 R1048576C16384 之间的任何单元格。</span><span class="sxs-lookup"><span data-stu-id="5c27e-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="5c27e-117">任何 Excel 4.0 宏函数 (例如`RUN`, `ECHO`)。</span><span class="sxs-lookup"><span data-stu-id="5c27e-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="5c27e-118">有关这些函数的完整列表, 请参阅[本文](https://www.microsoft.com/en-us/download/details.aspx?id=1465)。</span><span class="sxs-lookup"><span data-stu-id="5c27e-118">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="5c27e-119">命名冲突</span><span class="sxs-lookup"><span data-stu-id="5c27e-119">Naming conflicts</span></span>

<span data-ttu-id="5c27e-120">如果您的`name`函数与已存在的外`name`接程序中的函数相同, 则 **#REF!**</span><span class="sxs-lookup"><span data-stu-id="5c27e-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="5c27e-121">错误将出现在工作簿中。</span><span class="sxs-lookup"><span data-stu-id="5c27e-121">error will appear in your workbook.</span></span>

<span data-ttu-id="5c27e-122">若要修复命名冲突, 请更改`name`外接程序中的, 然后再次尝试该函数。</span><span class="sxs-lookup"><span data-stu-id="5c27e-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="5c27e-123">此外, 还可以使用冲突的名称卸载加载项。</span><span class="sxs-lookup"><span data-stu-id="5c27e-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="5c27e-124">或者, 如果要在不同的环境中测试外接程序, 请尝试使用不同的命名空间来区分您的函数`NAMESPACE_NAMEOFFUNCTION`(如)。</span><span class="sxs-lookup"><span data-stu-id="5c27e-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="5c27e-125">最佳做法</span><span class="sxs-lookup"><span data-stu-id="5c27e-125">Best practices</span></span>

- <span data-ttu-id="5c27e-126">请考虑向函数中添加多个参数, 而不是使用相同或相似的名称创建多个函数。</span><span class="sxs-lookup"><span data-stu-id="5c27e-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="5c27e-127">函数名称应指示函数的操作, 例如 ( `=GETZIPCODE`而不是) `ZIPCODE`。</span><span class="sxs-lookup"><span data-stu-id="5c27e-127">Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.</span></span>
- <span data-ttu-id="5c27e-128">避免函数名称中不明确的缩写。</span><span class="sxs-lookup"><span data-stu-id="5c27e-128">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="5c27e-129">清晰度比简洁性更重要。</span><span class="sxs-lookup"><span data-stu-id="5c27e-129">Clarity is more important than brevity.</span></span> <span data-ttu-id="5c27e-130">选择一个名称 ( `=INCREASETIME`而不`=INC`是)。</span><span class="sxs-lookup"><span data-stu-id="5c27e-130">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="5c27e-131">对执行类似操作的函数始终使用相同的动作。</span><span class="sxs-lookup"><span data-stu-id="5c27e-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="5c27e-132">`=DELETEZIPCODE`例如, 使用`=DELETEADDRESS`和, 而不是`=DELETEZIPCODE`和`=REMOVEADDRESS`。</span><span class="sxs-lookup"><span data-stu-id="5c27e-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>

## <a name="localizing-function-names"></a><span data-ttu-id="5c27e-133">对函数名称进行本地化</span><span class="sxs-lookup"><span data-stu-id="5c27e-133">Localizing function names</span></span>

<span data-ttu-id="5c27e-134">您可以使用单独的 JSON 文件本地化不同语言的函数名称, 并在外接程序清单文件中重写值。</span><span class="sxs-lookup"><span data-stu-id="5c27e-134">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="5c27e-135">作为一种最佳做法, 应避免在另`id`一`name`种语言中为函数提供内置 Excel 函数, 因为这可能会与本地化函数发生冲突。</span><span class="sxs-lookup"><span data-stu-id="5c27e-135">As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="5c27e-136">有关本地化的完整信息, 请参阅[本地化自定义函数](custom-functions-localize.md)</span><span class="sxs-lookup"><span data-stu-id="5c27e-136">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="5c27e-137">后续步骤</span><span class="sxs-lookup"><span data-stu-id="5c27e-137">Next steps</span></span>
<span data-ttu-id="5c27e-138">了解[错误处理最佳实践](custom-functions-errors.md)。</span><span class="sxs-lookup"><span data-stu-id="5c27e-138">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="5c27e-139">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5c27e-139">See also</span></span>

* [<span data-ttu-id="5c27e-140">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="5c27e-140">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="5c27e-141">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="5c27e-141">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="5c27e-142">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="5c27e-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="5c27e-143">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="5c27e-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
