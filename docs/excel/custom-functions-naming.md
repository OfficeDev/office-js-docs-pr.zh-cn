---
ms.date: 05/17/2020
description: 了解 Excel 自定义函数名称的要求并避免常见命名缺陷。
title: Excel 中自定义函数的命名准则
localization_priority: Normal
ms.openlocfilehash: ac0d824f49d359e574a0dc5caae8ef2f903dd4a1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609287"
---
# <a name="naming-guidelines"></a><span data-ttu-id="f461a-103">命名准则</span><span class="sxs-lookup"><span data-stu-id="f461a-103">Naming guidelines</span></span>

<span data-ttu-id="f461a-104">`id` `name` 在 JSON 元数据文件中，自定义函数由和属性标识。</span><span class="sxs-lookup"><span data-stu-id="f461a-104">A custom function is identified by an `id` and `name` property in the JSON metadata file.</span></span>

- <span data-ttu-id="f461a-105">函数 `id` 用于唯一标识 JavaScript 代码中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="f461a-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span>
- <span data-ttu-id="f461a-106">函数 `name` 用作在 Excel 中向用户显示的显示名称。</span><span class="sxs-lookup"><span data-stu-id="f461a-106">The function `name` is used as the display name that appears to a user in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f461a-107">函数 `name` 可以与函数不同，例如 `id` 出于本地化目的。</span><span class="sxs-lookup"><span data-stu-id="f461a-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="f461a-108">通常情况下， `name` `id` 如果没有理由让函数与相同，则函数应保持不变。</span><span class="sxs-lookup"><span data-stu-id="f461a-108">In general, a function's `name` should stay the same as the `id` if there is no reason for them to differ.</span></span>

<span data-ttu-id="f461a-109">函数的 `name` 并 `id` 共享一些常见要求：</span><span class="sxs-lookup"><span data-stu-id="f461a-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="f461a-110">函数 `id` 可能只使用字符 A 到 Z、从零到九、下划线和句点。</span><span class="sxs-lookup"><span data-stu-id="f461a-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="f461a-111">函数 `name` 可能使用任何 Unicode 字母字符、下划线和句点。</span><span class="sxs-lookup"><span data-stu-id="f461a-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="f461a-112">这两个函数都 `name` `id` 必须以字母开头，并且最小限制为三个字符。</span><span class="sxs-lookup"><span data-stu-id="f461a-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="f461a-113">Excel 使用大写字母作为内置函数名称（例如 `SUM` ）。</span><span class="sxs-lookup"><span data-stu-id="f461a-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="f461a-114">将大写字母用作自定义函数 `name` 和 `id` 最佳实践。</span><span class="sxs-lookup"><span data-stu-id="f461a-114">Use uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="f461a-115">函数 `name` 不应如下所示：</span><span class="sxs-lookup"><span data-stu-id="f461a-115">A function's `name` shouldn't be the same as:</span></span>

- <span data-ttu-id="f461a-116">A1 到 XFD1048576 之间的任何单元格，或从 R1C1 到 R1048576C16384 之间的任何单元格。</span><span class="sxs-lookup"><span data-stu-id="f461a-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="f461a-117">任何 Excel 4.0 宏函数（例如 `RUN` ， `ECHO` ）。</span><span class="sxs-lookup"><span data-stu-id="f461a-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="f461a-118">有关这些函数的完整列表，请参阅[此 Excel 宏函数参考文档](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)。</span><span class="sxs-lookup"><span data-stu-id="f461a-118">For a full list of these functions, see [this Excel Macro Functions Reference document](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="f461a-119">命名冲突</span><span class="sxs-lookup"><span data-stu-id="f461a-119">Naming conflicts</span></span>

<span data-ttu-id="f461a-120">如果您的函数与 `name` 已存在的外接程序中的函数相同 `name` ，则 **#REF！**</span><span class="sxs-lookup"><span data-stu-id="f461a-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="f461a-121">错误将出现在工作簿中。</span><span class="sxs-lookup"><span data-stu-id="f461a-121">error will appear in your workbook.</span></span>

<span data-ttu-id="f461a-122">若要修复命名冲突，请更改 `name` 外接程序中的，然后再次尝试该函数。</span><span class="sxs-lookup"><span data-stu-id="f461a-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="f461a-123">此外，还可以使用冲突的名称卸载加载项。</span><span class="sxs-lookup"><span data-stu-id="f461a-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="f461a-124">或者，如果要在不同的环境中测试外接程序，请尝试使用不同的命名空间来区分您的函数（如 `NAMESPACE_NAMEOFFUNCTION` ）。</span><span class="sxs-lookup"><span data-stu-id="f461a-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="f461a-125">最佳做法</span><span class="sxs-lookup"><span data-stu-id="f461a-125">Best practices</span></span>

- <span data-ttu-id="f461a-126">请考虑向函数中添加多个参数，而不是使用相同或相似的名称创建多个函数。</span><span class="sxs-lookup"><span data-stu-id="f461a-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="f461a-127">避免函数名称中不明确的缩写。</span><span class="sxs-lookup"><span data-stu-id="f461a-127">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="f461a-128">清晰度比简洁性更重要。</span><span class="sxs-lookup"><span data-stu-id="f461a-128">Clarity is more important than brevity.</span></span> <span data-ttu-id="f461a-129">选择一个名称（ `=INCREASETIME` 而不是） `=INC` 。</span><span class="sxs-lookup"><span data-stu-id="f461a-129">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="f461a-130">函数名称应指示函数的操作，如 = GETZIPCODE 而不是邮政编码。</span><span class="sxs-lookup"><span data-stu-id="f461a-130">Function names should indicate the action of the function, such as =GETZIPCODE instead of ZIPCODE.</span></span>
- <span data-ttu-id="f461a-131">对执行类似操作的函数始终使用相同的动作。</span><span class="sxs-lookup"><span data-stu-id="f461a-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="f461a-132">例如，使用 `=DELETEZIPCODE` 和 `=DELETEADDRESS` ，而不是 `=DELETEZIPCODE` 和 `=REMOVEADDRESS` 。</span><span class="sxs-lookup"><span data-stu-id="f461a-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>
- <span data-ttu-id="f461a-133">在命名流式处理函数时，请考虑在函数的说明中添加对该效果的注释或添加 `STREAM` 到函数名称的末尾。</span><span class="sxs-lookup"><span data-stu-id="f461a-133">When naming a streaming function, consider adding a note to that effect in the description of the function or adding `STREAM` to the end of the function's name.</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a><span data-ttu-id="f461a-134">对函数名称进行本地化</span><span class="sxs-lookup"><span data-stu-id="f461a-134">Localizing function names</span></span>

<span data-ttu-id="f461a-135">您可以使用单独的 JSON 文件本地化不同语言的函数名称，并在外接程序清单文件中重写值。</span><span class="sxs-lookup"><span data-stu-id="f461a-135">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="f461a-136">避免为您的函数 `id` 提供 `name` 另一种语言的内置 Excel 函数，因为这可能会与本地化函数发生冲突。</span><span class="sxs-lookup"><span data-stu-id="f461a-136">Avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="f461a-137">有关本地化的完整信息，请参阅[本地化自定义函数](custom-functions-localize.md)</span><span class="sxs-lookup"><span data-stu-id="f461a-137">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="f461a-138">后续步骤</span><span class="sxs-lookup"><span data-stu-id="f461a-138">Next steps</span></span>
<span data-ttu-id="f461a-139">了解[错误处理最佳实践](custom-functions-errors.md)。</span><span class="sxs-lookup"><span data-stu-id="f461a-139">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f461a-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f461a-140">See also</span></span>

* [<span data-ttu-id="f461a-141">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="f461a-141">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f461a-142">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="f461a-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
