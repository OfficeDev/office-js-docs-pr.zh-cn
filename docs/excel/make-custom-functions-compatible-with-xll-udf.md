---
title: 使用 XLL 用户定义的函数扩展自定义函数
description: 启用与自定义函数具有等效功能的 Excel XLL 用户定义函数的兼容性 (预览)
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 93e1b52606fca7ea6fbbb9ae3545e4edd7f78742
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628102"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions-preview"></a><span data-ttu-id="2a03a-103">使用 XLL 用户定义的函数扩展自定义函数 (预览)</span><span class="sxs-lookup"><span data-stu-id="2a03a-103">Extend custom functions with XLL user-defined functions (preview)</span></span>

<span data-ttu-id="2a03a-104">如果您有现有的 Excel Xll, 则可以在 Excel 外接程序中构建等效的自定义函数, 以将解决方案功能扩展到其他平台 (如 online 或 macOS)。</span><span class="sxs-lookup"><span data-stu-id="2a03a-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="2a03a-105">但是, Excel 外接程序没有在 Xll 中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="2a03a-105">However, Excel add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="2a03a-106">根据您的解决方案使用的功能, XLL 可以提供比 Excel for Windows 上的 Excel 加载项自定义函数更好的体验。</span><span class="sxs-lookup"><span data-stu-id="2a03a-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Excel add-in custom functions on Excel for Windows.</span></span>

[!include[COM add-in and XLL UDF compatibility note](../includes/xll-compatibility-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="2a03a-107">在清单中指定等效 XLL</span><span class="sxs-lookup"><span data-stu-id="2a03a-107">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="2a03a-108">若要启用与现有 XLL 的兼容性, 请在您的 Excel 外接程序清单中标识等效 XLL。</span><span class="sxs-lookup"><span data-stu-id="2a03a-108">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Excel add-in.</span></span> <span data-ttu-id="2a03a-109">在 Windows 上运行时, Excel 将使用 XLL 的函数而不是 Excel 加载项自定义函数。</span><span class="sxs-lookup"><span data-stu-id="2a03a-109">Then Excel will use the XLL's functions instead of your Excel add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="2a03a-110">若要设置自定义函数的等效 XLL, 请指定`FileName` XLL 的。</span><span class="sxs-lookup"><span data-stu-id="2a03a-110">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="2a03a-111">当用户使用 XLL 中的函数打开工作簿时, Excel 会将函数转换为兼容函数。</span><span class="sxs-lookup"><span data-stu-id="2a03a-111">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="2a03a-112">在 Windows 上的 Excel 中打开时, 工作簿将使用 XLL, 并且在联机或在 macOS 中打开时, 它将使用 Excel 外接程序中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="2a03a-112">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Excel add-in when opened online or on macOS.</span></span>

<span data-ttu-id="2a03a-113">下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。</span><span class="sxs-lookup"><span data-stu-id="2a03a-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="2a03a-114">通常, 出于完整性的考虑, 这两个示例都会在上下文中显示这两个示例。</span><span class="sxs-lookup"><span data-stu-id="2a03a-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="2a03a-115">它们`ProgID` `FileName`分别由各自标识。</span><span class="sxs-lookup"><span data-stu-id="2a03a-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="2a03a-116">有关 COM 加载项兼容性的详细信息, 请参阅[使您的 Excel 外接程序与现有的 com 外](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)接程序兼容。</span><span class="sxs-lookup"><span data-stu-id="2a03a-116">For more information on COM add-in compatibility, see [Make your Excel add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="2a03a-117">如果外接程序声明其自定义函数是 XLL 兼容的, 则稍后更改清单可能会破坏用户的工作簿, 因为它会更改文件格式。</span><span class="sxs-lookup"><span data-stu-id="2a03a-117">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.</span></span>

## <a name="excel-add-in-updates"></a><span data-ttu-id="2a03a-118">Excel 加载项更新</span><span class="sxs-lookup"><span data-stu-id="2a03a-118">Excel add-in updates</span></span>

<span data-ttu-id="2a03a-119">为 Excel 加载项指定等效 XLL 后, Excel 将停止处理 Excel 加载项的更新。</span><span class="sxs-lookup"><span data-stu-id="2a03a-119">Once you specify an equivalent XLL for your Excel add-in, Excel stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="2a03a-120">用户必须卸载 XLL 才能获取 Excel 外接程序的最新更新。</span><span class="sxs-lookup"><span data-stu-id="2a03a-120">The user must uninstall the XLL in order to get the latest updates for the Excel add-in.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="2a03a-121">XLL 兼容函数的自定义函数行为</span><span class="sxs-lookup"><span data-stu-id="2a03a-121">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="2a03a-122">如果打开的电子表格中包含的 XLL 函数也有等效的加载项, 则 XLL 的函数将转换为 XLL 兼容的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="2a03a-122">When a spreadsheet is opened that contains XLL functions for which there is also an equivalent add-in, the XLL's functions are converted to XLL compatible custom functions.</span></span> <span data-ttu-id="2a03a-123">在下一次保存时, 它们将在兼容模式下写入文件, 以便它们使用 XLL 和 Excel 外接程序自定义函数 (当在其他平台上)。</span><span class="sxs-lookup"><span data-stu-id="2a03a-123">On the next save, they are written to the file in a compatible mode so that they work with both the XLL and Excel add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="2a03a-124">下表比较了 XLL 用户定义函数、XLL 兼容的自定义函数和 Excel 加载项自定义函数之间的功能。</span><span class="sxs-lookup"><span data-stu-id="2a03a-124">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel add-in custom functions.</span></span>

|         |<span data-ttu-id="2a03a-125">XLL 用户定义的函数</span><span class="sxs-lookup"><span data-stu-id="2a03a-125">XLL user-defined function</span></span> |<span data-ttu-id="2a03a-126">XLL 兼容的自定义函数</span><span class="sxs-lookup"><span data-stu-id="2a03a-126">XLL compatible custom functions</span></span> |<span data-ttu-id="2a03a-127">Excel 加载项自定义函数</span><span class="sxs-lookup"><span data-stu-id="2a03a-127">Excel add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="2a03a-128">支持的平台</span><span class="sxs-lookup"><span data-stu-id="2a03a-128">Supported platforms</span></span> | <span data-ttu-id="2a03a-129">Windows</span><span class="sxs-lookup"><span data-stu-id="2a03a-129">Windows</span></span> | <span data-ttu-id="2a03a-130">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="2a03a-130">Windows, macOS, Excel online</span></span> | <span data-ttu-id="2a03a-131">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="2a03a-131">Windows, macOS, Excel online</span></span> |
| <span data-ttu-id="2a03a-132">支持的文件格式</span><span class="sxs-lookup"><span data-stu-id="2a03a-132">Supported file formats</span></span> | <span data-ttu-id="2a03a-133">.XLSX、XLSB、XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="2a03a-133">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="2a03a-134">.XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="2a03a-134">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="2a03a-135">.XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="2a03a-135">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="2a03a-136">公式自动完成</span><span class="sxs-lookup"><span data-stu-id="2a03a-136">Formula autocomplete</span></span> | <span data-ttu-id="2a03a-137">否</span><span class="sxs-lookup"><span data-stu-id="2a03a-137">No</span></span> | <span data-ttu-id="2a03a-138">可访问</span><span class="sxs-lookup"><span data-stu-id="2a03a-138">Yes</span></span> | <span data-ttu-id="2a03a-139">是</span><span class="sxs-lookup"><span data-stu-id="2a03a-139">Yes</span></span> |
| <span data-ttu-id="2a03a-140">媒体</span><span class="sxs-lookup"><span data-stu-id="2a03a-140">Streaming</span></span> | <span data-ttu-id="2a03a-141">可通过 xlfRTD 和 XLL 回调实现。</span><span class="sxs-lookup"><span data-stu-id="2a03a-141">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="2a03a-142">否</span><span class="sxs-lookup"><span data-stu-id="2a03a-142">No</span></span> | <span data-ttu-id="2a03a-143">可访问</span><span class="sxs-lookup"><span data-stu-id="2a03a-143">Yes</span></span> |
| <span data-ttu-id="2a03a-144">函数的本地化</span><span class="sxs-lookup"><span data-stu-id="2a03a-144">Localization of functions</span></span> | <span data-ttu-id="2a03a-145">否</span><span class="sxs-lookup"><span data-stu-id="2a03a-145">No</span></span> | <span data-ttu-id="2a03a-146">否。</span><span class="sxs-lookup"><span data-stu-id="2a03a-146">No.</span></span> <span data-ttu-id="2a03a-147">名称和 ID 必须与现有 XLL 的函数相匹配。</span><span class="sxs-lookup"><span data-stu-id="2a03a-147">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="2a03a-148">是</span><span class="sxs-lookup"><span data-stu-id="2a03a-148">Yes</span></span> |
| <span data-ttu-id="2a03a-149">可变函数</span><span class="sxs-lookup"><span data-stu-id="2a03a-149">Volatile functions</span></span> | <span data-ttu-id="2a03a-150">是</span><span class="sxs-lookup"><span data-stu-id="2a03a-150">Yes</span></span> | <span data-ttu-id="2a03a-151">是</span><span class="sxs-lookup"><span data-stu-id="2a03a-151">Yes</span></span> | <span data-ttu-id="2a03a-152">是</span><span class="sxs-lookup"><span data-stu-id="2a03a-152">Yes</span></span> |
| <span data-ttu-id="2a03a-153">多线程重新计算支持</span><span class="sxs-lookup"><span data-stu-id="2a03a-153">Multi-threaded recalculation support</span></span> | <span data-ttu-id="2a03a-154">是</span><span class="sxs-lookup"><span data-stu-id="2a03a-154">Yes</span></span> | <span data-ttu-id="2a03a-155">是</span><span class="sxs-lookup"><span data-stu-id="2a03a-155">Yes</span></span> | <span data-ttu-id="2a03a-156">是</span><span class="sxs-lookup"><span data-stu-id="2a03a-156">Yes</span></span> |
| <span data-ttu-id="2a03a-157">计算行为</span><span class="sxs-lookup"><span data-stu-id="2a03a-157">Calculation behavior</span></span> | <span data-ttu-id="2a03a-158">无 UI。</span><span class="sxs-lookup"><span data-stu-id="2a03a-158">No UI.</span></span> <span data-ttu-id="2a03a-159">在计算过程中, Excel 可能会无响应。</span><span class="sxs-lookup"><span data-stu-id="2a03a-159">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="2a03a-160">用户将看到 #BUSY!</span><span class="sxs-lookup"><span data-stu-id="2a03a-160">Users will see #BUSY!</span></span> <span data-ttu-id="2a03a-161">在返回结果之前。</span><span class="sxs-lookup"><span data-stu-id="2a03a-161">until a result is returned.</span></span> | <span data-ttu-id="2a03a-162">用户将看到 #BUSY!</span><span class="sxs-lookup"><span data-stu-id="2a03a-162">Users will see #BUSY!</span></span> <span data-ttu-id="2a03a-163">在返回结果之前。</span><span class="sxs-lookup"><span data-stu-id="2a03a-163">until a result is returned.</span></span> |
| <span data-ttu-id="2a03a-164">要求集</span><span class="sxs-lookup"><span data-stu-id="2a03a-164">Requirement sets</span></span> | <span data-ttu-id="2a03a-165">无</span><span class="sxs-lookup"><span data-stu-id="2a03a-165">N/A</span></span> | <span data-ttu-id="2a03a-166">Customfunctions.js 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="2a03a-166">CustomFunctions 1.1 and later</span></span> | <span data-ttu-id="2a03a-167">Customfunctions.js 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="2a03a-167">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="2a03a-168">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2a03a-168">See also</span></span>

- [<span data-ttu-id="2a03a-169">使 Excel 外接程序与现有 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="2a03a-169">Make your Excel add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="2a03a-170">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="2a03a-170">Custom functions best practices</span></span>](custom-functions-best-practices.md)
- [<span data-ttu-id="2a03a-171">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="2a03a-171">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
