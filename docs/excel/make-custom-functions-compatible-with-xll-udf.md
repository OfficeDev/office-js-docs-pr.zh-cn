---
title: 使自定义函数与 XLL 用户定义的函数兼容
description: 启用与自定义函数具有等效功能的 Excel XLL 用户定义函数的兼容性
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 09914e040c1721dd8b9e91952e5814e7a6b914e5
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356849"
---
# <a name="make-your-custom-functions-compatible-with-xll-user-defined-functions"></a><span data-ttu-id="6d7de-103">使自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="6d7de-103">Make your custom functions compatible with XLL user-defined functions</span></span>

<span data-ttu-id="6d7de-104">如果您有现有的 Excel xll, 则可以在 Office 外接程序中构建等效的自定义函数, 以将解决方案功能扩展到其他平台 (如 online 或 macOS)。</span><span class="sxs-lookup"><span data-stu-id="6d7de-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Office Add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="6d7de-105">但是, Office 外接程序没有 xll 中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="6d7de-105">However, Office Add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="6d7de-106">根据您的解决方案使用的功能, XLL 可能比 Excel for Windows 中的 Office 外接程序自定义函数提供更好的体验。</span><span class="sxs-lookup"><span data-stu-id="6d7de-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Office Add-in custom functions on Excel for Windows.</span></span>

<span data-ttu-id="6d7de-107">您可以配置 Office 外接程序, 以便在用户计算机上已安装等效 XLL 时, Excel 将运行 XLL 而不是 Office 外接程序自定义函数。</span><span class="sxs-lookup"><span data-stu-id="6d7de-107">You can configure your Office Add-in so that when an equivalent XLL is already installed on the user's computer, Excel runs the XLL instead of your Office Add-in custom functions.</span></span> <span data-ttu-id="6d7de-108">xll 被称作等效操作, 因为 Excel 将在 XLL 和 Office 加载项自定义函数之间进行无缝转换, 具体取决于 Windows 上安装的功能。</span><span class="sxs-lookup"><span data-stu-id="6d7de-108">The XLL is called equivalent because Excel will seamlessly transition between the XLL and the Office Add-in custom functions depending on which is installed on Windows.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="6d7de-109">在清单中指定等效 XLL</span><span class="sxs-lookup"><span data-stu-id="6d7de-109">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="6d7de-110">若要启用与现有 XLL 的兼容性, 请在 Office 外接程序的清单中标识等效的 XLL。</span><span class="sxs-lookup"><span data-stu-id="6d7de-110">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Office Add-in.</span></span> <span data-ttu-id="6d7de-111">在 Windows 上运行时, Excel 将使用 XLL 的函数而不是 Office 外接程序自定义函数。</span><span class="sxs-lookup"><span data-stu-id="6d7de-111">Then Excel will use the XLL's functions instead of your Office Add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="6d7de-112">若要设置自定义函数的等效 XLL, 请指定`FileName` XLL 的。</span><span class="sxs-lookup"><span data-stu-id="6d7de-112">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="6d7de-113">当用户使用 XLL 中的函数打开工作簿时, Excel 会将函数转换为兼容函数。</span><span class="sxs-lookup"><span data-stu-id="6d7de-113">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="6d7de-114">在 Windows Excel 中打开时, 工作簿将使用 XLL, 并且在联机或在 macOS 中打开时, 它将使用 Office 外接程序中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="6d7de-114">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Office Add-in when opened online or on macOS.</span></span>

<span data-ttu-id="6d7de-115">下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。</span><span class="sxs-lookup"><span data-stu-id="6d7de-115">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="6d7de-116">通常, 出于完整性的考虑, 这两个示例都会在上下文中显示这两个示例。</span><span class="sxs-lookup"><span data-stu-id="6d7de-116">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="6d7de-117">它们`ProgID` `FileName`分别由各自标识。</span><span class="sxs-lookup"><span data-stu-id="6d7de-117">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="6d7de-118">有关 COM 加载项兼容性的详细信息, 请参阅[使 Office 外接程序与现有 COM 加载项兼容](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="6d7de-118">For more information on COM add-in compatibility, see [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

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
> <span data-ttu-id="6d7de-119">如果外接程序声明其自定义函数是 XLL 兼容的, 则稍后更改清单可能会破坏用户的工作簿, 因为它会更改文件格式。</span><span class="sxs-lookup"><span data-stu-id="6d7de-119">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.</span></span>

## <a name="office-add-in-updates"></a><span data-ttu-id="6d7de-120">Office 外接程序更新</span><span class="sxs-lookup"><span data-stu-id="6d7de-120">Office Add-in updates</span></span>

<span data-ttu-id="6d7de-121">为 office 外接程序指定等效 XLL 后, Excel 将停止处理 office 外接程序的更新。</span><span class="sxs-lookup"><span data-stu-id="6d7de-121">Once you specify an equivalent XLL for your Office Add-in, Excel stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="6d7de-122">用户必须卸载 XLL 才能获取 Office 外接程序的最新更新。</span><span class="sxs-lookup"><span data-stu-id="6d7de-122">The user must uninstall the XLL in order to get the latest updates for the Office Add-in.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="6d7de-123">XLL 兼容函数的自定义函数行为</span><span class="sxs-lookup"><span data-stu-id="6d7de-123">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="6d7de-124">如果打开的电子表格中包含的 xll 函数也有等效的加载项, 则 XLL 的函数将转换为 xll 兼容的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="6d7de-124">When a spreadsheet is opened that contains XLL functions for which there is also an equivalent add-in, the XLL's functions are converted to XLL compatible custom functions.</span></span> <span data-ttu-id="6d7de-125">在下一次保存时, 它们将在兼容模式下写入文件, 以便它们使用 XLL 和 Office 外接程序自定义函数 (当在其他平台上)。</span><span class="sxs-lookup"><span data-stu-id="6d7de-125">On the next save, they are written to the file in a compatible mode so that they work with both the XLL and Office Add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="6d7de-126">下表比较了 XLL 用户定义函数、XLL 兼容的自定义函数和 Office 加载项自定义函数之间的功能。</span><span class="sxs-lookup"><span data-stu-id="6d7de-126">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Office Add-in custom functions.</span></span>

|         |<span data-ttu-id="6d7de-127">XLL 用户定义的函数</span><span class="sxs-lookup"><span data-stu-id="6d7de-127">XLL user-defined function</span></span> |<span data-ttu-id="6d7de-128">XLL 兼容的自定义函数</span><span class="sxs-lookup"><span data-stu-id="6d7de-128">XLL compatible custom functions</span></span> |<span data-ttu-id="6d7de-129">Office 外接自定义函数</span><span class="sxs-lookup"><span data-stu-id="6d7de-129">Office Add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="6d7de-130">支持的平台</span><span class="sxs-lookup"><span data-stu-id="6d7de-130">Supported platforms</span></span> | <span data-ttu-id="6d7de-131">Windows</span><span class="sxs-lookup"><span data-stu-id="6d7de-131">Windows</span></span> | <span data-ttu-id="6d7de-132">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="6d7de-132">Windows, macOS, Excel online</span></span> | <span data-ttu-id="6d7de-133">Windows、macOS、Excel online</span><span class="sxs-lookup"><span data-stu-id="6d7de-133">Windows, macOS, Excel online</span></span> |
| <span data-ttu-id="6d7de-134">支持的文件格式</span><span class="sxs-lookup"><span data-stu-id="6d7de-134">Supported file formats</span></span> | <span data-ttu-id="6d7de-135">.XLSX、XLSB、XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="6d7de-135">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="6d7de-136">.XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="6d7de-136">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="6d7de-137">.XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="6d7de-137">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="6d7de-138">公式自动完成</span><span class="sxs-lookup"><span data-stu-id="6d7de-138">Formula autocomplete</span></span> | <span data-ttu-id="6d7de-139">否</span><span class="sxs-lookup"><span data-stu-id="6d7de-139">No</span></span> | <span data-ttu-id="6d7de-140">可访问</span><span class="sxs-lookup"><span data-stu-id="6d7de-140">Yes</span></span> | <span data-ttu-id="6d7de-141">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-141">Yes</span></span> |
| <span data-ttu-id="6d7de-142">媒体</span><span class="sxs-lookup"><span data-stu-id="6d7de-142">Streaming</span></span> | <span data-ttu-id="6d7de-143">可通过 xlfRTD 和 XLL 回调实现。</span><span class="sxs-lookup"><span data-stu-id="6d7de-143">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="6d7de-144">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-144">Yes</span></span> | <span data-ttu-id="6d7de-145">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-145">Yes</span></span> |
| <span data-ttu-id="6d7de-146">函数的本地化</span><span class="sxs-lookup"><span data-stu-id="6d7de-146">Localization of functions</span></span> | <span data-ttu-id="6d7de-147">否</span><span class="sxs-lookup"><span data-stu-id="6d7de-147">No</span></span> | <span data-ttu-id="6d7de-148">否。</span><span class="sxs-lookup"><span data-stu-id="6d7de-148">No.</span></span> <span data-ttu-id="6d7de-149">名称和 ID 必须与现有 XLL 的函数相匹配。</span><span class="sxs-lookup"><span data-stu-id="6d7de-149">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="6d7de-150">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-150">Yes</span></span> |
| <span data-ttu-id="6d7de-151">可变函数</span><span class="sxs-lookup"><span data-stu-id="6d7de-151">Volatile functions</span></span> | <span data-ttu-id="6d7de-152">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-152">Yes</span></span> | <span data-ttu-id="6d7de-153">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-153">Yes</span></span> | <span data-ttu-id="6d7de-154">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-154">Yes</span></span> |
| <span data-ttu-id="6d7de-155">多线程重新计算支持</span><span class="sxs-lookup"><span data-stu-id="6d7de-155">Multi-threaded recalculation support</span></span> | <span data-ttu-id="6d7de-156">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-156">Yes</span></span> | <span data-ttu-id="6d7de-157">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-157">Yes</span></span> | <span data-ttu-id="6d7de-158">是</span><span class="sxs-lookup"><span data-stu-id="6d7de-158">Yes</span></span> |
| <span data-ttu-id="6d7de-159">计算行为</span><span class="sxs-lookup"><span data-stu-id="6d7de-159">Calculation behavior</span></span> | <span data-ttu-id="6d7de-160">无 UI。</span><span class="sxs-lookup"><span data-stu-id="6d7de-160">No UI.</span></span> <span data-ttu-id="6d7de-161">在计算过程中, Excel 可能会无响应。</span><span class="sxs-lookup"><span data-stu-id="6d7de-161">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="6d7de-162">用户将看到 #BUSY!</span><span class="sxs-lookup"><span data-stu-id="6d7de-162">Users will see #BUSY!</span></span> <span data-ttu-id="6d7de-163">在返回结果之前。</span><span class="sxs-lookup"><span data-stu-id="6d7de-163">until a result is returned.</span></span> | <span data-ttu-id="6d7de-164">用户将看到 #BUSY!</span><span class="sxs-lookup"><span data-stu-id="6d7de-164">Users will see #BUSY!</span></span> <span data-ttu-id="6d7de-165">在返回结果之前。</span><span class="sxs-lookup"><span data-stu-id="6d7de-165">until a result is returned.</span></span> |
| <span data-ttu-id="6d7de-166">要求集</span><span class="sxs-lookup"><span data-stu-id="6d7de-166">Requirement sets</span></span> | <span data-ttu-id="6d7de-167">无</span><span class="sxs-lookup"><span data-stu-id="6d7de-167">N/A</span></span> | <span data-ttu-id="6d7de-168">仅 customfunctions.js 1。1</span><span class="sxs-lookup"><span data-stu-id="6d7de-168">CustomFunctions 1.1 only</span></span> | <span data-ttu-id="6d7de-169">customfunctions.js 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="6d7de-169">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="6d7de-170">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6d7de-170">See also</span></span>

- [<span data-ttu-id="6d7de-171">使您的 Office 外接程序与现有的 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="6d7de-171">Make your Office Add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="6d7de-172">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="6d7de-172">Custom functions best practices</span></span>](custom-functions-best-practices.md)
- [<span data-ttu-id="6d7de-173">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="6d7de-173">Custom functions changelog</span></span>](custom-functions-changelog.md)
- [<span data-ttu-id="6d7de-174">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="6d7de-174">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)