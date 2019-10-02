---
title: 使用 XLL 用户定义的函数扩展自定义函数
description: 启用与自定义函数具有等效功能的 Excel XLL 用户定义函数的兼容性
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: a0a98dab1ec046151d2dd0d80a4a3a4542654574
ms.sourcegitcommit: 528577145b2cf0a42bc64c56145d661c4d019fb8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/02/2019
ms.locfileid: "37353879"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a><span data-ttu-id="a6760-103">使用 XLL 用户定义的函数扩展自定义函数</span><span class="sxs-lookup"><span data-stu-id="a6760-103">Extend custom functions with XLL user-defined functions</span></span>

<span data-ttu-id="a6760-104">如果您有现有的 Excel Xll，则可以在 Excel 外接程序中构建等效的自定义函数，以将解决方案功能扩展到其他平台（如 online 或 macOS）。</span><span class="sxs-lookup"><span data-stu-id="a6760-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="a6760-105">但是，Excel 外接程序没有在 Xll 中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="a6760-105">However, Excel add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="a6760-106">根据您的解决方案使用的功能，XLL 可以提供比 excel 在 Windows 上运行的 Excel 外接程序自定义函数更好的体验。</span><span class="sxs-lookup"><span data-stu-id="a6760-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Excel add-in custom functions in Excel on Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="a6760-107">当连接到 Office 365 订阅时，以下平台支持 COM 加载项和 XLL UDF 兼容性：</span><span class="sxs-lookup"><span data-stu-id="a6760-107">COM add-in and XLL UDF compatibility is supported by the following platforms, when connected to an Office 365 subscription:</span></span>
> - <span data-ttu-id="a6760-108">在 web 上的 Excel</span><span class="sxs-lookup"><span data-stu-id="a6760-108">Excel on the web</span></span>
> - <span data-ttu-id="a6760-109">Windows 上的 Excel （版本1904或更高版本）</span><span class="sxs-lookup"><span data-stu-id="a6760-109">Excel on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="a6760-110">Excel for Mac （版本13.329 或更高版本）</span><span class="sxs-lookup"><span data-stu-id="a6760-110">Excel on Mac (version 13.329 or later)</span></span>
> 
> <span data-ttu-id="a6760-111">若要在 web 上的 Excel 中使用 COM 加载项并 XLL UDF 兼容性，请使用 Office 365 订阅或[Microsoft 帐户](https://account.microsoft.com/account)登录。</span><span class="sxs-lookup"><span data-stu-id="a6760-111">To use COM add-in and XLL UDF compatibility within Excel on the web, login by using either your Office 365 subscription or a [Microsoft account](https://account.microsoft.com/account).</span></span> <span data-ttu-id="a6760-112">如果还没有 Office 365 订阅，可以通过加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个订阅。</span><span class="sxs-lookup"><span data-stu-id="a6760-112">If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="a6760-113">在清单中指定等效 XLL</span><span class="sxs-lookup"><span data-stu-id="a6760-113">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="a6760-114">若要启用与现有 XLL 的兼容性，请在您的 Excel 外接程序清单中标识等效 XLL。</span><span class="sxs-lookup"><span data-stu-id="a6760-114">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Excel add-in.</span></span> <span data-ttu-id="a6760-115">在 Windows 上运行时，Excel 将使用 XLL 的函数而不是 Excel 加载项自定义函数。</span><span class="sxs-lookup"><span data-stu-id="a6760-115">Then Excel will use the XLL's functions instead of your Excel add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="a6760-116">若要设置自定义函数的等效 XLL，请指定`FileName` XLL 的。</span><span class="sxs-lookup"><span data-stu-id="a6760-116">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="a6760-117">当用户使用 XLL 中的函数打开工作簿时，Excel 会将函数转换为兼容函数。</span><span class="sxs-lookup"><span data-stu-id="a6760-117">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="a6760-118">在 Windows 上的 Excel 中打开时，工作簿将使用 XLL，并且在联机或在 macOS 中打开时，它将使用 Excel 外接程序中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="a6760-118">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Excel add-in when opened online or on macOS.</span></span>

<span data-ttu-id="a6760-119">下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。</span><span class="sxs-lookup"><span data-stu-id="a6760-119">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="a6760-120">通常，出于完整性的考虑，这两个示例都会在上下文中显示这两个示例。</span><span class="sxs-lookup"><span data-stu-id="a6760-120">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="a6760-121">它们`ProgId` `FileName`分别由各自标识。</span><span class="sxs-lookup"><span data-stu-id="a6760-121">They are identified by their `ProgId` and `FileName` respectively.</span></span> <span data-ttu-id="a6760-122">`EquivalentAddins`元素必须紧跟在结束`VersionOverrides`标记之前。</span><span class="sxs-lookup"><span data-stu-id="a6760-122">The `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span> <span data-ttu-id="a6760-123">有关 COM 加载项兼容性的详细信息，请参阅[使您的 Excel 外接程序与现有的 com 外](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)接程序兼容。</span><span class="sxs-lookup"><span data-stu-id="a6760-123">For more information on COM add-in compatibility, see [Make your Excel add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>

    <EquivalentAddin>
      <FileName>contosofunctions.xll</FileName>
      <Type>XLL</Type>
    </EquivalentAddin>
  <EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="a6760-124">如果外接程序声明其自定义函数是 XLL 兼容的，则稍后更改清单可能会破坏用户的工作簿，因为它会更改文件格式。</span><span class="sxs-lookup"><span data-stu-id="a6760-124">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.</span></span>

## <a name="excel-add-in-updates"></a><span data-ttu-id="a6760-125">Excel 加载项更新</span><span class="sxs-lookup"><span data-stu-id="a6760-125">Excel add-in updates</span></span>

<span data-ttu-id="a6760-126">为 Excel 加载项指定等效 XLL 后，Excel 将停止处理 Excel 加载项的更新。</span><span class="sxs-lookup"><span data-stu-id="a6760-126">Once you specify an equivalent XLL for your Excel add-in, Excel stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="a6760-127">用户必须卸载 XLL 才能获取 Excel 外接程序的最新更新。</span><span class="sxs-lookup"><span data-stu-id="a6760-127">The user must uninstall the XLL in order to get the latest updates for the Excel add-in.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="a6760-128">XLL 兼容函数的自定义函数行为</span><span class="sxs-lookup"><span data-stu-id="a6760-128">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="a6760-129">如果打开的电子表格中包含的 XLL 函数也有等效的加载项，则 XLL 的函数将转换为 XLL 兼容的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="a6760-129">When a spreadsheet is opened that contains XLL functions for which there is also an equivalent add-in, the XLL's functions are converted to XLL compatible custom functions.</span></span> <span data-ttu-id="a6760-130">在下一次保存时，它们将在兼容模式下写入文件，以便它们使用 XLL 和 Excel 外接程序自定义函数（当在其他平台上）。</span><span class="sxs-lookup"><span data-stu-id="a6760-130">On the next save, they are written to the file in a compatible mode so that they work with both the XLL and Excel add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="a6760-131">下表比较了 XLL 用户定义函数、XLL 兼容的自定义函数和 Excel 加载项自定义函数之间的功能。</span><span class="sxs-lookup"><span data-stu-id="a6760-131">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel add-in custom functions.</span></span>

|         |<span data-ttu-id="a6760-132">XLL 用户定义的函数</span><span class="sxs-lookup"><span data-stu-id="a6760-132">XLL user-defined function</span></span> |<span data-ttu-id="a6760-133">XLL 兼容的自定义函数</span><span class="sxs-lookup"><span data-stu-id="a6760-133">XLL compatible custom functions</span></span> |<span data-ttu-id="a6760-134">Excel 加载项自定义函数</span><span class="sxs-lookup"><span data-stu-id="a6760-134">Excel add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="a6760-135">支持的平台</span><span class="sxs-lookup"><span data-stu-id="a6760-135">Supported platforms</span></span> | <span data-ttu-id="a6760-136">Windows</span><span class="sxs-lookup"><span data-stu-id="a6760-136">Windows</span></span> | <span data-ttu-id="a6760-137">Windows、macOS、web 浏览器</span><span class="sxs-lookup"><span data-stu-id="a6760-137">Windows, macOS, web browser</span></span> | <span data-ttu-id="a6760-138">Windows、macOS、web 浏览器</span><span class="sxs-lookup"><span data-stu-id="a6760-138">Windows, macOS, web browser</span></span> |
| <span data-ttu-id="a6760-139">支持的文件格式</span><span class="sxs-lookup"><span data-stu-id="a6760-139">Supported file formats</span></span> | <span data-ttu-id="a6760-140">.XLSX、XLSB、XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="a6760-140">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="a6760-141">.XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="a6760-141">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="a6760-142">.XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="a6760-142">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="a6760-143">公式自动完成</span><span class="sxs-lookup"><span data-stu-id="a6760-143">Formula autocomplete</span></span> | <span data-ttu-id="a6760-144">否</span><span class="sxs-lookup"><span data-stu-id="a6760-144">No</span></span> | <span data-ttu-id="a6760-145">可访问</span><span class="sxs-lookup"><span data-stu-id="a6760-145">Yes</span></span> | <span data-ttu-id="a6760-146">是</span><span class="sxs-lookup"><span data-stu-id="a6760-146">Yes</span></span> |
| <span data-ttu-id="a6760-147">媒体</span><span class="sxs-lookup"><span data-stu-id="a6760-147">Streaming</span></span> | <span data-ttu-id="a6760-148">可通过 xlfRTD 和 XLL 回调实现。</span><span class="sxs-lookup"><span data-stu-id="a6760-148">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="a6760-149">否</span><span class="sxs-lookup"><span data-stu-id="a6760-149">No</span></span> | <span data-ttu-id="a6760-150">可访问</span><span class="sxs-lookup"><span data-stu-id="a6760-150">Yes</span></span> |
| <span data-ttu-id="a6760-151">函数的本地化</span><span class="sxs-lookup"><span data-stu-id="a6760-151">Localization of functions</span></span> | <span data-ttu-id="a6760-152">否</span><span class="sxs-lookup"><span data-stu-id="a6760-152">No</span></span> | <span data-ttu-id="a6760-153">否。</span><span class="sxs-lookup"><span data-stu-id="a6760-153">No.</span></span> <span data-ttu-id="a6760-154">名称和 ID 必须与现有 XLL 的函数相匹配。</span><span class="sxs-lookup"><span data-stu-id="a6760-154">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="a6760-155">是</span><span class="sxs-lookup"><span data-stu-id="a6760-155">Yes</span></span> |
| <span data-ttu-id="a6760-156">可变函数</span><span class="sxs-lookup"><span data-stu-id="a6760-156">Volatile functions</span></span> | <span data-ttu-id="a6760-157">是</span><span class="sxs-lookup"><span data-stu-id="a6760-157">Yes</span></span> | <span data-ttu-id="a6760-158">是</span><span class="sxs-lookup"><span data-stu-id="a6760-158">Yes</span></span> | <span data-ttu-id="a6760-159">是</span><span class="sxs-lookup"><span data-stu-id="a6760-159">Yes</span></span> |
| <span data-ttu-id="a6760-160">多线程重新计算支持</span><span class="sxs-lookup"><span data-stu-id="a6760-160">Multi-threaded recalculation support</span></span> | <span data-ttu-id="a6760-161">是</span><span class="sxs-lookup"><span data-stu-id="a6760-161">Yes</span></span> | <span data-ttu-id="a6760-162">是</span><span class="sxs-lookup"><span data-stu-id="a6760-162">Yes</span></span> | <span data-ttu-id="a6760-163">是</span><span class="sxs-lookup"><span data-stu-id="a6760-163">Yes</span></span> |
| <span data-ttu-id="a6760-164">计算行为</span><span class="sxs-lookup"><span data-stu-id="a6760-164">Calculation behavior</span></span> | <span data-ttu-id="a6760-165">无 UI。</span><span class="sxs-lookup"><span data-stu-id="a6760-165">No UI.</span></span> <span data-ttu-id="a6760-166">在计算过程中，Excel 可能会无响应。</span><span class="sxs-lookup"><span data-stu-id="a6760-166">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="a6760-167">用户将看到 #BUSY！</span><span class="sxs-lookup"><span data-stu-id="a6760-167">Users will see #BUSY!</span></span> <span data-ttu-id="a6760-168">在返回结果之前。</span><span class="sxs-lookup"><span data-stu-id="a6760-168">until a result is returned.</span></span> | <span data-ttu-id="a6760-169">用户将看到 #BUSY！</span><span class="sxs-lookup"><span data-stu-id="a6760-169">Users will see #BUSY!</span></span> <span data-ttu-id="a6760-170">在返回结果之前。</span><span class="sxs-lookup"><span data-stu-id="a6760-170">until a result is returned.</span></span> |
| <span data-ttu-id="a6760-171">要求集</span><span class="sxs-lookup"><span data-stu-id="a6760-171">Requirement sets</span></span> | <span data-ttu-id="a6760-172">不适用</span><span class="sxs-lookup"><span data-stu-id="a6760-172">N/A</span></span> | <span data-ttu-id="a6760-173">Customfunctions.js 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="a6760-173">CustomFunctions 1.1 and later</span></span> | <span data-ttu-id="a6760-174">Customfunctions.js 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="a6760-174">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="a6760-175">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a6760-175">See also</span></span>

- [<span data-ttu-id="a6760-176">使 Excel 外接程序与现有 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="a6760-176">Make your Excel add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="a6760-177">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="a6760-177">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
