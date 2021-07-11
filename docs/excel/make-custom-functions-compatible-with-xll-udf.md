---
title: 使用 XLL 用户定义函数扩展自定义函数
description: 启用与Excel等效功能的 XLL 用户定义函数的兼容性
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 33c7ee9309196d627520b37a02d5a1bca44cb767
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349390"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a><span data-ttu-id="b1781-103">使用 XLL 用户定义函数扩展自定义函数</span><span class="sxs-lookup"><span data-stu-id="b1781-103">Extend custom functions with XLL user-defined functions</span></span>

<span data-ttu-id="b1781-104">如果您已有 Excel XLL，可以在 Excel 外接程序中生成等效的自定义函数，以将解决方案功能扩展到其他平台（如联机或 Mac 上）。</span><span class="sxs-lookup"><span data-stu-id="b1781-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Excel add-in to extend your solution features to other platforms such as online or on a Mac.</span></span> <span data-ttu-id="b1781-105">但是Excel加载项并不具有 XLL 中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="b1781-105">However, Excel add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="b1781-106">根据您的解决方案使用的功能，XLL 可能会提供比 Excel 中的 Excel 外接程序自定义函数更好的Excel体验Windows。</span><span class="sxs-lookup"><span data-stu-id="b1781-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Excel add-in custom functions in Excel on Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="b1781-107">连接到订阅时，以下平台支持 COM 加载项和 XLL UDF Microsoft 365兼容性。</span><span class="sxs-lookup"><span data-stu-id="b1781-107">COM add-in and XLL UDF compatibility is supported by the following platforms, when connected to a Microsoft 365 subscription.</span></span>
>
> - <span data-ttu-id="b1781-108">Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="b1781-108">Excel on the web</span></span>
> - <span data-ttu-id="b1781-109">Excel版本Windows (版本 1904 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="b1781-109">Excel on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="b1781-110">Excel Mac (版本 13.329 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="b1781-110">Excel on Mac (version 13.329 or later)</span></span>
>
> <span data-ttu-id="b1781-111">若要在加载项内使用 COM 加载项和 XLL UDF Excel web 版，请使用你的 Microsoft 365 订阅或 Microsoft[帐户登录](https://account.microsoft.com/account)。</span><span class="sxs-lookup"><span data-stu-id="b1781-111">To use COM add-in and XLL UDF compatibility within Excel on the web, login by using either your Microsoft 365 subscription or a [Microsoft account](https://account.microsoft.com/account).</span></span> <span data-ttu-id="b1781-112">如果你还没有免费订阅，Microsoft 365开发人员计划，获得为期 90 天的免费可续订 Microsoft 365[订阅Microsoft 365订阅](https://developer.microsoft.com/office/dev-program)。</span><span class="sxs-lookup"><span data-stu-id="b1781-112">If you don't already have a Microsoft 365 subscription, you can a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="b1781-113">在清单中指定等效的 XLL</span><span class="sxs-lookup"><span data-stu-id="b1781-113">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="b1781-114">若要启用与现有 XLL 的兼容性，请标识加载项清单中的等效 XLL Excel XLL。</span><span class="sxs-lookup"><span data-stu-id="b1781-114">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Excel add-in.</span></span> <span data-ttu-id="b1781-115">Excel在加载项上运行时，Excel使用 XLL 函数，而不是Windows。</span><span class="sxs-lookup"><span data-stu-id="b1781-115">Excel will then use the XLL's functions instead of your Excel add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="b1781-116">若要为自定义函数设置等效的 XLL，请 `FileName` 指定 XLL 的 。</span><span class="sxs-lookup"><span data-stu-id="b1781-116">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="b1781-117">当用户使用 XLL 中的函数打开工作簿时，Excel函数转换为兼容函数。</span><span class="sxs-lookup"><span data-stu-id="b1781-117">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="b1781-118">然后，在 Windows 上的 Excel 中打开工作簿时，工作簿将使用 XLL，当联机打开或在 Mac 上打开时，它将使用 Excel 加载项中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="b1781-118">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Excel add-in when opened online or on a Mac.</span></span>

<span data-ttu-id="b1781-119">以下示例演示如何将 COM 加载项和 XLL 指定为等效项。</span><span class="sxs-lookup"><span data-stu-id="b1781-119">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="b1781-120">通常，您将同时指定这两者。</span><span class="sxs-lookup"><span data-stu-id="b1781-120">Often you will specify both.</span></span> <span data-ttu-id="b1781-121">为完整，此示例在上下文中显示这两者。</span><span class="sxs-lookup"><span data-stu-id="b1781-121">For completeness, this example shows both in context.</span></span> <span data-ttu-id="b1781-122">它们分别由它们 `ProgId` 和 `FileName` 标识。</span><span class="sxs-lookup"><span data-stu-id="b1781-122">They are identified by their `ProgId` and `FileName` respectively.</span></span> <span data-ttu-id="b1781-123">`EquivalentAddins`元素必须紧接在结束标记 `VersionOverrides` 的之前。</span><span class="sxs-lookup"><span data-stu-id="b1781-123">The `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span> <span data-ttu-id="b1781-124">有关 COM 加载项兼容性的详细信息，请参阅使Office[加载项与现有 COM](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)加载项兼容。</span><span class="sxs-lookup"><span data-stu-id="b1781-124">For more information on COM add-in compatibility, see [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

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
  </EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="b1781-125">如果加载项声明其自定义函数与 XLL 兼容，以后更改清单可能会破坏用户的工作簿，因为它将更改文件格式。</span><span class="sxs-lookup"><span data-stu-id="b1781-125">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user's workbook because it will change the file format.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="b1781-126">XLL 兼容函数的自定义函数行为</span><span class="sxs-lookup"><span data-stu-id="b1781-126">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="b1781-127">打开电子表格且有等效的加载项可用时，加载项的 XLL 函数将转换为 XLL 兼容的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="b1781-127">An add-in's XLL functions are converted to XLL compatible custom functions when a spreadsheet is opened and there is an equivalent add-in available.</span></span> <span data-ttu-id="b1781-128">下一次保存时，XLL 函数会以兼容模式写入文件，以便它们适用于 XLL 和 Excel 加载项自定义函数 (在其他平台上) 。</span><span class="sxs-lookup"><span data-stu-id="b1781-128">On the next save, the XLL functions are written to the file in a compatible mode so that they work with both the XLL and Excel add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="b1781-129">下表对 XLL 用户定义函数、XLL 兼容自定义函数和加载项自定义Excel功能进行比较。</span><span class="sxs-lookup"><span data-stu-id="b1781-129">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel add-in custom functions.</span></span>

|         |<span data-ttu-id="b1781-130">XLL 用户定义函数</span><span class="sxs-lookup"><span data-stu-id="b1781-130">XLL user-defined function</span></span> |<span data-ttu-id="b1781-131">XLL 兼容的自定义函数</span><span class="sxs-lookup"><span data-stu-id="b1781-131">XLL compatible custom functions</span></span> |<span data-ttu-id="b1781-132">Excel加载项自定义函数</span><span class="sxs-lookup"><span data-stu-id="b1781-132">Excel add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="b1781-133">**支持的平台**</span><span class="sxs-lookup"><span data-stu-id="b1781-133">**Supported platforms**</span></span> | <span data-ttu-id="b1781-134">Windows</span><span class="sxs-lookup"><span data-stu-id="b1781-134">Windows</span></span> | <span data-ttu-id="b1781-135">Windows、macOS、Web 浏览器</span><span class="sxs-lookup"><span data-stu-id="b1781-135">Windows, macOS, web browser</span></span> | <span data-ttu-id="b1781-136">Windows、macOS、Web 浏览器</span><span class="sxs-lookup"><span data-stu-id="b1781-136">Windows, macOS, web browser</span></span> |
| <span data-ttu-id="b1781-137">**支持的文件格式**</span><span class="sxs-lookup"><span data-stu-id="b1781-137">**Supported file formats**</span></span> | <span data-ttu-id="b1781-138">XLSX、XLSB、XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="b1781-138">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="b1781-139">XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="b1781-139">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="b1781-140">XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="b1781-140">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="b1781-141">**公式自动完成**</span><span class="sxs-lookup"><span data-stu-id="b1781-141">**Formula autocomplete**</span></span> | <span data-ttu-id="b1781-142">否</span><span class="sxs-lookup"><span data-stu-id="b1781-142">No</span></span> | <span data-ttu-id="b1781-143">是</span><span class="sxs-lookup"><span data-stu-id="b1781-143">Yes</span></span> | <span data-ttu-id="b1781-144">是</span><span class="sxs-lookup"><span data-stu-id="b1781-144">Yes</span></span> |
| <span data-ttu-id="b1781-145">**流式**</span><span class="sxs-lookup"><span data-stu-id="b1781-145">**Streaming**</span></span> | <span data-ttu-id="b1781-146">可通过 xlfRTD 和 XLL 回调实现。</span><span class="sxs-lookup"><span data-stu-id="b1781-146">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="b1781-147">是</span><span class="sxs-lookup"><span data-stu-id="b1781-147">Yes</span></span> | <span data-ttu-id="b1781-148">是</span><span class="sxs-lookup"><span data-stu-id="b1781-148">Yes</span></span> |
| <span data-ttu-id="b1781-149">**函数本地化**</span><span class="sxs-lookup"><span data-stu-id="b1781-149">**Localization of functions**</span></span> | <span data-ttu-id="b1781-150">否</span><span class="sxs-lookup"><span data-stu-id="b1781-150">No</span></span> | <span data-ttu-id="b1781-151">不正确。</span><span class="sxs-lookup"><span data-stu-id="b1781-151">No.</span></span> <span data-ttu-id="b1781-152">Name 和 ID 必须与现有的 XLL 函数匹配。</span><span class="sxs-lookup"><span data-stu-id="b1781-152">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="b1781-153">是</span><span class="sxs-lookup"><span data-stu-id="b1781-153">Yes</span></span> |
| <span data-ttu-id="b1781-154">**可变函数**</span><span class="sxs-lookup"><span data-stu-id="b1781-154">**Volatile functions**</span></span> | <span data-ttu-id="b1781-155">是</span><span class="sxs-lookup"><span data-stu-id="b1781-155">Yes</span></span> | <span data-ttu-id="b1781-156">是</span><span class="sxs-lookup"><span data-stu-id="b1781-156">Yes</span></span> | <span data-ttu-id="b1781-157">是</span><span class="sxs-lookup"><span data-stu-id="b1781-157">Yes</span></span> |
| <span data-ttu-id="b1781-158">**多线程重新计算支持**</span><span class="sxs-lookup"><span data-stu-id="b1781-158">**Multi-threaded recalculation support**</span></span> | <span data-ttu-id="b1781-159">是</span><span class="sxs-lookup"><span data-stu-id="b1781-159">Yes</span></span> | <span data-ttu-id="b1781-160">是</span><span class="sxs-lookup"><span data-stu-id="b1781-160">Yes</span></span> | <span data-ttu-id="b1781-161">是</span><span class="sxs-lookup"><span data-stu-id="b1781-161">Yes</span></span> |
| <span data-ttu-id="b1781-162">**计算行为**</span><span class="sxs-lookup"><span data-stu-id="b1781-162">**Calculation behavior**</span></span> | <span data-ttu-id="b1781-163">无 UI。</span><span class="sxs-lookup"><span data-stu-id="b1781-163">No UI.</span></span> <span data-ttu-id="b1781-164">Excel计算期间可能无响应。</span><span class="sxs-lookup"><span data-stu-id="b1781-164">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="b1781-165">用户将看到#BUSY！</span><span class="sxs-lookup"><span data-stu-id="b1781-165">Users will see #BUSY!</span></span> <span data-ttu-id="b1781-166">直到返回结果。</span><span class="sxs-lookup"><span data-stu-id="b1781-166">until a result is returned.</span></span> | <span data-ttu-id="b1781-167">用户将看到#BUSY！</span><span class="sxs-lookup"><span data-stu-id="b1781-167">Users will see #BUSY!</span></span> <span data-ttu-id="b1781-168">直到返回结果。</span><span class="sxs-lookup"><span data-stu-id="b1781-168">until a result is returned.</span></span> |
| <span data-ttu-id="b1781-169">**要求集**</span><span class="sxs-lookup"><span data-stu-id="b1781-169">**Requirement sets**</span></span> | <span data-ttu-id="b1781-170">不适用</span><span class="sxs-lookup"><span data-stu-id="b1781-170">N/A</span></span> | <span data-ttu-id="b1781-171">CustomFunctions 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="b1781-171">CustomFunctions 1.1 and later</span></span> | <span data-ttu-id="b1781-172">CustomFunctions 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="b1781-172">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="b1781-173">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b1781-173">See also</span></span>

- [<span data-ttu-id="b1781-174">让 Office 加载项与现有 COM 加载项兼容</span><span class="sxs-lookup"><span data-stu-id="b1781-174">Make your Office Add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="b1781-175">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="b1781-175">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
