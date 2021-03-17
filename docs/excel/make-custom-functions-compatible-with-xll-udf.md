---
title: 使用 XLL 用户定义函数扩展自定义函数
description: 启用与 Excel XLL 用户定义函数的兼容性，这些函数具有与自定义函数等效的功能
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 32146e7eebb963e8d800b619ef052457e40f2ac6
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836814"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a><span data-ttu-id="ef63e-103">使用 XLL 用户定义函数扩展自定义函数</span><span class="sxs-lookup"><span data-stu-id="ef63e-103">Extend custom functions with XLL user-defined functions</span></span>

<span data-ttu-id="ef63e-104">如果已有 Excel XLL，可以在 Excel 加载项中生成等效自定义函数，以将解决方案功能扩展到其他平台（如联机平台或 Mac 平台）。</span><span class="sxs-lookup"><span data-stu-id="ef63e-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Excel add-in to extend your solution features to other platforms such as online or on a Mac.</span></span> <span data-ttu-id="ef63e-105">但是，Excel 加载项并不具有 XLL 中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="ef63e-105">However, Excel add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="ef63e-106">根据您的解决方案使用的功能，XLL 可能会提供比 Windows 上的 Excel 中的 Excel 外接程序自定义函数更好的体验。</span><span class="sxs-lookup"><span data-stu-id="ef63e-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Excel add-in custom functions in Excel on Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="ef63e-107">连接到 Microsoft 365 订阅时，以下平台支持 COM 加载项和 XLL UDF 兼容性：</span><span class="sxs-lookup"><span data-stu-id="ef63e-107">COM add-in and XLL UDF compatibility is supported by the following platforms, when connected to a Microsoft 365 subscription:</span></span>
>
> - <span data-ttu-id="ef63e-108">Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="ef63e-108">Excel on the web</span></span>
> - <span data-ttu-id="ef63e-109">Windows 版 Excel (版本 1904 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="ef63e-109">Excel on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="ef63e-110">Mac 版 Excel (版本 13.329 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="ef63e-110">Excel on Mac (version 13.329 or later)</span></span>
>
> <span data-ttu-id="ef63e-111">若要在 Excel 网页版内使用 COM 加载项和 XLL UDF 兼容性，请使用 Microsoft 365 订阅或 Microsoft 帐户 [登录](https://account.microsoft.com/account)。</span><span class="sxs-lookup"><span data-stu-id="ef63e-111">To use COM add-in and XLL UDF compatibility within Excel on the web, login by using either your Microsoft 365 subscription or a [Microsoft account](https://account.microsoft.com/account).</span></span> <span data-ttu-id="ef63e-112">如果你还没有 Microsoft 365 订阅，则可以通过加入 Microsoft 365 开发人员计划获得为期 90 天的免费可续订 [Microsoft 365 订阅](https://developer.microsoft.com/office/dev-program)。</span><span class="sxs-lookup"><span data-stu-id="ef63e-112">If you don't already have a Microsoft 365 subscription, you can a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="ef63e-113">在清单中指定等效的 XLL</span><span class="sxs-lookup"><span data-stu-id="ef63e-113">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="ef63e-114">若要启用与现有 XLL 的兼容性，请确定 Excel 加载项清单中的等效 XLL。</span><span class="sxs-lookup"><span data-stu-id="ef63e-114">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Excel add-in.</span></span> <span data-ttu-id="ef63e-115">然后，当在 Windows 上运行时，Excel 将使用 XLL 函数，而不是 Excel 加载项自定义函数。</span><span class="sxs-lookup"><span data-stu-id="ef63e-115">Excel will then use the XLL's functions instead of your Excel add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="ef63e-116">若要为自定义函数设置等效的 XLL，请 `FileName` 指定 XLL 的 。</span><span class="sxs-lookup"><span data-stu-id="ef63e-116">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="ef63e-117">当用户使用 XLL 中的函数打开工作簿时，Excel 会将函数转换为兼容函数。</span><span class="sxs-lookup"><span data-stu-id="ef63e-117">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="ef63e-118">然后，在 Windows 上的 Excel 中打开工作簿时，工作簿将使用 XLL，当联机打开或在 Mac 上打开时，它将使用 Excel 加载项中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="ef63e-118">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Excel add-in when opened online or on a Mac.</span></span>

<span data-ttu-id="ef63e-119">以下示例演示如何将 COM 加载项和 XLL 指定为等效项。</span><span class="sxs-lookup"><span data-stu-id="ef63e-119">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="ef63e-120">通常，您将同时指定这两者。</span><span class="sxs-lookup"><span data-stu-id="ef63e-120">Often you will specify both.</span></span> <span data-ttu-id="ef63e-121">为完整，此示例在上下文中显示这两者。</span><span class="sxs-lookup"><span data-stu-id="ef63e-121">For completeness, this example shows both in context.</span></span> <span data-ttu-id="ef63e-122">它们分别由它们 `ProgId` 和 `FileName` 标识。</span><span class="sxs-lookup"><span data-stu-id="ef63e-122">They are identified by their `ProgId` and `FileName` respectively.</span></span> <span data-ttu-id="ef63e-123">`EquivalentAddins`元素必须紧接在结束标记 `VersionOverrides` 的之前。</span><span class="sxs-lookup"><span data-stu-id="ef63e-123">The `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span> <span data-ttu-id="ef63e-124">有关 COM 加载项兼容性的详细信息，请参阅使 [Office](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)加载项与现有 COM 加载项兼容。</span><span class="sxs-lookup"><span data-stu-id="ef63e-124">For more information on COM add-in compatibility, see [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

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
> <span data-ttu-id="ef63e-125">如果加载项声明其自定义函数与 XLL 兼容，以后更改清单可能会破坏用户的工作簿，因为它将更改文件格式。</span><span class="sxs-lookup"><span data-stu-id="ef63e-125">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user's workbook because it will change the file format.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="ef63e-126">XLL 兼容函数的自定义函数行为</span><span class="sxs-lookup"><span data-stu-id="ef63e-126">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="ef63e-127">打开电子表格且有等效的加载项可用时，加载项的 XLL 函数将转换为 XLL 兼容的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="ef63e-127">An add-in's XLL functions are converted to XLL compatible custom functions when a spreadsheet is opened and there is an equivalent add-in available.</span></span> <span data-ttu-id="ef63e-128">下一次保存时，XLL 函数会以兼容模式写入文件，以便它们在其他平台上使用 XLL 和 Excel (自定义函数) 。</span><span class="sxs-lookup"><span data-stu-id="ef63e-128">On the next save, the XLL functions are written to the file in a compatible mode so that they work with both the XLL and Excel add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="ef63e-129">下表比较了 XLL 用户定义函数、XLL 兼容自定义函数和 Excel 加载项自定义函数之间的功能。</span><span class="sxs-lookup"><span data-stu-id="ef63e-129">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel add-in custom functions.</span></span>

|         |<span data-ttu-id="ef63e-130">XLL 用户定义函数</span><span class="sxs-lookup"><span data-stu-id="ef63e-130">XLL user-defined function</span></span> |<span data-ttu-id="ef63e-131">XLL 兼容的自定义函数</span><span class="sxs-lookup"><span data-stu-id="ef63e-131">XLL compatible custom functions</span></span> |<span data-ttu-id="ef63e-132">Excel 加载项自定义函数</span><span class="sxs-lookup"><span data-stu-id="ef63e-132">Excel add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="ef63e-133">**支持的平台**</span><span class="sxs-lookup"><span data-stu-id="ef63e-133">**Supported platforms**</span></span> | <span data-ttu-id="ef63e-134">Windows</span><span class="sxs-lookup"><span data-stu-id="ef63e-134">Windows</span></span> | <span data-ttu-id="ef63e-135">Windows、macOS、Web 浏览器</span><span class="sxs-lookup"><span data-stu-id="ef63e-135">Windows, macOS, web browser</span></span> | <span data-ttu-id="ef63e-136">Windows、macOS、Web 浏览器</span><span class="sxs-lookup"><span data-stu-id="ef63e-136">Windows, macOS, web browser</span></span> |
| <span data-ttu-id="ef63e-137">**支持的文件格式**</span><span class="sxs-lookup"><span data-stu-id="ef63e-137">**Supported file formats**</span></span> | <span data-ttu-id="ef63e-138">XLSX、XLSB、XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="ef63e-138">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="ef63e-139">XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="ef63e-139">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="ef63e-140">XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="ef63e-140">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="ef63e-141">**公式自动完成**</span><span class="sxs-lookup"><span data-stu-id="ef63e-141">**Formula autocomplete**</span></span> | <span data-ttu-id="ef63e-142">否</span><span class="sxs-lookup"><span data-stu-id="ef63e-142">No</span></span> | <span data-ttu-id="ef63e-143">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-143">Yes</span></span> | <span data-ttu-id="ef63e-144">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-144">Yes</span></span> |
| <span data-ttu-id="ef63e-145">**流式**</span><span class="sxs-lookup"><span data-stu-id="ef63e-145">**Streaming**</span></span> | <span data-ttu-id="ef63e-146">可通过 xlfRTD 和 XLL 回调实现。</span><span class="sxs-lookup"><span data-stu-id="ef63e-146">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="ef63e-147">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-147">Yes</span></span> | <span data-ttu-id="ef63e-148">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-148">Yes</span></span> |
| <span data-ttu-id="ef63e-149">**函数本地化**</span><span class="sxs-lookup"><span data-stu-id="ef63e-149">**Localization of functions**</span></span> | <span data-ttu-id="ef63e-150">否</span><span class="sxs-lookup"><span data-stu-id="ef63e-150">No</span></span> | <span data-ttu-id="ef63e-151">不正确。</span><span class="sxs-lookup"><span data-stu-id="ef63e-151">No.</span></span> <span data-ttu-id="ef63e-152">Name 和 ID 必须与现有的 XLL 函数匹配。</span><span class="sxs-lookup"><span data-stu-id="ef63e-152">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="ef63e-153">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-153">Yes</span></span> |
| <span data-ttu-id="ef63e-154">**可变函数**</span><span class="sxs-lookup"><span data-stu-id="ef63e-154">**Volatile functions**</span></span> | <span data-ttu-id="ef63e-155">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-155">Yes</span></span> | <span data-ttu-id="ef63e-156">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-156">Yes</span></span> | <span data-ttu-id="ef63e-157">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-157">Yes</span></span> |
| <span data-ttu-id="ef63e-158">**多线程重新计算支持**</span><span class="sxs-lookup"><span data-stu-id="ef63e-158">**Multi-threaded recalculation support**</span></span> | <span data-ttu-id="ef63e-159">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-159">Yes</span></span> | <span data-ttu-id="ef63e-160">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-160">Yes</span></span> | <span data-ttu-id="ef63e-161">是</span><span class="sxs-lookup"><span data-stu-id="ef63e-161">Yes</span></span> |
| <span data-ttu-id="ef63e-162">**计算行为**</span><span class="sxs-lookup"><span data-stu-id="ef63e-162">**Calculation behavior**</span></span> | <span data-ttu-id="ef63e-163">无 UI。</span><span class="sxs-lookup"><span data-stu-id="ef63e-163">No UI.</span></span> <span data-ttu-id="ef63e-164">Excel 在计算过程中可能无响应。</span><span class="sxs-lookup"><span data-stu-id="ef63e-164">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="ef63e-165">用户将看到#BUSY！</span><span class="sxs-lookup"><span data-stu-id="ef63e-165">Users will see #BUSY!</span></span> <span data-ttu-id="ef63e-166">直到返回结果。</span><span class="sxs-lookup"><span data-stu-id="ef63e-166">until a result is returned.</span></span> | <span data-ttu-id="ef63e-167">用户将看到#BUSY！</span><span class="sxs-lookup"><span data-stu-id="ef63e-167">Users will see #BUSY!</span></span> <span data-ttu-id="ef63e-168">直到返回结果。</span><span class="sxs-lookup"><span data-stu-id="ef63e-168">until a result is returned.</span></span> |
| <span data-ttu-id="ef63e-169">**要求集**</span><span class="sxs-lookup"><span data-stu-id="ef63e-169">**Requirement sets**</span></span> | <span data-ttu-id="ef63e-170">不适用</span><span class="sxs-lookup"><span data-stu-id="ef63e-170">N/A</span></span> | <span data-ttu-id="ef63e-171">CustomFunctions 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="ef63e-171">CustomFunctions 1.1 and later</span></span> | <span data-ttu-id="ef63e-172">CustomFunctions 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="ef63e-172">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="ef63e-173">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ef63e-173">See also</span></span>

- [<span data-ttu-id="ef63e-174">让 Office 加载项与现有 COM 加载项兼容</span><span class="sxs-lookup"><span data-stu-id="ef63e-174">Make your Office Add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="ef63e-175">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="ef63e-175">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
