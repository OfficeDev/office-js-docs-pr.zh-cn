---
title: 使用 XLL 用户定义的函数扩展自定义函数
description: 启用与自定义函数具有等效功能的 Excel XLL 用户定义函数的兼容性
ms.date: 04/29/2020
localization_priority: Normal
ms.openlocfilehash: 06a776678c60a4b66494097385f304ddac9a4ce0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610359"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a><span data-ttu-id="13cdb-103">使用 XLL 用户定义的函数扩展自定义函数</span><span class="sxs-lookup"><span data-stu-id="13cdb-103">Extend custom functions with XLL user-defined functions</span></span>

<span data-ttu-id="13cdb-104">如果您有现有的 Excel Xll，则可以在 Excel 外接程序中构建等效的自定义函数，以将解决方案功能扩展到其他平台（如联机或 Mac）。</span><span class="sxs-lookup"><span data-stu-id="13cdb-104">If you have existing Excel XLLs, you can build equivalent custom functions in an Excel add-in to extend your solution features to other platforms such as online or on a Mac.</span></span> <span data-ttu-id="13cdb-105">但是，Excel 外接程序没有在 Xll 中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="13cdb-105">However, Excel add-ins don't have all of the functionality available in XLLs.</span></span> <span data-ttu-id="13cdb-106">根据您的解决方案使用的功能，XLL 可以提供比 excel 在 Windows 上运行的 Excel 外接程序自定义函数更好的体验。</span><span class="sxs-lookup"><span data-stu-id="13cdb-106">Depending on the functionality your solution uses, the XLL may provide a better experience than the Excel add-in custom functions in Excel on Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="13cdb-107">当连接到 Office 365 订阅时，以下平台支持 COM 加载项和 XLL UDF 兼容性：</span><span class="sxs-lookup"><span data-stu-id="13cdb-107">COM add-in and XLL UDF compatibility is supported by the following platforms, when connected to an Office 365 subscription:</span></span>
> - <span data-ttu-id="13cdb-108">Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="13cdb-108">Excel on the web</span></span>
> - <span data-ttu-id="13cdb-109">Windows 上的 Excel （版本1904或更高版本）</span><span class="sxs-lookup"><span data-stu-id="13cdb-109">Excel on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="13cdb-110">Excel for Mac （版本13.329 或更高版本）</span><span class="sxs-lookup"><span data-stu-id="13cdb-110">Excel on Mac (version 13.329 or later)</span></span>
> 
> <span data-ttu-id="13cdb-111">若要在 web 上的 Excel 中使用 COM 加载项并 XLL UDF 兼容性，请使用 Office 365 订阅或[Microsoft 帐户](https://account.microsoft.com/account)登录。</span><span class="sxs-lookup"><span data-stu-id="13cdb-111">To use COM add-in and XLL UDF compatibility within Excel on the web, login by using either your Office 365 subscription or a [Microsoft account](https://account.microsoft.com/account).</span></span> <span data-ttu-id="13cdb-112">如果还没有 Office 365 订阅，则可以通过加入[office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)免费的90天 renewable office 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="13cdb-112">If you don't already have an Office 365 subscription, you can a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="specify-equivalent-xll-in-the-manifest"></a><span data-ttu-id="13cdb-113">在清单中指定等效 XLL</span><span class="sxs-lookup"><span data-stu-id="13cdb-113">Specify equivalent XLL in the manifest</span></span>

<span data-ttu-id="13cdb-114">若要启用与现有 XLL 的兼容性，请在您的 Excel 外接程序清单中标识等效 XLL。</span><span class="sxs-lookup"><span data-stu-id="13cdb-114">To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Excel add-in.</span></span> <span data-ttu-id="13cdb-115">然后，在 Windows 上运行时，excel 将使用 XLL 的函数而不是 Excel 加载项自定义函数。</span><span class="sxs-lookup"><span data-stu-id="13cdb-115">Excel will then use the XLL's functions instead of your Excel add-in custom functions when running on Windows.</span></span>

<span data-ttu-id="13cdb-116">若要设置自定义函数的等效 XLL，请指定 `FileName` XLL 的。</span><span class="sxs-lookup"><span data-stu-id="13cdb-116">To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL.</span></span> <span data-ttu-id="13cdb-117">当用户使用 XLL 中的函数打开工作簿时，Excel 会将函数转换为兼容函数。</span><span class="sxs-lookup"><span data-stu-id="13cdb-117">When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions.</span></span> <span data-ttu-id="13cdb-118">在 Windows 上的 Excel 中打开时，工作簿将使用 XLL，并且在联机或在 Mac 上打开时，它将使用 Excel 加载项中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="13cdb-118">The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Excel add-in when opened online or on a Mac.</span></span>

<span data-ttu-id="13cdb-119">下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。</span><span class="sxs-lookup"><span data-stu-id="13cdb-119">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="13cdb-120">通常会同时指定这两个。</span><span class="sxs-lookup"><span data-stu-id="13cdb-120">Often you will specify both.</span></span> <span data-ttu-id="13cdb-121">为了实现完整性，本示例同时显示了上下文中的内容。</span><span class="sxs-lookup"><span data-stu-id="13cdb-121">For completeness, this example shows both in context.</span></span> <span data-ttu-id="13cdb-122">它们分别由各自标识 `ProgId` `FileName` 。</span><span class="sxs-lookup"><span data-stu-id="13cdb-122">They are identified by their `ProgId` and `FileName` respectively.</span></span> <span data-ttu-id="13cdb-123">`EquivalentAddins`元素必须紧跟在结束 `VersionOverrides` 标记之前。</span><span class="sxs-lookup"><span data-stu-id="13cdb-123">The `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span> <span data-ttu-id="13cdb-124">有关 COM 加载项兼容性的详细信息，请参阅[使您的 Excel 外接程序与现有的 com 外](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)接程序兼容。</span><span class="sxs-lookup"><span data-stu-id="13cdb-124">For more information on COM add-in compatibility, see [Make your Excel add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).</span></span>

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
> <span data-ttu-id="13cdb-125">如果外接程序声明其自定义函数是 XLL 兼容的，则稍后更改清单可能会破坏用户的工作簿，因为它会更改文件格式。</span><span class="sxs-lookup"><span data-stu-id="13cdb-125">If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user's workbook because it will change the file format.</span></span>

## <a name="custom-function-behavior-for-xll-compatible-functions"></a><span data-ttu-id="13cdb-126">XLL 兼容函数的自定义函数行为</span><span class="sxs-lookup"><span data-stu-id="13cdb-126">Custom function behavior for XLL compatible functions</span></span>

<span data-ttu-id="13cdb-127">打开电子表格且存在等效的加载项时，外接程序的 XLL 函数将转换为 XLL 兼容的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="13cdb-127">An add-in's XLL functions are converted to XLL compatible custom functions when a spreadsheet is opened and there is an equivalent add-in available.</span></span> <span data-ttu-id="13cdb-128">在下一次保存时，XLL 函数将在兼容模式下写入文件中，以便它们使用 XLL 和 Excel 外接程序自定义函数（当在其他平台上）。</span><span class="sxs-lookup"><span data-stu-id="13cdb-128">On the next save, the XLL functions are written to the file in a compatible mode so that they work with both the XLL and Excel add-in custom functions (when on other platforms).</span></span>

<span data-ttu-id="13cdb-129">下表比较了 XLL 用户定义函数、XLL 兼容的自定义函数和 Excel 加载项自定义函数之间的功能。</span><span class="sxs-lookup"><span data-stu-id="13cdb-129">The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel add-in custom functions.</span></span>

|         |<span data-ttu-id="13cdb-130">XLL 用户定义的函数</span><span class="sxs-lookup"><span data-stu-id="13cdb-130">XLL user-defined function</span></span> |<span data-ttu-id="13cdb-131">XLL 兼容的自定义函数</span><span class="sxs-lookup"><span data-stu-id="13cdb-131">XLL compatible custom functions</span></span> |<span data-ttu-id="13cdb-132">Excel 加载项自定义函数</span><span class="sxs-lookup"><span data-stu-id="13cdb-132">Excel add-in custom function</span></span> |
|---------|---------|---------|---------|
| <span data-ttu-id="13cdb-133">支持的平台</span><span class="sxs-lookup"><span data-stu-id="13cdb-133">Supported platforms</span></span> | <span data-ttu-id="13cdb-134">Windows</span><span class="sxs-lookup"><span data-stu-id="13cdb-134">Windows</span></span> | <span data-ttu-id="13cdb-135">Windows、macOS、web 浏览器</span><span class="sxs-lookup"><span data-stu-id="13cdb-135">Windows, macOS, web browser</span></span> | <span data-ttu-id="13cdb-136">Windows、macOS、web 浏览器</span><span class="sxs-lookup"><span data-stu-id="13cdb-136">Windows, macOS, web browser</span></span> |
| <span data-ttu-id="13cdb-137">支持的文件格式</span><span class="sxs-lookup"><span data-stu-id="13cdb-137">Supported file formats</span></span> | <span data-ttu-id="13cdb-138">.XLSX、XLSB、XLSM、XLS</span><span class="sxs-lookup"><span data-stu-id="13cdb-138">XLSX, XLSB, XLSM, XLS</span></span> | <span data-ttu-id="13cdb-139">.XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="13cdb-139">XLSX, XLSB, XLSM</span></span> | <span data-ttu-id="13cdb-140">.XLSX、XLSB、XLSM</span><span class="sxs-lookup"><span data-stu-id="13cdb-140">XLSX, XLSB, XLSM</span></span> |
| <span data-ttu-id="13cdb-141">公式自动完成</span><span class="sxs-lookup"><span data-stu-id="13cdb-141">Formula autocomplete</span></span> | <span data-ttu-id="13cdb-142">否</span><span class="sxs-lookup"><span data-stu-id="13cdb-142">No</span></span> | <span data-ttu-id="13cdb-143">可访问</span><span class="sxs-lookup"><span data-stu-id="13cdb-143">Yes</span></span> | <span data-ttu-id="13cdb-144">是</span><span class="sxs-lookup"><span data-stu-id="13cdb-144">Yes</span></span> |
| <span data-ttu-id="13cdb-145">媒体</span><span class="sxs-lookup"><span data-stu-id="13cdb-145">Streaming</span></span> | <span data-ttu-id="13cdb-146">可通过 xlfRTD 和 XLL 回调实现。</span><span class="sxs-lookup"><span data-stu-id="13cdb-146">Possible via xlfRTD and XLL callback.</span></span> | <span data-ttu-id="13cdb-147">否</span><span class="sxs-lookup"><span data-stu-id="13cdb-147">No</span></span> | <span data-ttu-id="13cdb-148">可访问</span><span class="sxs-lookup"><span data-stu-id="13cdb-148">Yes</span></span> |
| <span data-ttu-id="13cdb-149">函数的本地化</span><span class="sxs-lookup"><span data-stu-id="13cdb-149">Localization of functions</span></span> | <span data-ttu-id="13cdb-150">否</span><span class="sxs-lookup"><span data-stu-id="13cdb-150">No</span></span> | <span data-ttu-id="13cdb-151">否。</span><span class="sxs-lookup"><span data-stu-id="13cdb-151">No.</span></span> <span data-ttu-id="13cdb-152">名称和 ID 必须与现有 XLL 的函数相匹配。</span><span class="sxs-lookup"><span data-stu-id="13cdb-152">The Name and ID must match the existing XLL's functions.</span></span> | <span data-ttu-id="13cdb-153">是</span><span class="sxs-lookup"><span data-stu-id="13cdb-153">Yes</span></span> |
| <span data-ttu-id="13cdb-154">可变函数</span><span class="sxs-lookup"><span data-stu-id="13cdb-154">Volatile functions</span></span> | <span data-ttu-id="13cdb-155">是</span><span class="sxs-lookup"><span data-stu-id="13cdb-155">Yes</span></span> | <span data-ttu-id="13cdb-156">是</span><span class="sxs-lookup"><span data-stu-id="13cdb-156">Yes</span></span> | <span data-ttu-id="13cdb-157">是</span><span class="sxs-lookup"><span data-stu-id="13cdb-157">Yes</span></span> |
| <span data-ttu-id="13cdb-158">多线程重新计算支持</span><span class="sxs-lookup"><span data-stu-id="13cdb-158">Multi-threaded recalculation support</span></span> | <span data-ttu-id="13cdb-159">是</span><span class="sxs-lookup"><span data-stu-id="13cdb-159">Yes</span></span> | <span data-ttu-id="13cdb-160">是</span><span class="sxs-lookup"><span data-stu-id="13cdb-160">Yes</span></span> | <span data-ttu-id="13cdb-161">是</span><span class="sxs-lookup"><span data-stu-id="13cdb-161">Yes</span></span> |
| <span data-ttu-id="13cdb-162">计算行为</span><span class="sxs-lookup"><span data-stu-id="13cdb-162">Calculation behavior</span></span> | <span data-ttu-id="13cdb-163">无 UI。</span><span class="sxs-lookup"><span data-stu-id="13cdb-163">No UI.</span></span> <span data-ttu-id="13cdb-164">在计算过程中，Excel 可能会无响应。</span><span class="sxs-lookup"><span data-stu-id="13cdb-164">Excel can be unresponsive during calculation.</span></span> | <span data-ttu-id="13cdb-165">用户将看到 #BUSY！</span><span class="sxs-lookup"><span data-stu-id="13cdb-165">Users will see #BUSY!</span></span> <span data-ttu-id="13cdb-166">在返回结果之前。</span><span class="sxs-lookup"><span data-stu-id="13cdb-166">until a result is returned.</span></span> | <span data-ttu-id="13cdb-167">用户将看到 #BUSY！</span><span class="sxs-lookup"><span data-stu-id="13cdb-167">Users will see #BUSY!</span></span> <span data-ttu-id="13cdb-168">在返回结果之前。</span><span class="sxs-lookup"><span data-stu-id="13cdb-168">until a result is returned.</span></span> |
| <span data-ttu-id="13cdb-169">要求集</span><span class="sxs-lookup"><span data-stu-id="13cdb-169">Requirement sets</span></span> | <span data-ttu-id="13cdb-170">不适用</span><span class="sxs-lookup"><span data-stu-id="13cdb-170">N/A</span></span> | <span data-ttu-id="13cdb-171">Customfunctions.js 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="13cdb-171">CustomFunctions 1.1 and later</span></span> | <span data-ttu-id="13cdb-172">Customfunctions.js 1.1 及更高版本</span><span class="sxs-lookup"><span data-stu-id="13cdb-172">CustomFunctions 1.1 and later</span></span> |

## <a name="see-also"></a><span data-ttu-id="13cdb-173">另请参阅</span><span class="sxs-lookup"><span data-stu-id="13cdb-173">See also</span></span>

- [<span data-ttu-id="13cdb-174">使 Excel 外接程序与现有 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="13cdb-174">Make your Excel add-in compatible with an existing COM add-in</span></span>](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [<span data-ttu-id="13cdb-175">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="13cdb-175">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
