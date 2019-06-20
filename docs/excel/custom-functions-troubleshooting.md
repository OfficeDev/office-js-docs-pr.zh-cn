---
ms.date: 06/17/2019
description: Excel 自定义函数中的常见问题疑难解答。
title: 自定义函数疑难解答
localization_priority: Priority
ms.openlocfilehash: f407e103d8f628710c5f58a9787b3a802dcd39c8
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059900"
---
# <a name="troubleshoot-custom-functions"></a><span data-ttu-id="ef960-103">自定义函数疑难解答</span><span class="sxs-lookup"><span data-stu-id="ef960-103">Troubleshoot custom functions</span></span>

<span data-ttu-id="ef960-104">开发自定义函数时，创建和测试函数可能会遇到产品错误。</span><span class="sxs-lookup"><span data-stu-id="ef960-104">When developing custom functions, you may encounter errors in the product while creating and testing your functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="ef960-105">若要解决这些问题，可以[启用运行时日志记录以捕获错误](#enable-runtime-logging)，并参考[Excel 的本机错误消息](#check-for-excel-error-messages)。</span><span class="sxs-lookup"><span data-stu-id="ef960-105">To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages).</span></span> <span data-ttu-id="ef960-106">另外，检查常见错误，例如未正确[有未解析的 promise](#ensure-promises-return) 以及忘记[关联函数](#my-functions-wont-load-associate-functions)。</span><span class="sxs-lookup"><span data-stu-id="ef960-106">Also, check for common mistakes such as [leaving promises unresolved](#ensure-promises-return) and forgetting to [associate your functions](#my-functions-wont-load-associate-functions).</span></span>

## <a name="enable-runtime-logging"></a><span data-ttu-id="ef960-107">启用运行时日志记录</span><span class="sxs-lookup"><span data-stu-id="ef960-107">Enable runtime logging</span></span>

<span data-ttu-id="ef960-108">如果在 Windows 上的 Office 中测试加载项，应[启用运行时日志记录](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="ef960-108">If you are testing your add-in in Office on Windows, you should [enable runtime logging](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span> <span data-ttu-id="ef960-109">运行时日志记录将 `console.log` 语句传递给创建的单独日志文件，以帮助发现问题。</span><span class="sxs-lookup"><span data-stu-id="ef960-109">Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues.</span></span> <span data-ttu-id="ef960-110">这些语句涵盖了各种错误，其中包括加载项的 XML 清单文件、运行时条件或自定义函数安装的相关错误。</span><span class="sxs-lookup"><span data-stu-id="ef960-110">The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.</span></span>  <span data-ttu-id="ef960-111">有关运行时日志记录的详细信息，请参阅[使用运行时日志记录调试加载项](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="ef960-111">For more information about runtime logging, see [Use runtime logging to debug your add-in](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span>  

### <a name="check-for-excel-error-messages"></a><span data-ttu-id="ef960-112">检查 Excel 错误消息</span><span class="sxs-lookup"><span data-stu-id="ef960-112">Check for Excel error messages</span></span>

<span data-ttu-id="ef960-113">Excel 有许多内置错误消息，如果存在计算错误，系统会将向单元格返回这些错误消息。</span><span class="sxs-lookup"><span data-stu-id="ef960-113">Excel has a number of built-in error messages which are returned to a cell if there is calculation error.</span></span> <span data-ttu-id="ef960-114">自定义函数仅使用以下错误消息：`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A` 和 `#BUSY!`。</span><span class="sxs-lookup"><span data-stu-id="ef960-114">Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.</span></span>

<span data-ttu-id="ef960-115">通常情况下，这些错误可能对应于你在 Excel 中熟悉的错误。</span><span class="sxs-lookup"><span data-stu-id="ef960-115">Generally, these errors correspond to the errors you might already be familiar with in Excel.</span></span> <span data-ttu-id="ef960-116">有一些特定于自定义函数的异常，如下所示：</span><span class="sxs-lookup"><span data-stu-id="ef960-116">The are only a few exceptions specific to custom functions, listed here:</span></span>

- <span data-ttu-id="ef960-117">`#NAME` 错误通常意味着注册函数时出错。</span><span class="sxs-lookup"><span data-stu-id="ef960-117">A `#NAME` error generally means there has been an issue registering your functions.</span></span>
- <span data-ttu-id="ef960-118">`#VALUE` 错误通常是指函数的脚本文件中出现了错误。</span><span class="sxs-lookup"><span data-stu-id="ef960-118">A `#VALUE` error typically indicates an error in the functions' script file.</span></span>
- <span data-ttu-id="ef960-119">`#N/A` 错误也可能是注册的函数无法运行的迹象。</span><span class="sxs-lookup"><span data-stu-id="ef960-119">A `#N/A` error is also maybe a sign that that function while registered could not be run.</span></span> <span data-ttu-id="ef960-120">这通常是因为缺少 `CustomFunctions.associate` 命令。</span><span class="sxs-lookup"><span data-stu-id="ef960-120">This is typically due to a missing `CustomFunctions.associate` command.</span></span>
- <span data-ttu-id="ef960-121">`#REF!` 错误可能指示函数名称与已存在的加载项中的函数名称相同。</span><span class="sxs-lookup"><span data-stu-id="ef960-121">A `#REF!` error may indicate that your function name is the same as a function name in an add-in that already exists.</span></span>

## <a name="clear-the-office-cache"></a><span data-ttu-id="ef960-122">清除 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="ef960-122">Clear the Office cache</span></span>

<span data-ttu-id="ef960-123">与自定义函数相关的信息由 Office 缓存。</span><span class="sxs-lookup"><span data-stu-id="ef960-123">Information about custom functions is cached by Office.</span></span> <span data-ttu-id="ef960-124">有时候，开发和反复重新加载带有自定义函数的加载项时，变更可能不会显示。</span><span class="sxs-lookup"><span data-stu-id="ef960-124">Sometimes while developing and repeatedly reloading an add-in with custom functions your changes may not appear.</span></span> <span data-ttu-id="ef960-125">可以通过清除 Office 缓存修复此问题。</span><span class="sxs-lookup"><span data-stu-id="ef960-125">You can fix this by clearing the Office cache.</span></span> <span data-ttu-id="ef960-126">有关详细信息，请参阅[使用清单验证和解决问题](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)一文中的“清除 Office 缓存”部分。</span><span class="sxs-lookup"><span data-stu-id="ef960-126">For more information, see the "Clear the Office cache" section in the article [Validate and troubleshoot issues with your manifest](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)</span></span>

## <a name="common-issues"></a><span data-ttu-id="ef960-127">常见问题</span><span class="sxs-lookup"><span data-stu-id="ef960-127">Common issues</span></span>

### <a name="my-functions-wont-load-associate-functions"></a><span data-ttu-id="ef960-128">我的函数无法加载：关联函数</span><span class="sxs-lookup"><span data-stu-id="ef960-128">My functions won't load: associate functions</span></span>

<span data-ttu-id="ef960-129">在自定义函数的脚本文件中，需要将每个自定义函数与在 [JSON 元数据文件](custom-functions-json.md)中指定的 ID 相关联。</span><span class="sxs-lookup"><span data-stu-id="ef960-129">In your custom functions' script file, you need to associate each custom function with its ID specified in the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="ef960-130">使用 `CustomFunctions.associate()` 方法可实现此操作。</span><span class="sxs-lookup"><span data-stu-id="ef960-130">This is done by using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="ef960-131">通常，在每个函数之后或脚本文件的末尾调用此方法。</span><span class="sxs-lookup"><span data-stu-id="ef960-131">Typically this method call is made after each function or at the end of the script file.</span></span> <span data-ttu-id="ef960-132">如果没有关联自定义函数，它将不起作用。</span><span class="sxs-lookup"><span data-stu-id="ef960-132">If a custom function is not associated, it will not work.</span></span>

<span data-ttu-id="ef960-133">下面的示例显示了一个 add 函数，后跟一个与相应的 JSON ID `ADD` 相关联的函数名称 `add`。</span><span class="sxs-lookup"><span data-stu-id="ef960-133">The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="ef960-134">有关此过程的更多信息，请参阅[将函数名称与 JSON 元数据相关联](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="ef960-134">For more information on this process, see [Associating function names with json metadata](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).</span></span>

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a><span data-ttu-id="ef960-135">无法从 localhost 打开加载项：使用本地环回异常</span><span class="sxs-lookup"><span data-stu-id="ef960-135">Can't open add-in from localhost: use a local loopback exception</span></span>

<span data-ttu-id="ef960-136">如果看到错误“我们无法从 localhost 打开此加载项”，则需要启用本地环回异常。</span><span class="sxs-lookup"><span data-stu-id="ef960-136">If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exception.</span></span> <span data-ttu-id="ef960-137">有关如何执行此操作的详细信息，请参阅[此 Microsoft 支持文章](https://support.microsoft.com/zh-CN/help/4490419/local-loopback-exemption-does-not-work)。</span><span class="sxs-lookup"><span data-stu-id="ef960-137">For details on how to do this, see [this Microsoft support article](https://support.microsoft.com/en-us/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="runtime-logging-reports-typeerror-network-request-failed-on-excel-for-windows"></a><span data-ttu-id="ef960-138">Excel for Windows 上的运行时日志记录报告“TypeError：网络请求失败”</span><span class="sxs-lookup"><span data-stu-id="ef960-138">Runtime logging reports "TypeError: Network request failed" on Excel for Windows</span></span>

<span data-ttu-id="ef960-139">如果在调用 localhost 服务器时在[运行时日志](custom-functions-troubleshooting.md#enable-runtime-logging)中看到错误“TypeError：网络请求失败”，则需要启用本地环回异常。</span><span class="sxs-lookup"><span data-stu-id="ef960-139">If you see the error "TypeError: Network request failed" in your [runtime log](custom-functions-troubleshooting.md#enable-runtime-logging) while making calls to your localhost server, you'll need to enable a local loopback exception.</span></span> <span data-ttu-id="ef960-140">有关如何执行此操作的详细信息，请参阅[此 Microsoft 支持文章](https://support.microsoft.com/zh-CN/help/4490419/local-loopback-exemption-does-not-work)。</span><span class="sxs-lookup"><span data-stu-id="ef960-140">For details on how to do this, see [this Microsoft support article](https://support.microsoft.com/en-us/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="ensure-promises-return"></a><span data-ttu-id="ef960-141">确保返回 promise</span><span class="sxs-lookup"><span data-stu-id="ef960-141">Ensure promises return</span></span>

<span data-ttu-id="ef960-142">在 Excel 等待自定义函数完成时，它会在单元格中</span><span class="sxs-lookup"><span data-stu-id="ef960-142">When Excel is waiting for a custom function to complete, it displays #BUSY!</span></span> <span data-ttu-id="ef960-143">显示 #BUSY!。</span><span class="sxs-lookup"><span data-stu-id="ef960-143">in the cell.</span></span> <span data-ttu-id="ef960-144">如果自定义函数代码返回一个 promise，但 promise 不返回结果，则 Excel 将继续显示 #BUSY!。</span><span class="sxs-lookup"><span data-stu-id="ef960-144">If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing #BUSY!.</span></span> <span data-ttu-id="ef960-145">查看函数以确保所有 promise 都正确地向单元格返回结果。</span><span class="sxs-lookup"><span data-stu-id="ef960-145">Check your functions to make sure that any promises are properly returning a result to a cell.</span></span>

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a><span data-ttu-id="ef960-146">错误：开发服务器已在端口 3000 上运行</span><span class="sxs-lookup"><span data-stu-id="ef960-146">Error: The dev server is already running on port 3000</span></span>

<span data-ttu-id="ef960-147">有时候，运行 `npm start` 时，你可能会看到开发服务器已在端口 3000（或加载项使用的任何端口）上运行的错误。</span><span class="sxs-lookup"><span data-stu-id="ef960-147">Sometimes when running `npm start` you may see an error that the dev server is already running on port 3000 (or whichever port your add-in uses).</span></span> <span data-ttu-id="ef960-148">可以通过运行 `npm stop` 或关闭 Node.js 窗口停止开发服务器运行。</span><span class="sxs-lookup"><span data-stu-id="ef960-148">You can stop the dev server by running `npm stop` or by closing the Node.js window.</span></span> <span data-ttu-id="ef960-149">但在某些情况下，开发服务器可能需要几分钟才能实际停止运行。</span><span class="sxs-lookup"><span data-stu-id="ef960-149">But in some cases in can take a few minutes for the dev server to actually stop running.</span></span>

## <a name="reporting-feedback"></a><span data-ttu-id="ef960-150">报告反馈</span><span class="sxs-lookup"><span data-stu-id="ef960-150">Reporting feedback</span></span>

<span data-ttu-id="ef960-151">如果遇到本文中未记录的问题，请告诉我们。</span><span class="sxs-lookup"><span data-stu-id="ef960-151">If you are encountering issues that aren't documented here, let us know.</span></span> <span data-ttu-id="ef960-152">有两种方法可以报告问题。</span><span class="sxs-lookup"><span data-stu-id="ef960-152">There are two ways to report issues.</span></span>

### <a name="in-excel-on-windows-or-mac"></a><span data-ttu-id="ef960-153">在 Wndows 或 Mac 上的 Excel 中</span><span class="sxs-lookup"><span data-stu-id="ef960-153">In Excel on Windows or Mac</span></span>

<span data-ttu-id="ef960-154">如果使用 Windows 版 Excel 或 Mac 版 Excel，可以直接从 Excel 向 Office 扩展性团队报告反馈。</span><span class="sxs-lookup"><span data-stu-id="ef960-154">If using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="ef960-155">为此，请选择“文件”->“反馈”->“发送哭脸”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="ef960-155">To do this, select **File -> Feedback -> Send a Frown**.</span></span> <span data-ttu-id="ef960-156">发送哭脸将提供必要的日志，以帮助我们了解你遇到的问题。</span><span class="sxs-lookup"><span data-stu-id="ef960-156">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

### <a name="in-github"></a><span data-ttu-id="ef960-157">在 Github 中</span><span class="sxs-lookup"><span data-stu-id="ef960-157">In Github</span></span>

<span data-ttu-id="ef960-158">可以随时通过任何文档页底部的“内容反馈”功能提交所遇到的问题，也可以[直接向自定义功能存储库提交新问题](https://github.com/OfficeDev/Excel-Custom-Functions/issues)。</span><span class="sxs-lookup"><span data-stu-id="ef960-158">Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="ef960-159">后续步骤</span><span class="sxs-lookup"><span data-stu-id="ef960-159">Next steps</span></span>
<span data-ttu-id="ef960-160">了解如何[调试自定义函数](custom-functions-debugging.md)。</span><span class="sxs-lookup"><span data-stu-id="ef960-160">Learn how to [debug your custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ef960-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ef960-161">See also</span></span>

* [<span data-ttu-id="ef960-162">自定义函数元数据自动生成</span><span class="sxs-lookup"><span data-stu-id="ef960-162">Custom functions metadata autogeneration</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="ef960-163">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="ef960-163">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="ef960-164">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="ef960-164">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="ef960-165">让自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="ef960-165">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="ef960-166">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="ef960-166">Create custom functions in Excel</span></span>](custom-functions-overview.md)
