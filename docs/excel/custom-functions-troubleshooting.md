---
ms.date: 04/15/2019
description: Excel 自定义函数中的常见问题疑难解答。
title: 自定义函数疑难解答（预览版）
localization_priority: Priority
ms.openlocfilehash: 6a11b733c528028a2ea9fc48b08e9308a2cf6e97
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914219"
---
# <a name="troubleshoot-custom-functions"></a><span data-ttu-id="c2247-103">自定义函数疑难解答</span><span class="sxs-lookup"><span data-stu-id="c2247-103">Troubleshoot custom functions</span></span>

<span data-ttu-id="c2247-104">开发自定义函数时，创建和测试函数可能会遇到产品错误。</span><span class="sxs-lookup"><span data-stu-id="c2247-104">When developing custom functions, you may encounter errors in the product while creating and testing your functions.</span></span>

<span data-ttu-id="c2247-105">若要解决这些问题，可以[启用运行时日志记录以捕获错误](#enable-runtime-logging)，并参考[Excel 的本机错误消息](#check-for-excel-error-messages)。</span><span class="sxs-lookup"><span data-stu-id="c2247-105">To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages).</span></span> <span data-ttu-id="c2247-106">另外，检查常见错误，例如未正确[验证 SSL 证书](#my-add-in-wont-load-verify-certificates)、[有未解析的 promise](#ensure-promises-return)，以及忘记[关联函数](#my-functions-wont-load-associate-functions)。</span><span class="sxs-lookup"><span data-stu-id="c2247-106">Also, check for common mistakes such as not [verifying ssl certificates](#my-add-in-wont-load-verify-certificates) properly, [leaving promises unresolved](#ensure-promises-return), and forgetting to [associate your functions](#my-functions-wont-load-associate-functions).</span></span>

## <a name="enable-runtime-logging"></a><span data-ttu-id="c2247-107">启用运行时日志记录</span><span class="sxs-lookup"><span data-stu-id="c2247-107">Enable runtime logging</span></span>

<span data-ttu-id="c2247-108">如果在 Windows 上的 Office 中测试加载项，应[启用运行时日志记录](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="c2247-108">If you are testing your add-in in Office on Windows, you should [enable runtime logging](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span> <span data-ttu-id="c2247-109">运行时日志记录将 `console.log` 语句传递给创建的单独日志文件，以帮助发现问题。</span><span class="sxs-lookup"><span data-stu-id="c2247-109">Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues.</span></span> <span data-ttu-id="c2247-110">这些语句涵盖了各种错误，其中包括加载项的 XML 清单文件、运行时条件或自定义函数安装的相关错误。</span><span class="sxs-lookup"><span data-stu-id="c2247-110">The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.</span></span>  <span data-ttu-id="c2247-111">有关运行时日志记录的详细信息，请参阅[使用运行时日志记录调试加载项](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="c2247-111">For more information about runtime logging, see [Use runtime logging to debug your add-in](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span>  

### <a name="check-for-excel-error-messages"></a><span data-ttu-id="c2247-112">检查 Excel 错误消息</span><span class="sxs-lookup"><span data-stu-id="c2247-112">Check for Excel error messages</span></span>

<span data-ttu-id="c2247-113">Excel 有许多内置错误消息，如果存在计算错误，系统会将向单元格返回这些错误消息。</span><span class="sxs-lookup"><span data-stu-id="c2247-113">Excel has a number of built-in error messages which are returned to a cell if there is calculation error.</span></span> <span data-ttu-id="c2247-114">自定义函数仅使用以下错误消息：`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A` 和 `#BUSY!`。</span><span class="sxs-lookup"><span data-stu-id="c2247-114">Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.</span></span>

## <a name="common-issues"></a><span data-ttu-id="c2247-115">常见问题</span><span class="sxs-lookup"><span data-stu-id="c2247-115">Common issues</span></span>

### <a name="my-add-in-wont-load-verify-certificates"></a><span data-ttu-id="c2247-116">我的加载项无法加载：验证证书</span><span class="sxs-lookup"><span data-stu-id="c2247-116">My add-in won't load: verify certifications</span></span>

<span data-ttu-id="c2247-117">如果加载项无法安装，请验证是否为托管加载项的 Web 服务器正确配置了 SSL 证书。</span><span class="sxs-lookup"><span data-stu-id="c2247-117">If your add-in fails to install, verify that the SSL certificates are configured correctly for the web server that's hosting your add-in.</span></span> <span data-ttu-id="c2247-118">通常，如果 SSL 证书存在问题，将会在 Excel 警告中看到一条错误消息，提示无法正确安装加载项。</span><span class="sxs-lookup"><span data-stu-id="c2247-118">Typically if there is a problem with SSL certificates, you will see an error message in Excel warning you that your add-in could not be installed properly.</span></span> <span data-ttu-id="c2247-119">有关详细信息，请参阅[添加自签名证书作为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="c2247-119">For more information, see [Adding self-signed certificates as trusted root certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>

### <a name="my-functions-wont-load-associate-functions"></a><span data-ttu-id="c2247-120">我的函数无法加载：关联函数</span><span class="sxs-lookup"><span data-stu-id="c2247-120">My functions won't load: associate functions</span></span>

<span data-ttu-id="c2247-121">在自定义函数的脚本文件中，需要将每个自定义函数与在 [JSON 元数据文件](custom-functions-json.md)中指定的 ID 相关联。</span><span class="sxs-lookup"><span data-stu-id="c2247-121">In your custom functions' script file, you need to associate each custom function with its ID specified in the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="c2247-122">使用 `CustomFunctions.associate()` 方法可实现此操作。</span><span class="sxs-lookup"><span data-stu-id="c2247-122">This is done by using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="c2247-123">通常，在每个函数之后或脚本文件的末尾调用此方法。</span><span class="sxs-lookup"><span data-stu-id="c2247-123">Typically this method call is made after each function or at the end of the script file.</span></span> <span data-ttu-id="c2247-124">如果没有关联自定义函数，它将不起作用。</span><span class="sxs-lookup"><span data-stu-id="c2247-124">If a custom function is not associated, it will not work.</span></span>

<span data-ttu-id="c2247-125">下面的示例显示了一个 add 函数，后跟一个与相应的 JSON ID `ADD` 相关联的函数名称 `add`。</span><span class="sxs-lookup"><span data-stu-id="c2247-125">The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="c2247-126">有关此过程的更多信息，请参阅[将函数名称与 JSON 元数据相关联](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="c2247-126">For more information on this process, see [Associating function names with json metadata](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).</span></span>

### <a name="ensure-promises-return"></a><span data-ttu-id="c2247-127">确保返回 promise</span><span class="sxs-lookup"><span data-stu-id="c2247-127">Ensure promises return</span></span>

<span data-ttu-id="c2247-128">在 Excel 等待自定义函数完成时，它会在单元格中</span><span class="sxs-lookup"><span data-stu-id="c2247-128">When Excel is waiting for a custom function to complete, it displays #BUSY!</span></span> <span data-ttu-id="c2247-129">显示 #BUSY!。</span><span class="sxs-lookup"><span data-stu-id="c2247-129">in the cell.</span></span> <span data-ttu-id="c2247-130">如果自定义函数代码返回一个 promise，但 promise 不返回结果，则 Excel 将继续显示 #BUSY!。</span><span class="sxs-lookup"><span data-stu-id="c2247-130">If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing #BUSY!.</span></span> <span data-ttu-id="c2247-131">查看函数以确保所有 promise 都正确地向单元格返回结果。</span><span class="sxs-lookup"><span data-stu-id="c2247-131">Check your functions to make sure that any promises are properly returning a result to a cell.</span></span>

## <a name="reporting-feedback"></a><span data-ttu-id="c2247-132">报告反馈</span><span class="sxs-lookup"><span data-stu-id="c2247-132">Reporting Feedback</span></span>

<span data-ttu-id="c2247-133">如果遇到本文中未记录的问题，请告诉我们。</span><span class="sxs-lookup"><span data-stu-id="c2247-133">If you are encountering issues that aren't documented here, let us know.</span></span> <span data-ttu-id="c2247-134">有两种方法可以报告问题。</span><span class="sxs-lookup"><span data-stu-id="c2247-134">There are two ways to report issues.</span></span>

### <a name="in-excel-on-windows-or-mac"></a><span data-ttu-id="c2247-135">在 Wndows 或 Mac 上的 Excel 中</span><span class="sxs-lookup"><span data-stu-id="c2247-135">In Excel on Windows or Mac</span></span>

<span data-ttu-id="c2247-136">如果使用 Excel for Windows 或 Excel for Mac，可以直接从 Excel 向 Office 扩展性团队报告反馈。</span><span class="sxs-lookup"><span data-stu-id="c2247-136">If using Excel for Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="c2247-137">为此，请选择“文件”->“反馈”->“发送哭脸”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="c2247-137">To do this, select **File -> Feedback -> Send a Frown**.</span></span> <span data-ttu-id="c2247-138">发送哭脸将提供必要的日志，以帮助我们了解你遇到的问题。</span><span class="sxs-lookup"><span data-stu-id="c2247-138">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

### <a name="in-github"></a><span data-ttu-id="c2247-139">在 Github 中</span><span class="sxs-lookup"><span data-stu-id="c2247-139">In Github</span></span>

<span data-ttu-id="c2247-140">可以随时通过任何文档页底部的“内容反馈”功能提交所遇到的问题，也可以[直接向自定义功能存储库提交新问题](https://github.com/OfficeDev/Excel-Custom-Functions/issues)。</span><span class="sxs-lookup"><span data-stu-id="c2247-140">Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="c2247-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c2247-141">See also</span></span>

* [<span data-ttu-id="c2247-142">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="c2247-142">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="c2247-143">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="c2247-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="c2247-144">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="c2247-144">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="c2247-145">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="c2247-145">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="c2247-146">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="c2247-146">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
