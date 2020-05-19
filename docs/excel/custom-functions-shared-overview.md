---
ms.date: 05/17/2020
description: 了解如何在同一 JavaScript 运行时中运行自定义函数、功能区按钮和任务窗格代码，以便在加载项中协调方案。
title: 在共享 JavaScript 运行时中运行外接程序代码
localization_priority: Priority
ms.openlocfilehash: afb07c5223e26ba1e1adbf40c7a4b2e4f7c06349
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275929"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a><span data-ttu-id="e38c0-103">概述：在共享 JavaScript 运行时中运行外接程序代码</span><span class="sxs-lookup"><span data-stu-id="e38c0-103">Overview: Run your add-in code in a shared JavaScript runtimes</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="e38c0-104">运行 Windows 版 Excel 或 Mac 版 Excel 时，加载项将在单独的 JavaScript 运行时环境中运行功能区按钮、自定义函数和任务窗格的代码。</span><span class="sxs-lookup"><span data-stu-id="e38c0-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="e38c0-105">这会产生一些局限性，例如无法轻松共享全局数据，也不能通过自定义函数访问所有 CORS 功能。</span><span class="sxs-lookup"><span data-stu-id="e38c0-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="e38c0-106">但是，你可以将 Excel 加载项配置为在同一 JavaScript 运行时（也称为共享运行时）中共享代码。</span><span class="sxs-lookup"><span data-stu-id="e38c0-106">However, you can configure your Excel add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="e38c0-107">这可在加载项中实现更好的协调，并且可从加载项的所有部分访问任务窗格 DOM 和 CORS。</span><span class="sxs-lookup"><span data-stu-id="e38c0-107">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="e38c0-108">配置共享运行时可实现以下方案：</span><span class="sxs-lookup"><span data-stu-id="e38c0-108">Configuring a shared runtime enables the following scenarios:</span></span>

- <span data-ttu-id="e38c0-109">加载项将具有可供功能区、任务窗格和自定义函数访问的共享 DOM。</span><span class="sxs-lookup"><span data-stu-id="e38c0-109">Your add-in will have a shared DOM that the ribbon, task pane, and custom functions can all access.</span></span>
- <span data-ttu-id="e38c0-110">自定义函数将具有完整的 CORS 支持。</span><span class="sxs-lookup"><span data-stu-id="e38c0-110">Your custom functions will have full CORS support.</span></span>
- <span data-ttu-id="e38c0-111">自定义函数可调用 Office.js API 以读取电子表格文档数据。</span><span class="sxs-lookup"><span data-stu-id="e38c0-111">Your custom functions can call Office.js APIs to read spreadsheet document data.</span></span>
- <span data-ttu-id="e38c0-112">打开文档后，加载项即可运行代码。</span><span class="sxs-lookup"><span data-stu-id="e38c0-112">Your add-in can run code as soon as the document is opened.</span></span>
- <span data-ttu-id="e38c0-113">关闭任务窗格后，加载项可以继续运行代码。</span><span class="sxs-lookup"><span data-stu-id="e38c0-113">Your add-in can continue running code after the task pane is closed.</span></span>

<span data-ttu-id="e38c0-114">当使用任务窗格在共享运行时中运行自定义函数时，它将在不同平台上的浏览器实例中运行，如 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中所述。此外，Excel 加载项在功能区上显示的任何按钮都将在同一共享运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="e38c0-114">When you run custom functions in a shared runtime with the task pane, it will run in a browser instance on different platforms as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="e38c0-115">下图显示了自定义函数、功能区 UI 和任务窗格代码如何在同一 JavaScript 运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="e38c0-115">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![在包含 Excel 中的功能区按钮和任务窗格的共享运行时中运行的自定义函数](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a><span data-ttu-id="e38c0-117">设置共享运行时</span><span class="sxs-lookup"><span data-stu-id="e38c0-117">Set up a shared runtime</span></span>

<span data-ttu-id="e38c0-118">请参阅[配置共享运行时文章](./configure-your-add-in-to-use-a-shared-runtime.md)，了解如何将自定义函数设置为使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="e38c0-118">See the [configuring a shared runtime article](./configure-your-add-in-to-use-a-shared-runtime.md) to learn how to set up your custom functions to use a shared runtime.</span></span>

### <a name="debugging"></a><span data-ttu-id="e38c0-119">调试</span><span class="sxs-lookup"><span data-stu-id="e38c0-119">Debugging</span></span>

<span data-ttu-id="e38c0-120">使用共享运行时时，目前不能使用 Visual Studio Code 在 Windows 版 Excel 中调试自定义函数。</span><span class="sxs-lookup"><span data-stu-id="e38c0-120">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="e38c0-121">而是需要使用开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="e38c0-121">You'll need to use developer tools instead.</span></span> <span data-ttu-id="e38c0-122">有关详细信息，请参阅[使用 Windows 10 上的开发人员工具调试加载项](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)。</span><span class="sxs-lookup"><span data-stu-id="e38c0-122">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="e38c0-123">向我们提供反馈</span><span class="sxs-lookup"><span data-stu-id="e38c0-123">Give us feedback</span></span>

<span data-ttu-id="e38c0-124">我们非常乐意听取有关此功能的反馈。</span><span class="sxs-lookup"><span data-stu-id="e38c0-124">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="e38c0-125">如果你发现此功能存在任何 bug、问题或具有相关请求，请通过在 [office-js repo](https://github.com/OfficeDev/office-js) 中创建 GitHub 问题来告诉我们。</span><span class="sxs-lookup"><span data-stu-id="e38c0-125">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="e38c0-126">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e38c0-126">See also</span></span>

- [<span data-ttu-id="e38c0-127">教程：在 Excel 自定义函数和任务窗格之间共享数据和事件</span><span class="sxs-lookup"><span data-stu-id="e38c0-127">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="e38c0-128">从自定义函数调用 Excel Api</span><span class="sxs-lookup"><span data-stu-id="e38c0-128">Call Excel APIs from your custom function</span></span>](call-excel-apis-from-custom-function.md)
