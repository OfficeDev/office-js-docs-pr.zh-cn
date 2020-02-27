---
ms.date: 02/13/2020
description: 了解如何在同一 JavaScript 运行时中运行自定义函数、功能区按钮和任务窗格代码，以便在加载项中协调方案。
title: 在共享 JavaScript 运行时中运行加载项代码（预览版）
localization_priority: Priority
ms.openlocfilehash: d9d73a5ae2ff1da09d1a5fd7d02514cb28be0e2d
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284113"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtime-preview"></a><span data-ttu-id="52199-103">概述：在共享 JavaScript 运行时中运行加载项代码（预览版）</span><span class="sxs-lookup"><span data-stu-id="52199-103">Overview: Run your add-in code in a shared JavaScript runtime (preview)</span></span>

[!include[Running custom functions in shared JavaScript runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="52199-104">运行 Windows 版 Excel 或 Mac 版 Excel 时，加载项将在单独的 JavaScript 运行时环境中运行功能区按钮、自定义函数和任务窗格的代码。</span><span class="sxs-lookup"><span data-stu-id="52199-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="52199-105">这会产生一些局限性，例如无法轻松共享全局数据，也不能通过自定义函数访问所有 CORS 功能。</span><span class="sxs-lookup"><span data-stu-id="52199-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="52199-106">但是，你可以将 Excel 加载项配置为在同一 JavaScript 运行时（也称为共享运行时）中共享代码。</span><span class="sxs-lookup"><span data-stu-id="52199-106">However, you can configure your Excel add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="52199-107">这可在加载项中实现更好的协调，并且可从加载项的所有部分访问任务窗格 DOM 和 CORS。</span><span class="sxs-lookup"><span data-stu-id="52199-107">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="52199-108">配置共享运行时可实现以下方案：</span><span class="sxs-lookup"><span data-stu-id="52199-108">Configuring a shared runtime enables the following scenarios:</span></span>

- <span data-ttu-id="52199-109">加载项将具有可供功能区、任务窗格和自定义函数访问的共享 DOM。</span><span class="sxs-lookup"><span data-stu-id="52199-109">Your add-in will have a shared DOM that the ribbon, task pane, and custom functions can all access.</span></span>
- <span data-ttu-id="52199-110">自定义函数将具有完整的 CORS 支持。</span><span class="sxs-lookup"><span data-stu-id="52199-110">Your custom functions will have full CORS support.</span></span>
- <span data-ttu-id="52199-111">自定义函数可调用 Office.js API 以读取电子表格文档数据。</span><span class="sxs-lookup"><span data-stu-id="52199-111">Your custom functions can call Office.js APIs to read spreadsheet document data.</span></span>
- <span data-ttu-id="52199-112">打开文档后，加载项即可运行代码。</span><span class="sxs-lookup"><span data-stu-id="52199-112">Your add-in can run code as soon as the document is opened.</span></span>
- <span data-ttu-id="52199-113">关闭任务窗格后，加载项可以继续运行代码。</span><span class="sxs-lookup"><span data-stu-id="52199-113">Your add-in can continue running code after the task pane is closed.</span></span>

<span data-ttu-id="52199-114">当使用任务窗格在共享运行时中运行自定义函数时，它将在不同平台上的浏览器实例中运行，如 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中所述。此外，Excel 加载项在功能区上显示的任何按钮都将在同一共享运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="52199-114">When you run custom functions in a shared runtime with the task pane, it will run in a browser instance on different platforms as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="52199-115">下图显示了自定义函数、功能区 UI 和任务窗格代码如何在同一 JavaScript 运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="52199-115">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![使用 Excel 中的功能区按钮和任务窗格在共享运行时中运行的自定义函数](../images/custom-functions-in-browser-runtime.png)

## <a name="differences-when-running-custom-functions-in-a-shared-runtime"></a><span data-ttu-id="52199-117">在共享运行时中运行自定义函数时的差异</span><span class="sxs-lookup"><span data-stu-id="52199-117">Differences when running custom functions in a shared runtime</span></span>

<span data-ttu-id="52199-118">将 Excel 加载项项目配置为在共享运行时中运行自定义函数时，与使用自定义函数运行时有一些不同。</span><span class="sxs-lookup"><span data-stu-id="52199-118">When you configure your Excel add-in project to run custom functions in a shared runtime, there are a few differences from using the custom function runtime.</span></span>

### <a name="storage"></a><span data-ttu-id="52199-119">存储</span><span class="sxs-lookup"><span data-stu-id="52199-119">Storage</span></span>

<span data-ttu-id="52199-120">无需再在任务窗格、自定义函数或功能区 UI 之间使用**存储** API 来共享数据。</span><span class="sxs-lookup"><span data-stu-id="52199-120">You no longer need to use the **Storage** API to share data between the task pane, custom functions or ribbon UI.</span></span> <span data-ttu-id="52199-121">可将全局变量置于 **window** 对象中，或使用自己的首选状态管理方法。</span><span class="sxs-lookup"><span data-stu-id="52199-121">You can put global variables in the **window** object, or use your own preferred state management approach.</span></span>

### <a name="authentication"></a><span data-ttu-id="52199-122">身份验证</span><span class="sxs-lookup"><span data-stu-id="52199-122">Authentication</span></span>

<span data-ttu-id="52199-123">如果在身份验证过程中收到令牌，无需使用**存储** API 在任务窗格、自定义函数和功能区 UI 之间共享它们。</span><span class="sxs-lookup"><span data-stu-id="52199-123">When you receive tokens as part of authentication, you don't need to use the **Storage** API to share them between the task pane, custom functions and ribbon UI.</span></span> <span data-ttu-id="52199-124">你可以使用自己的首选存储技术和存储位置来共享它们，例如 `localStorage`。</span><span class="sxs-lookup"><span data-stu-id="52199-124">You can use your own preferred storage technique and storage location to share them, such as `localStorage`.</span></span>

### <a name="dialog-api"></a><span data-ttu-id="52199-125">对话框 API</span><span class="sxs-lookup"><span data-stu-id="52199-125">Dialog API</span></span>

<span data-ttu-id="52199-126">无需再使用 **OfficeRuntime.Dialog** API 来显示来自自定义函数的对话框。</span><span class="sxs-lookup"><span data-stu-id="52199-126">You no longer need to use the **OfficeRuntime.Dialog** API to display a dialog from a custom function.</span></span> <span data-ttu-id="52199-127">可以将同一[对话框 API](../develop/dialog-api-in-office-add-ins.md) 用于自定义函数、功能区按钮和任务窗格。</span><span class="sxs-lookup"><span data-stu-id="52199-127">You can use the same [Dialog API](../develop/dialog-api-in-office-add-ins.md) for custom functions, ribbon buttons, and the task pane.</span></span>

### <a name="debugging"></a><span data-ttu-id="52199-128">调试</span><span class="sxs-lookup"><span data-stu-id="52199-128">Debugging</span></span>

<span data-ttu-id="52199-129">使用共享运行时时，目前不能使用 Visual Studio Code 在 Windows 版 Excel 中调试自定义函数。</span><span class="sxs-lookup"><span data-stu-id="52199-129">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="52199-130">你需要使用开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="52199-130">You'll need to use developer tools.</span></span> <span data-ttu-id="52199-131">有关详细信息，请参阅[使用 Windows 10 上的开发人员工具调试加载项](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)。</span><span class="sxs-lookup"><span data-stu-id="52199-131">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

## <a name="get-started"></a><span data-ttu-id="52199-132">开始使用</span><span class="sxs-lookup"><span data-stu-id="52199-132">Get Started</span></span>

<span data-ttu-id="52199-133">若要将 Excel 加载项项目配置为在共享运行时中运行自定义函数，请参阅[将 Excel 加载项配置为使用共享 JavaScript 运行时（预览版）](configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="52199-133">To configure your Excel add-in project to run custom functions in a shared runtime, see [Configure your Excel add-in to use a shared JavaScript runtime (preview)](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="52199-134">向我们提供反馈</span><span class="sxs-lookup"><span data-stu-id="52199-134">Give us feedback</span></span>

<span data-ttu-id="52199-135">我们非常乐意听取有关此功能的反馈。</span><span class="sxs-lookup"><span data-stu-id="52199-135">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="52199-136">如果你发现此功能存在任何 bug、问题或具有相关请求，请通过在 [office-js repo](https://github.com/OfficeDev/office-js) 中创建 GitHub 问题来告诉我们。</span><span class="sxs-lookup"><span data-stu-id="52199-136">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="52199-137">另请参阅</span><span class="sxs-lookup"><span data-stu-id="52199-137">See also</span></span>

<span data-ttu-id="52199-138">共享运行时的相关文章列表</span><span class="sxs-lookup"><span data-stu-id="52199-138">List of related articles for shared runtime</span></span>
- [<span data-ttu-id="52199-139">教程：在 Excel 自定义函数和任务窗格之间共享数据和事件（预览）</span><span class="sxs-lookup"><span data-stu-id="52199-139">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="52199-140">从自定义函数中调用 Excel API（预览版）</span><span class="sxs-lookup"><span data-stu-id="52199-140">Call Excel APIs from your custom function (preview)</span></span>](call-excel-apis-from-custom-function.md)