---
title: Internet Explorer 11 测试
description: 在 Office 11 上测试Internet Explorer加载项。
ms.date: 06/18/2021
localization_priority: Normal
ms.openlocfilehash: fa9550884a24feffdd750171f3a7e08648f9432f
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076404"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a><span data-ttu-id="09683-103">在 Office 11 上测试Internet Explorer加载项</span><span class="sxs-lookup"><span data-stu-id="09683-103">Test your Office Add-in on Internet Explorer 11</span></span>

> [!IMPORTANT]
> <span data-ttu-id="09683-104">**Internet Explorer外接程序Office中使用的内容**</span><span class="sxs-lookup"><span data-stu-id="09683-104">**Internet Explorer still used in Office Add-ins**</span></span>
>
> <span data-ttu-id="09683-105">Microsoft 将终止对Internet Explorer的支持，但这不会显著Office外接程序。平台和 Office 版本（包括 Office 2019 的所有一次购买版本）的一些组合将继续使用 Internet Explorer 11 随附的 Webview 控件来托管外接程序，如[Office](../concepts/browsers-used-by-office-web-add-ins.md)外接程序使用的浏览器所说明。此外，提交到[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)的加载项Internet Explorer支持这些组合，因此也支持这些组合。</span><span class="sxs-lookup"><span data-stu-id="09683-105">Microsoft is ending support for Internet Explorer, but this doesn't significantly affect Office Add-ins. Some combinations of platforms and Office versions, including all one-time-purchase versions through Office 2019, will continue to use the webview control that comes with Internet Explorer 11 to host add-ins, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Moreover, support for these combinations, and hence for Internet Explorer, is still required for add-ins submitted to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span></span> <span data-ttu-id="09683-106">有两 *个变化* ：</span><span class="sxs-lookup"><span data-stu-id="09683-106">Two things *are* changing:</span></span>
>
> - <span data-ttu-id="09683-107">AppSource 不再使用作为浏览器Office web 版Internet Explorer加载项。</span><span class="sxs-lookup"><span data-stu-id="09683-107">AppSource no longer tests add-ins in Office on the web using Internet Explorer as the browser.</span></span> <span data-ttu-id="09683-108">但 AppSource 仍测试使用 Office *版本的平台* 和桌面Internet Explorer。</span><span class="sxs-lookup"><span data-stu-id="09683-108">But AppSource still tests for combinations of platform and Office *desktop* versions that use Internet Explorer.</span></span>
> - <span data-ttu-id="09683-109">2021 Script Lab，Internet Explorer工具将停止工作。 [](../overview/explore-with-script-lab.md)</span><span class="sxs-lookup"><span data-stu-id="09683-109">The [Script Lab tool](../overview/explore-with-script-lab.md) will stop working in Internet Explorer sometime in 2021.</span></span>

<span data-ttu-id="09683-110">如果计划通过 AppSource 销售加载项或计划支持较旧版本的 Windows 和 Office，加载项必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中运行。</span><span class="sxs-lookup"><span data-stu-id="09683-110">If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11).</span></span> <span data-ttu-id="09683-111">可以使用命令行从外接程序使用的更现代运行时切换到 Internet Explorer 11 运行时进行此测试。</span><span class="sxs-lookup"><span data-stu-id="09683-111">You can use a command line to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span> <span data-ttu-id="09683-112">有关哪些版本的 Windows 和 Office使用 Internet Explorer 11 Web 视图控件的信息，请参阅 Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器。</span><span class="sxs-lookup"><span data-stu-id="09683-112">For information about which versions of Windows and Office use the Internet Explorer 11 web view control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="09683-113">Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。</span><span class="sxs-lookup"><span data-stu-id="09683-113">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="09683-114">如果要使用 ECMAScript 2015 或更高版本的语法和功能，有两个选项：</span><span class="sxs-lookup"><span data-stu-id="09683-114">If you want to use the syntax and features of ECMAScript 2015 or later, you have two options:</span></span>
>
> - <span data-ttu-id="09683-115">在 ECMAScript 2015 (（也称为 ES6) 或更高版本 JavaScript）中编写代码，或在 TypeScript 中编写代码，然后使用编译器（如 [#A0](https://babeljs.io/) 或 [tsc）](https://www.typescriptlang.org/index.html)将代码编译为 ES5 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="09683-115">Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).</span></span>
> - <span data-ttu-id="09683-116">在 ECMAScript 2015 或更高版本的 JavaScript[](https://en.wikipedia.org/wiki/Polyfill_(programming))中编写，但也加载填充库（如[core-js，](https://github.com/zloirock/core-js)它使 IE 能够运行代码）。</span><span class="sxs-lookup"><span data-stu-id="09683-116">Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.</span></span>
>
> <span data-ttu-id="09683-117">有关这些选项的详细信息，请参阅 Support [Internet Explorer 11](../develop/support-ie-11.md)。</span><span class="sxs-lookup"><span data-stu-id="09683-117">For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).</span></span>
>
> <span data-ttu-id="09683-118">此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。</span><span class="sxs-lookup"><span data-stu-id="09683-118">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="09683-119">若要在 Internet Explorer 11 浏览器上测试外接程序，Office web 版中Internet Explorer并[旁加载外接程序](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="09683-119">To test your add-in on the Internet Explorer 11 browser, open Office on the web in Internet Explorer and [sideload the add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="09683-120">先决条件</span><span class="sxs-lookup"><span data-stu-id="09683-120">Prerequisites</span></span>

- <span data-ttu-id="09683-121">[Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）</span><span class="sxs-lookup"><span data-stu-id="09683-121">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

<span data-ttu-id="09683-122">这些说明假定你之前已经设置了 Yo Office生成器项目。</span><span class="sxs-lookup"><span data-stu-id="09683-122">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="09683-123">如果之前尚未这样做，请考虑阅读快速入门，例如适用于Excel[入门](../quickstarts/excel-quickstart-jquery.md)。</span><span class="sxs-lookup"><span data-stu-id="09683-123">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="switching-to-the-internet-explorer-11-webview"></a><span data-ttu-id="09683-124">切换到 Internet Explorer 11 Webview</span><span class="sxs-lookup"><span data-stu-id="09683-124">Switching to the Internet Explorer 11 webview</span></span>

1. <span data-ttu-id="09683-125">创建 Yo Office生成器项目。</span><span class="sxs-lookup"><span data-stu-id="09683-125">Create a Yo Office generator project.</span></span> <span data-ttu-id="09683-126">选择哪种项目并不重要，此工具将用于所有项目类型。</span><span class="sxs-lookup"><span data-stu-id="09683-126">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

    > [!NOTE]
    > <span data-ttu-id="09683-127">如果您有一个现有项目，并且想要在不创建新项目的情况下添加此工具，请跳过此步骤并移至下一步。</span><span class="sxs-lookup"><span data-stu-id="09683-127">If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

1. <span data-ttu-id="09683-128">在项目的根文件夹中，在命令行中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="09683-128">In the root folder of your project, run the following in the command line.</span></span> <span data-ttu-id="09683-129">此示例假定项目的清单文件位于根中。</span><span class="sxs-lookup"><span data-stu-id="09683-129">This example assumes that your project's manifest file is in the root.</span></span> <span data-ttu-id="09683-130">如果不是，请指定清单文件的相对路径。</span><span class="sxs-lookup"><span data-stu-id="09683-130">If it isn't, specify the relative path to the manifest file.</span></span> <span data-ttu-id="09683-131">您应该在命令行中看到一条消息，指出 Web 视图类型现在设置为 IE。</span><span class="sxs-lookup"><span data-stu-id="09683-131">You should see a message in the command line that the web view type is now set to IE.</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> <span data-ttu-id="09683-132">虽然不需要使用此命令，但它应有助于调试与 11 运行时Internet Explorer大多数问题。</span><span class="sxs-lookup"><span data-stu-id="09683-132">It isn't necessary to use this command, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="09683-133">为提供完整的稳定性，应测试使用具有 Windows 7、8.1 和 10 的各种版本以及不同版本的 Office 的计算机。</span><span class="sxs-lookup"><span data-stu-id="09683-133">For complete robustness, you should test using computers with various combinations of Windows 7, 8.1, and 10 and various versions of Office.</span></span> <span data-ttu-id="09683-134">有关详细信息，请参阅Office[外接程序](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器和如何还原到早期版本[Office。](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841)</span><span class="sxs-lookup"><span data-stu-id="09683-134">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span></span>

### <a name="command-options"></a><span data-ttu-id="09683-135">命令选项</span><span class="sxs-lookup"><span data-stu-id="09683-135">Command options</span></span>

<span data-ttu-id="09683-136">该命令 `office-addin-dev-settings webview` 还可以将多个运行时用作参数：</span><span class="sxs-lookup"><span data-stu-id="09683-136">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="09683-137">ie</span><span class="sxs-lookup"><span data-stu-id="09683-137">ie</span></span>
- <span data-ttu-id="09683-138">edge</span><span class="sxs-lookup"><span data-stu-id="09683-138">edge</span></span>
- <span data-ttu-id="09683-139">default</span><span class="sxs-lookup"><span data-stu-id="09683-139">default</span></span>

## <a name="see-also"></a><span data-ttu-id="09683-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="09683-140">See also</span></span>

* [<span data-ttu-id="09683-141">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="09683-141">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="09683-142">旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="09683-142">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="09683-143">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="09683-143">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="09683-144">从任务窗格附加调试器</span><span class="sxs-lookup"><span data-stu-id="09683-144">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
