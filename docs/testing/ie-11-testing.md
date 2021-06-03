---
title: Internet Explorer 11 测试
description: 在 Office 11 上测试Internet Explorer加载项。
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: de256ee8b0633f18d3188c5bbfae52cb24ff2c35
ms.sourcegitcommit: 0d3bf72f8ddd1b287bf95f832b7ecb9d9fa62a24
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/02/2021
ms.locfileid: "52727932"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a><span data-ttu-id="097e3-103">在 Office 11 上测试Internet Explorer加载项</span><span class="sxs-lookup"><span data-stu-id="097e3-103">Test your Office Add-in on Internet Explorer 11</span></span>

<span data-ttu-id="097e3-104">如果计划通过 AppSource 销售加载项或计划支持较旧版本的 Windows 和 Office，加载项必须在基于 Internet Explorer 11 (IE11) 的可嵌入浏览器控件中运行。</span><span class="sxs-lookup"><span data-stu-id="097e3-104">If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11).</span></span> <span data-ttu-id="097e3-105">可以使用命令行从外接程序使用的更现代运行时切换到 Internet Explorer 11 运行时进行此测试。</span><span class="sxs-lookup"><span data-stu-id="097e3-105">You can use a command line to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span> <span data-ttu-id="097e3-106">有关哪些版本的 Windows 和 Office使用 Internet Explorer 11 Web 视图控件的信息，请参阅 Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器。</span><span class="sxs-lookup"><span data-stu-id="097e3-106">For information about which versions of Windows and Office use the Internet Explorer 11 web view control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="097e3-107">Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。</span><span class="sxs-lookup"><span data-stu-id="097e3-107">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="097e3-108">如果要使用 ECMAScript 2015 或更高版本的语法和功能，有两个选项：</span><span class="sxs-lookup"><span data-stu-id="097e3-108">If you want to use the syntax and features of ECMAScript 2015 or later, you have two options:</span></span>
>
> - <span data-ttu-id="097e3-109">在 ECMAScript 2015 (（也称为 ES6) 或更高版本 JavaScript）中编写代码，或在 TypeScript 中编写代码，然后使用编译器（如 [#A0](https://babeljs.io/) 或 [tsc）](https://www.typescriptlang.org/index.html)将代码编译为 ES5 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="097e3-109">Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).</span></span>
> - <span data-ttu-id="097e3-110">在 ECMAScript 2015 或更高版本的 JavaScript[](https://en.wikipedia.org/wiki/Polyfill_(programming))中编写，但也加载填充库（如[core-js，](https://github.com/zloirock/core-js)它使 IE 能够运行代码）。</span><span class="sxs-lookup"><span data-stu-id="097e3-110">Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.</span></span>
>
> <span data-ttu-id="097e3-111">有关这些选项的详细信息，请参阅 Support [Internet Explorer 11](../develop/support-ie-11.md)。</span><span class="sxs-lookup"><span data-stu-id="097e3-111">For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).</span></span>
>
> <span data-ttu-id="097e3-112">此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。</span><span class="sxs-lookup"><span data-stu-id="097e3-112">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="097e3-113">若要在 Internet Explorer 11 浏览器上测试外接程序，Office web 版中Internet Explorer并[旁加载外接程序](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="097e3-113">To test your add-in on the Internet Explorer 11 browser, open Office on the web in Internet Explorer and [sideload the add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="097e3-114">先决条件</span><span class="sxs-lookup"><span data-stu-id="097e3-114">Prerequisites</span></span>

- <span data-ttu-id="097e3-115">[Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）</span><span class="sxs-lookup"><span data-stu-id="097e3-115">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

<span data-ttu-id="097e3-116">这些说明假定你之前已经设置了 Yo Office生成器项目。</span><span class="sxs-lookup"><span data-stu-id="097e3-116">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="097e3-117">如果之前尚未这样做，请考虑阅读快速入门，例如适用于Excel[入门](../quickstarts/excel-quickstart-jquery.md)。</span><span class="sxs-lookup"><span data-stu-id="097e3-117">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="switching-to-the-internet-explorer-11-webview"></a><span data-ttu-id="097e3-118">切换到 Internet Explorer 11 Webview</span><span class="sxs-lookup"><span data-stu-id="097e3-118">Switching to the Internet Explorer 11 webview</span></span>

1. <span data-ttu-id="097e3-119">创建 Yo Office生成器项目。</span><span class="sxs-lookup"><span data-stu-id="097e3-119">Create a Yo Office generator project.</span></span> <span data-ttu-id="097e3-120">选择哪种项目并不重要，此工具将用于所有项目类型。</span><span class="sxs-lookup"><span data-stu-id="097e3-120">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

    > [!NOTE]
    > <span data-ttu-id="097e3-121">如果您有一个现有项目，并且想要在不创建新项目的情况下添加此工具，请跳过此步骤并移至下一步。</span><span class="sxs-lookup"><span data-stu-id="097e3-121">If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

1. <span data-ttu-id="097e3-122">在项目的根文件夹中，在命令行中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="097e3-122">In the root folder of your project, run the following in the command line.</span></span> <span data-ttu-id="097e3-123">此示例假定项目的清单文件位于根中。</span><span class="sxs-lookup"><span data-stu-id="097e3-123">This example assumes that your project's manifest file is in the root.</span></span> <span data-ttu-id="097e3-124">如果不是，请指定清单文件的相对路径。</span><span class="sxs-lookup"><span data-stu-id="097e3-124">If it isn't, specify the relative path to the manifest file.</span></span> <span data-ttu-id="097e3-125">您应该在命令行中看到一条消息，指出 Web 视图类型现在设置为 IE。</span><span class="sxs-lookup"><span data-stu-id="097e3-125">You should see a message in the command line that the web view type is now set to IE.</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> <span data-ttu-id="097e3-126">虽然不需要使用此命令，但它应有助于调试与 11 运行时Internet Explorer大多数问题。</span><span class="sxs-lookup"><span data-stu-id="097e3-126">It isn't necessary to use this command, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="097e3-127">为提供完整的稳定性，应测试使用具有 Windows 7、8.1 和 10 的各种版本以及不同版本的 Office 的计算机。</span><span class="sxs-lookup"><span data-stu-id="097e3-127">For complete robustness, you should test using computers with various combinations of Windows 7, 8.1, and 10 and various versions of Office.</span></span> <span data-ttu-id="097e3-128">有关详细信息，请参阅Office[外接程序](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器和如何还原到早期版本[Office。](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841)</span><span class="sxs-lookup"><span data-stu-id="097e3-128">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span></span>

### <a name="command-options"></a><span data-ttu-id="097e3-129">命令选项</span><span class="sxs-lookup"><span data-stu-id="097e3-129">Command options</span></span>

<span data-ttu-id="097e3-130">该命令 `office-addin-dev-settings webview` 还可以将多个运行时用作参数：</span><span class="sxs-lookup"><span data-stu-id="097e3-130">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="097e3-131">ie</span><span class="sxs-lookup"><span data-stu-id="097e3-131">ie</span></span>
- <span data-ttu-id="097e3-132">edge</span><span class="sxs-lookup"><span data-stu-id="097e3-132">edge</span></span>
- <span data-ttu-id="097e3-133">default</span><span class="sxs-lookup"><span data-stu-id="097e3-133">default</span></span>

## <a name="see-also"></a><span data-ttu-id="097e3-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="097e3-134">See also</span></span>

* [<span data-ttu-id="097e3-135">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="097e3-135">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="097e3-136">旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="097e3-136">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="097e3-137">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="097e3-137">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="097e3-138">从任务窗格附加调试器</span><span class="sxs-lookup"><span data-stu-id="097e3-138">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
