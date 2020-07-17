---
ms.date: 05/16/2020
description: 使用 Internet Explorer 11 测试 Office 外接程序。
title: Internet Explorer 11 测试
localization_priority: Normal
ms.openlocfilehash: 1d6852d08308088a020e86ce7f5ab9cfdb9ab978
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006435"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a><span data-ttu-id="76874-103">使用 Internet Explorer 11 测试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="76874-103">Test your Office Add-in using Internet Explorer 11</span></span>

<span data-ttu-id="76874-104">根据你的外接程序的规范，你可能会计划支持较早版本的 Windows 和 Office，这需要在 Internet Explorer 11 上进行测试。</span><span class="sxs-lookup"><span data-stu-id="76874-104">Depending on the specifications of your add-in, you may plan to support older versions of Windows and Office, which require testing on Internet Explorer 11.</span></span> <span data-ttu-id="76874-105">在将外接程序提交到 AppSource 时，通常需要执行此过程。</span><span class="sxs-lookup"><span data-stu-id="76874-105">This is often necessary as part of submitting your add-in to AppSource.</span></span> <span data-ttu-id="76874-106">您可以使用以下命令行工具从外接程序使用的更新式运行时切换到 Internet Explorer 11 运行时进行此测试。</span><span class="sxs-lookup"><span data-stu-id="76874-106">You can use the following command line tooling to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span>

## <a name="pre-requisites"></a><span data-ttu-id="76874-107">先决条件</span><span class="sxs-lookup"><span data-stu-id="76874-107">Pre-requisites</span></span>

- <span data-ttu-id="76874-108">[Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）</span><span class="sxs-lookup"><span data-stu-id="76874-108">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>
- <span data-ttu-id="76874-109">一个代码编辑器。</span><span class="sxs-lookup"><span data-stu-id="76874-109">A code editor.</span></span> <span data-ttu-id="76874-110">建议[Visual Studio Code](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="76874-110">We recommend [Visual Studio Code](https://code.visualstudio.com/)</span></span>
- [<span data-ttu-id="76874-111">是 Office 预览体验计划的一部分</span><span class="sxs-lookup"><span data-stu-id="76874-111">Be part of the Office Insider program</span></span>](https://insider.office.com)

<span data-ttu-id="76874-112">这些说明假定您先设置了 "Yo Office 生成器" 项目。</span><span class="sxs-lookup"><span data-stu-id="76874-112">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="76874-113">如果你之前未执行此操作，请考虑阅读快速启动，例如， [Excel 外接程序](../quickstarts/excel-quickstart-jquery.md)。</span><span class="sxs-lookup"><span data-stu-id="76874-113">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="using-ie11-tooling"></a><span data-ttu-id="76874-114">使用 IE11 工具</span><span class="sxs-lookup"><span data-stu-id="76874-114">Using IE11 tooling</span></span>

1. <span data-ttu-id="76874-115">创建 "Yo Office 生成器" 项目。</span><span class="sxs-lookup"><span data-stu-id="76874-115">Create a Yo Office generator project.</span></span> <span data-ttu-id="76874-116">无论选择哪种类型的项目，此工具都将适用于所有项目类型。</span><span class="sxs-lookup"><span data-stu-id="76874-116">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

> <span data-ttu-id="76874-117">!便笺如果您有一个现有项目，并且想要添加此工具而不创建新项目，请跳过此步骤并移动到下一步。</span><span class="sxs-lookup"><span data-stu-id="76874-117">![NOTE] If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

2. <span data-ttu-id="76874-118">在新项目的根文件夹中，在命令行中运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="76874-118">In the root folder of your new project, run the following in the command line:</span></span>

```command&nbsp;line
npx office-addin-dev-settings webview manifest.xml ie
```
<span data-ttu-id="76874-119">您应该会在命令行中看到一条注释，web 视图类型现在设置为 IE。</span><span class="sxs-lookup"><span data-stu-id="76874-119">You should see a note in the command line that the web view type is now set to IE.</span></span>

> <span data-ttu-id="76874-120">!尖不必使用此工具，但应帮助调试与 Internet Explorer 11 运行时相关的大多数问题。</span><span class="sxs-lookup"><span data-stu-id="76874-120">![TIP] It isn't necessary to use this tooling, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="76874-121">为实现全面的可靠性，应使用安装了 Windows 7 和 Office 2013 副本的计算机进行测试。</span><span class="sxs-lookup"><span data-stu-id="76874-121">For complete robustness, you should test using a computer with a copy of Windows 7 and Office 2013 installed.</span></span>

## <a name="command-settings"></a><span data-ttu-id="76874-122">命令设置</span><span class="sxs-lookup"><span data-stu-id="76874-122">Command settings</span></span>

<span data-ttu-id="76874-123">如果您有一个不同的清单路径，请在命令中指定此路径，如下所示：</span><span class="sxs-lookup"><span data-stu-id="76874-123">Should you have a different manifest path, specify this in the command, as shown in the following:</span></span>

`npx office-addin-dev-settings webview [path to your manifest] ie`

<span data-ttu-id="76874-124">该 `office-addin-dev-settings webview` 命令还可以采用若干个运行时作为参数：</span><span class="sxs-lookup"><span data-stu-id="76874-124">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="76874-125">限于</span><span class="sxs-lookup"><span data-stu-id="76874-125">ie</span></span>
- <span data-ttu-id="76874-126">距</span><span class="sxs-lookup"><span data-stu-id="76874-126">edge</span></span>
- <span data-ttu-id="76874-127"> 默认值</span><span class="sxs-lookup"><span data-stu-id="76874-127">default</span></span>

## <a name="see-also"></a><span data-ttu-id="76874-128">另请参阅</span><span class="sxs-lookup"><span data-stu-id="76874-128">See also</span></span>
* [<span data-ttu-id="76874-129">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="76874-129">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="76874-130">旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="76874-130">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="76874-131">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="76874-131">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="76874-132">从任务窗格附加调试器</span><span class="sxs-lookup"><span data-stu-id="76874-132">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
