---
title: VSTO 加载项开发人员指南
description: 资深 VSTO 加载项开发人员了解 Office Web 加载项资源的建议路径。
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6da72dbdc5dc25d222cc7c2a269d905d9271ce15
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238013"
---
# <a name="vsto-add-in-developers-guide"></a><span data-ttu-id="cd415-103">VSTO 加载项开发人员指南</span><span class="sxs-lookup"><span data-stu-id="cd415-103">VSTO add-in developer's guide</span></span>

<span data-ttu-id="cd415-104">因此，你为在 Windows 上运行的 Office 应用创建了一些 VSTO 加载项，现在正在探索扩展将在 Windows、Mac 上所运行 Office 和 Office 套件联机版的新方式：Office Web 加载项。</span><span class="sxs-lookup"><span data-stu-id="cd415-104">So, you've made some VSTO add-ins for Office applications that run on Windows and now you're exploring the new way of extending Office that will run on Windows, Mac, and the online version of the Office suite: Office Web Add-ins.</span></span>

<span data-ttu-id="cd415-105">对 Excel、Word 和其他 Office 应用程序的对象模型的理解将非常有用，因为 Office Web 加载项中的对象模型遵循类似的模式。</span><span class="sxs-lookup"><span data-stu-id="cd415-105">Your understanding of the object models for the Excel, Word, and the other Office applications will be a huge help because the object models in Office Web Add-ins follow similar patterns.</span></span> <span data-ttu-id="cd415-106">但是将面临一些挑战：</span><span class="sxs-lookup"><span data-stu-id="cd415-106">But there are going to be some challenges:</span></span>

- <span data-ttu-id="cd415-107">你将使用其他语言（JavaScript 或 TypeScript）而不是 C＃或 Visual Basic .NET。</span><span class="sxs-lookup"><span data-stu-id="cd415-107">You will be working with a different language (either JavaScript or TypeScript) instead of C# or Visual Basic .NET.</span></span> <span data-ttu-id="cd415-108">（还有一种方法，如下所述，可以重复使用 Web 加载项中存在的代码。）</span><span class="sxs-lookup"><span data-stu-id="cd415-108">(There is also a way, described below, to reuse some of your existing code in a web add-in.)</span></span>
- <span data-ttu-id="cd415-109">Office Web 加载项的部署方式不同于 VSTO 加载项。</span><span class="sxs-lookup"><span data-stu-id="cd415-109">Office Web Add-ins are deployed differently from VSTO add-ins.</span></span>
- <span data-ttu-id="cd415-110">Office Web 加载项是在 Office 应用程序中嵌入的简化浏览器窗口中运行的 Web 应用程序，因此需要对 Web 应用程序以及如何在Web服务器或云帐户上托管有基本的了解。</span><span class="sxs-lookup"><span data-stu-id="cd415-110">Office Web Add-ins are web applications that run in a simplified browser window that is embedded in the Office application, so you need to gain a basic understanding of web applications and how they are hosted on web servers or cloud accounts.</span></span> 

<span data-ttu-id="cd415-111">出于以上原因，本文的大部分内容都向完整的 Office 扩展初学者介绍了我们的学习路径：[入门指南](learning-path-beginner.md)。</span><span class="sxs-lookup"><span data-stu-id="cd415-111">For these reasons, much of this article duplicates our learning path for complete beginners to Office extensions: [Beginner's guide](learning-path-beginner.md).</span></span> <span data-ttu-id="cd415-112">我们添加了一些其他学习资源，以帮助 VSTO 加载项开发人员利用他们的经验，并帮助他们重用现有代码。</span><span class="sxs-lookup"><span data-stu-id="cd415-112">What we have added are some additional learning resources to help VSTO add-in developers leverage their experience, and also help them reuse their existing code.</span></span>

## <a name="step-0-prerequisites"></a><span data-ttu-id="cd415-113">步骤 0：先决条件</span><span class="sxs-lookup"><span data-stu-id="cd415-113">Step 0: Prerequisites</span></span>

- <span data-ttu-id="cd415-114">Office Web 加载项（也称为 Office 加载项）本质上是嵌入在 Office 中的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="cd415-114">Office Web Add-ins (also referred to as Office Add-ins) are essentially web applications embedded in Office.</span></span> <span data-ttu-id="cd415-115">因此，你首先应该对 Web 应用程序以及如何在 Web 上托管它们有基本的了解。</span><span class="sxs-lookup"><span data-stu-id="cd415-115">So, you should first have a basic understanding of web applications and how they are hosted on the web.</span></span> <span data-ttu-id="cd415-116">Internet、书籍和在线课程提供了有关它的大量信息。</span><span class="sxs-lookup"><span data-stu-id="cd415-116">There's an enormous amount of information about this on the Internet, in books, and in online courses.</span></span> <span data-ttu-id="cd415-117">如果你根本不了解 Web 应用程序，那么一个很好的开始方法是在</span><span class="sxs-lookup"><span data-stu-id="cd415-117">A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?"</span></span> <span data-ttu-id="cd415-118">必应上搜索“什么是 Web 应用程序？”。</span><span class="sxs-lookup"><span data-stu-id="cd415-118">on Bing.</span></span>
- <span data-ttu-id="cd415-119">创建 Office 加载项将使用的主要编程语言是 JavaScript 或 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="cd415-119">The primary programming language you'll use to create Office Add-ins is JavaScript or TypeScript.</span></span> <span data-ttu-id="cd415-120">可将 TypeScript 视为 JavaScript 的强类型版本。</span><span class="sxs-lookup"><span data-stu-id="cd415-120">You can think of TypeScript as a strongly-typed version of JavaScript.</span></span> <span data-ttu-id="cd415-121">如果你不熟悉这两种语言，但是你有使用 VBA、VB.Net、C# 的经验，则你可能会发现 TypeScript 更容易学习。</span><span class="sxs-lookup"><span data-stu-id="cd415-121">If you're not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you'll probably find TypeScript easier to learn.</span></span> <span data-ttu-id="cd415-122">此外，Internet、书籍和在线课程提供了有关这些语言的大量信息。</span><span class="sxs-lookup"><span data-stu-id="cd415-122">Again, there's a wealth of information about these languages on the Internet, in books, and in online courses.</span></span>

## <a name="step-1-begin-with-fundamentals"></a><span data-ttu-id="cd415-123">步骤 1：从基础知识开始</span><span class="sxs-lookup"><span data-stu-id="cd415-123">Step 1: Begin with fundamentals</span></span>

<span data-ttu-id="cd415-124">我们知道你渴望开始编码，但是在打开 IDE 或代码编辑器之前，你应该先阅读一些有关 Office 加载项的信息。</span><span class="sxs-lookup"><span data-stu-id="cd415-124">We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.</span></span>

- <span data-ttu-id="cd415-125">[Office 加载项平台概述](office-add-ins.md)：了解什么是 Office Web 加载项以及它们与扩展 Office（如 VSTO 加载项）的旧方法有何区别。</span><span class="sxs-lookup"><span data-stu-id="cd415-125">[Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.</span></span>
- <span data-ttu-id="cd415-126">[开发 Office 加载项](../develop/develop-overview.md)：获取 Office 加载项的开发和生命周期概述，包括工具、创建加载项 UI 以及使用 JavaScript API 与 Office 文档交互。</span><span class="sxs-lookup"><span data-stu-id="cd415-126">[Develop Office Add-ins](../develop/develop-overview.md): Get an overview of Office Add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.</span></span>

<span data-ttu-id="cd415-127">这些文章中有许多链接，但是如果你正在过渡至 Office Web 加载项的初学者，我们建议你在阅读完后返回此处并继续下一部分。</span><span class="sxs-lookup"><span data-stu-id="cd415-127">There are a lot of links in those articles, but if you're transitioning to Office Web Add-ins, we recommend that you come back here when you've read them and continue with the next section.</span></span>

## <a name="step-2-install-tools-and-create-your-first-add-in"></a><span data-ttu-id="cd415-128">步骤 2：安装工具并创建首个加载项</span><span class="sxs-lookup"><span data-stu-id="cd415-128">Step 2: Install tools and create your first add-in</span></span>

<span data-ttu-id="cd415-129">现在，你已有了大致的了解，下面需要深入了解其中一个快速入门。</span><span class="sxs-lookup"><span data-stu-id="cd415-129">You've got the big picture now, so dive in with one of our quick starts.</span></span> <span data-ttu-id="cd415-130">出于学习平台的目的，我们推荐使用 Excel 快速入门。</span><span class="sxs-lookup"><span data-stu-id="cd415-130">For purposes of learning the platform, we recommend the Excel quick start.</span></span> <span data-ttu-id="cd415-131">一个版本基于 Visual Studio，另一个版本基于 Node.js 和 Visual Studio Code。</span><span class="sxs-lookup"><span data-stu-id="cd415-131">There's a version based on Visual Studio and another based on Node.js and Visual Studio Code.</span></span> <span data-ttu-id="cd415-132">如果正在从 VSTO 加载项转换，可能会发现 Visual Studio 版本更易于使用。</span><span class="sxs-lookup"><span data-stu-id="cd415-132">If you're transitioning from VSTO add-ins, you'll probably find the Visual Studio version easier to work with.</span></span>

- [<span data-ttu-id="cd415-133">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="cd415-133">Visual Studio</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="cd415-134">Node.js 和 Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="cd415-134">Node.js and Visual Studio Code</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a><span data-ttu-id="cd415-135">步骤 3：代码</span><span class="sxs-lookup"><span data-stu-id="cd415-135">Step 3: Code</span></span>

<span data-ttu-id="cd415-136">你无法通过阅读车主手册学会开车，因此请从此 [Excel 教程](../tutorials/excel-tutorial.md)开始编码吧。</span><span class="sxs-lookup"><span data-stu-id="cd415-136">You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md).</span></span> <span data-ttu-id="cd415-137">你将使用 Office JavaScript 库和加载项清单中的一些 XML。</span><span class="sxs-lookup"><span data-stu-id="cd415-137">You'll be using the Office JavaScript library and some XML in the add-in's manifest.</span></span> <span data-ttu-id="cd415-138">无需记住任何内容，因为在后面的步骤中，你将获得关于这两者的更多背景知识。</span><span class="sxs-lookup"><span data-stu-id="cd415-138">There's no need to memorize anything, because you'll be getting more background about both in a later step.</span></span>

## <a name="step-4-understand-the-javascript-library"></a><span data-ttu-id="cd415-139">步骤 4：了解 JavaScript 库</span><span class="sxs-lookup"><span data-stu-id="cd415-139">Step 4: Understand the JavaScript library</span></span>

<span data-ttu-id="cd415-140">通过来自 Microsoft Learn 的本教程大致了解 Office JavaScript 库：[了解 Office JavaScript API](/learn/modules/intro-office-add-ins/3-apis)。</span><span class="sxs-lookup"><span data-stu-id="cd415-140">Get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](/learn/modules/intro-office-add-ins/3-apis).</span></span>

<span data-ttu-id="cd415-141">然后，使用 [Script Lab 工具](explore-with-script-lab.md)（一种用于运行和探索 API 的沙箱）来探索 Office JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="cd415-141">Then explore the Office JavaScript APIs with the [Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.</span></span>

### <a name="special-resource-for-vsto-add-in-developers"></a><span data-ttu-id="cd415-142">适用于 VSTO 加载项开发人员的特殊支援</span><span class="sxs-lookup"><span data-stu-id="cd415-142">Special resource for VSTO add-in developers</span></span>

<span data-ttu-id="cd415-143">这里将介绍如何查看示例加载项、[Excel 加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)。</span><span class="sxs-lookup"><span data-stu-id="cd415-143">This would be a good place to take a look at the sample add-in, [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).</span></span> <span data-ttu-id="cd415-144">创建的目的是为了突出显示 VSTO 加载项和 Office Web 加载项之间的异同，并且示例的自述文件指出了比较的重点。</span><span class="sxs-lookup"><span data-stu-id="cd415-144">It was created to highlight the similarities and differences between VSTO add-ins and Office Web Add-ins, and the readme of the sample calls out the important points of comparison.</span></span>

## <a name="step-5-understand-the-manifest"></a><span data-ttu-id="cd415-145">步骤 5：了解清单</span><span class="sxs-lookup"><span data-stu-id="cd415-145">Step 5: Understand the manifest</span></span>

<span data-ttu-id="cd415-146">在 [Office 加载项 XML 清单](../develop/add-in-manifests.md)中了解 web 加载项清单的用途以及有关其 XML 标记的简介。</span><span class="sxs-lookup"><span data-stu-id="cd415-146">Get an understanding of the purposes of the web add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="step-6-for-vsto-developers-only-reuse-your-vsto-code"></a><span data-ttu-id="cd415-147">步骤 6（仅适用于 VSTO 开发人员）：重复使用 VSTO 代码</span><span class="sxs-lookup"><span data-stu-id="cd415-147">Step 6 (for VSTO developers only): Reuse your VSTO code</span></span>

<span data-ttu-id="cd415-148">可以在 Office Web 加载项中重复使用某些 VSTO 加载项代码，方法是将其移到服务器上 Web 应用程序的后端，然后将其作为 Web API 供 JavaScript 或 TypeScript 使用。</span><span class="sxs-lookup"><span data-stu-id="cd415-148">You can reuse some of your VSTO add-in code in an Office web add-in by moving it to your web application's back end on the server and making it available to your JavaScript or TypeScript as a web API.</span></span> <span data-ttu-id="cd415-149">有关指南，参见 [教程：使用共享代码库在 VSTO 加载项与 Office 加载项之间共享代码](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="cd415-149">For guidance, see [Tutorial: Share code between both a VSTO Add-in and an Office Add-in by using a shared code library](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="cd415-150">后续步骤</span><span class="sxs-lookup"><span data-stu-id="cd415-150">Next Steps</span></span>

<span data-ttu-id="cd415-151">恭喜你完成了 VSTO 加载项的 Office Web 加载项学习之路！</span><span class="sxs-lookup"><span data-stu-id="cd415-151">Congratulations on finishing the VSTO add-in developer's learning path for Office Web Add-ins!</span></span> <span data-ttu-id="cd415-152">以下是进一步探索我们的文档的一些建议：</span><span class="sxs-lookup"><span data-stu-id="cd415-152">Here are some suggestions for further exploration of our documentation:</span></span>

- <span data-ttu-id="cd415-153">其他 Office 应用程序的教程或快速入门：</span><span class="sxs-lookup"><span data-stu-id="cd415-153">Tutorials or quick starts for other Office applications:</span></span>

  - [<span data-ttu-id="cd415-154">OneNote 快速入门</span><span class="sxs-lookup"><span data-stu-id="cd415-154">OneNote quick start</span></span>](../quickstarts/onenote-quickstart.md)
  - [<span data-ttu-id="cd415-155">Outlook 教程</span><span class="sxs-lookup"><span data-stu-id="cd415-155">Outlook tutorial</span></span>](/outlook/add-ins/addin-tutorial)
  - [<span data-ttu-id="cd415-156">PowerPoint 教程</span><span class="sxs-lookup"><span data-stu-id="cd415-156">PowerPoint tutorial</span></span>](../tutorials/powerpoint-tutorial.md)
  - [<span data-ttu-id="cd415-157">Project 快速入门</span><span class="sxs-lookup"><span data-stu-id="cd415-157">Project quick start</span></span>](../quickstarts/project-quickstart.md)
  - [<span data-ttu-id="cd415-158">Word 教程</span><span class="sxs-lookup"><span data-stu-id="cd415-158">Word tutorial</span></span>](../tutorials/word-tutorial.md)

- <span data-ttu-id="cd415-159">其他重要主题：</span><span class="sxs-lookup"><span data-stu-id="cd415-159">Other important subjects:</span></span>

  - [<span data-ttu-id="cd415-160">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="cd415-160">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
  - [<span data-ttu-id="cd415-161">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="cd415-161">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
  - [<span data-ttu-id="cd415-162">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="cd415-162">Design Office Add-ins</span></span>](../design/add-in-design.md)
  - [<span data-ttu-id="cd415-163">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="cd415-163">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
  - [<span data-ttu-id="cd415-164">部署和发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="cd415-164">Deploy and publish Office Add-ins</span></span>](../publish/publish.md)
  - [<span data-ttu-id="cd415-165">资源</span><span class="sxs-lookup"><span data-stu-id="cd415-165">Resources</span></span>](../resources/resources-links-help.md)
  - [<span data-ttu-id="cd415-166">了解 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="cd415-166">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
