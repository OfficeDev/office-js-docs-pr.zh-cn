---
title: 从这里开始！ 面向初学者的 Office 加载项构建指南
description: 通过 Office 加载项的学习资源为初学者提供指导的推荐路径。
ms.date: 04/16/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 026f90ea62960cbbf5ab4420d40a4a9165139cae
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547617"
---
# <a name="start-here-a-guide-for-beginners-making-office-add-ins"></a><span data-ttu-id="3a537-104">从这里开始！</span><span class="sxs-lookup"><span data-stu-id="3a537-104">Start Here!</span></span> <span data-ttu-id="3a537-105">面向初学者的 Office 加载项构建指南</span><span class="sxs-lookup"><span data-stu-id="3a537-105">A guide for beginners making Office Add-ins</span></span>

<span data-ttu-id="3a537-106">想要开始构建自己的跨平台 Office 扩展？</span><span class="sxs-lookup"><span data-stu-id="3a537-106">Want to get started building your own cross-platform Office extensions?</span></span> <span data-ttu-id="3a537-107">以下步骤显示了需要先阅读的内容、要安装的工具以及要完成的推荐教程。</span><span class="sxs-lookup"><span data-stu-id="3a537-107">The following steps show you what to read first, what tools to install, and recommended tutorials to complete.</span></span>

## <a name="step-0-prerequisites"></a><span data-ttu-id="3a537-108">步骤 0：先决条件</span><span class="sxs-lookup"><span data-stu-id="3a537-108">Step 0: Prerequisites</span></span>

- <span data-ttu-id="3a537-109">Office 加载项本质上是嵌入在 Office 中的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="3a537-109">Office Add-ins are essentially web applications embedded in Office.</span></span> <span data-ttu-id="3a537-110">因此，你首先应该对 Web 应用程序以及如何在 Web 上托管它们有基本的了解。</span><span class="sxs-lookup"><span data-stu-id="3a537-110">So, you should first have a basic understanding of web applications and how they are hosted on the web.</span></span> <span data-ttu-id="3a537-111">Internet、书籍和在线课程提供了有关它的大量信息。</span><span class="sxs-lookup"><span data-stu-id="3a537-111">There is an enormous amount of information about this on the Internet, in books, and in online courses.</span></span> <span data-ttu-id="3a537-112">如果你根本不了解 Web 应用程序，那么一个很好的开始方法是在</span><span class="sxs-lookup"><span data-stu-id="3a537-112">A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?"</span></span> <span data-ttu-id="3a537-113">必应上搜索“什么是 Web 应用程序？”。</span><span class="sxs-lookup"><span data-stu-id="3a537-113">on Bing.</span></span>
- <span data-ttu-id="3a537-114">创建 Office 加载项时将使用的主要编程语言是 JavaScript 或 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="3a537-114">The primary programming language you will use in creating Office Add-ins is JavaScript or TypeScript.</span></span> <span data-ttu-id="3a537-115">可将 TypeScript 视为 JavaScript 的强类型版本。</span><span class="sxs-lookup"><span data-stu-id="3a537-115">You can think of TypeScript as a strongly-typed version of JavaScript.</span></span> <span data-ttu-id="3a537-116">如果你不熟悉这两种语言，但是你有使用 VBA、VB.Net、C# 的经验，则你可能会发现 TypeScript 更容易学习。</span><span class="sxs-lookup"><span data-stu-id="3a537-116">If you are not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you will probably find TypeScript easier to learn.</span></span> <span data-ttu-id="3a537-117">此外，Internet、书籍和在线课程提供了有关这些语言的大量信息。</span><span class="sxs-lookup"><span data-stu-id="3a537-117">Again, there is a wealth of information about these languages on the Internet, in books, and in online courses.</span></span>

## <a name="step-1-begin-with-fundamentals"></a><span data-ttu-id="3a537-118">步骤 1：从基础知识开始</span><span class="sxs-lookup"><span data-stu-id="3a537-118">Step 1: Begin with fundamentals</span></span>

<span data-ttu-id="3a537-119">我们知道你渴望开始编码，但是在打开 IDE 或代码编辑器之前，你应该先阅读一些有关 Office 加载项的信息。</span><span class="sxs-lookup"><span data-stu-id="3a537-119">We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.</span></span>

- <span data-ttu-id="3a537-120">[Office 加载项平台概述](office-add-ins.md)：了解什么是 Office Web 加载项以及它们与扩展 Office（如 VSTO 加载项）的旧方法有何区别。</span><span class="sxs-lookup"><span data-stu-id="3a537-120">[Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.</span></span>
- <span data-ttu-id="3a537-121">[构建 Office 加载项](office-add-ins-fundamentals.md)：概述 Office 加载项的开发和生命周期，包括工具、创建加载项 UI 以及使用 JavaScript API 与 Office 文档进行交互。</span><span class="sxs-lookup"><span data-stu-id="3a537-121">[Building Office Add-ins](office-add-ins-fundamentals.md): Get an overview of Office add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.</span></span>

<span data-ttu-id="3a537-122">这些文章中有许多链接，但是如果你是 Office 加载项的初学者，我们建议你在阅读完后返回此处并继续下一部分。</span><span class="sxs-lookup"><span data-stu-id="3a537-122">There are a lot of links in those articles, but if you're a beginner with Office Add-ins, we recommend that you come back here when you've read them and continue with the next section.</span></span>

## <a name="step-2-install-tools-and-create-your-first-add-in"></a><span data-ttu-id="3a537-123">步骤 2：安装工具并创建首个加载项</span><span class="sxs-lookup"><span data-stu-id="3a537-123">Step 2: Install tools and create your first add-in</span></span>

<span data-ttu-id="3a537-124">现在，你已有了大致的了解，下面需要深入了解其中一个快速入门。</span><span class="sxs-lookup"><span data-stu-id="3a537-124">You've got the big picture now, so dive in with one of our quick starts.</span></span> <span data-ttu-id="3a537-125">出于学习平台的目的，我们推荐使用 Excel 快速入门。</span><span class="sxs-lookup"><span data-stu-id="3a537-125">For purposes of learning the platform, we recommend the Excel quick start.</span></span> <span data-ttu-id="3a537-126">我们提供基于 Visual Studio 的版本以及基于 Node.js 和 Visual Studio Code 的版本。</span><span class="sxs-lookup"><span data-stu-id="3a537-126">There is a version that is based on Visual Studio and a version that is based in Node.js and Visual Studio Code.</span></span>

- [<span data-ttu-id="3a537-127">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="3a537-127">Visual Studio</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="3a537-128">Node.js 和 Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="3a537-128">Node.js and Visual Studio Code</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a><span data-ttu-id="3a537-129">步骤 3：代码</span><span class="sxs-lookup"><span data-stu-id="3a537-129">Step 3: Code</span></span>

<span data-ttu-id="3a537-130">你无法通过阅读车主手册学会开车，因此请从此 [Excel 教程](../tutorials/excel-tutorial.md)开始编码吧。</span><span class="sxs-lookup"><span data-stu-id="3a537-130">You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md).</span></span> <span data-ttu-id="3a537-131">你将使用 Office JavaScript 库和加载项清单中的一些 XML。</span><span class="sxs-lookup"><span data-stu-id="3a537-131">You'll be using the Office JavaScript library and some XML in the add-in's manifest.</span></span> <span data-ttu-id="3a537-132">无需记住任何内容，因为在后面的步骤中，你将获得关于这两者的更多背景知识。</span><span class="sxs-lookup"><span data-stu-id="3a537-132">There's no need to memorize anything, because you'll be getting more background about both in a later steps.</span></span>

## <a name="step-4-understand-the-javascript-library"></a><span data-ttu-id="3a537-133">步骤 4：了解 JavaScript 库</span><span class="sxs-lookup"><span data-stu-id="3a537-133">Step 4: Understand the JavaScript library</span></span>

<span data-ttu-id="3a537-134">首先，通过来自 Microsoft Learn 的本教程大致了解 Office JavaScript 库：[了解 Office JavaScript API](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index)。</span><span class="sxs-lookup"><span data-stu-id="3a537-134">First, get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).</span></span>

<span data-ttu-id="3a537-135">然后，使用我们的 [Script Lab 工具](explore-with-script-lab.md)（一种用于运行和探索 API 的沙箱）来探索 Office JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="3a537-135">Then explore the Office JavaScript APIs with our [the Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.</span></span>

## <a name="step-5-understand-the-manifest"></a><span data-ttu-id="3a537-136">步骤 5：了解清单</span><span class="sxs-lookup"><span data-stu-id="3a537-136">Step 5: Understand the manifest</span></span>

<span data-ttu-id="3a537-137">在 [Office 加载项 XML 清单](../develop/add-in-manifests.md)中了解加载项清单的用途以及有关其 XML 标记的简介。</span><span class="sxs-lookup"><span data-stu-id="3a537-137">Get an understanding of the purposes of the add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="3a537-138">后续步骤</span><span class="sxs-lookup"><span data-stu-id="3a537-138">Next Steps</span></span>

<span data-ttu-id="3a537-139">恭喜你完成了初学者的 Office 加载项学习之路！</span><span class="sxs-lookup"><span data-stu-id="3a537-139">Congratulations on finishing the beginner's learning path for Office Add-ins!</span></span> <span data-ttu-id="3a537-140">以下是进一步探索我们的文档的一些建议：</span><span class="sxs-lookup"><span data-stu-id="3a537-140">Here are some suggestions for further exploration of our documentation:</span></span>

- <span data-ttu-id="3a537-141">其他 Office 应用程序的教程或快速入门：</span><span class="sxs-lookup"><span data-stu-id="3a537-141">Tutorials or quick starts for other Office applications:</span></span>

  - [<span data-ttu-id="3a537-142">OneNote 快速入门</span><span class="sxs-lookup"><span data-stu-id="3a537-142">OneNote quick start</span></span>](../quickstarts/onenote-quickstart.md)
  - [<span data-ttu-id="3a537-143">Outlook 教程</span><span class="sxs-lookup"><span data-stu-id="3a537-143">Outlook tutorial</span></span>](/outlook/add-ins/addin-tutorial)
  - [<span data-ttu-id="3a537-144">PowerPoint 教程</span><span class="sxs-lookup"><span data-stu-id="3a537-144">PowerPoint tutorial</span></span>](../tutorials/powerpoint-tutorial.md)
  - [<span data-ttu-id="3a537-145">Project 快速入门</span><span class="sxs-lookup"><span data-stu-id="3a537-145">Project quick start</span></span>](../quickstarts/project-quickstart.md)
  - [<span data-ttu-id="3a537-146">Word 教程</span><span class="sxs-lookup"><span data-stu-id="3a537-146">Word tutorial</span></span>](../tutorials/word-tutorial.md)

- <span data-ttu-id="3a537-147">其他重要主题：</span><span class="sxs-lookup"><span data-stu-id="3a537-147">Other important subjects:</span></span>

  - [<span data-ttu-id="3a537-148">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="3a537-148">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
  - [<span data-ttu-id="3a537-149">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="3a537-149">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
  - [<span data-ttu-id="3a537-150">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="3a537-150">Design Office Add-ins</span></span>](../design/add-in-design.md)
  - [<span data-ttu-id="3a537-151">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="3a537-151">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
  - [<span data-ttu-id="3a537-152">部署和发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="3a537-152">Deploy and publish Office Add-ins</span></span>](../publish/publish.md)
  - [<span data-ttu-id="3a537-153">资源</span><span class="sxs-lookup"><span data-stu-id="3a537-153">Resources</span></span>](../resources/resources-links-help.md)
