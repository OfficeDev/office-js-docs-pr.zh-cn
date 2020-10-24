---
title: 初学者指南
description: 通过 Office 加载项的学习资源为初学者提供指导的推荐路径。
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: a51ffc437c9d1946b886d1e665836dd6d76f52d2
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741069"
---
# <a name="beginners-guide"></a><span data-ttu-id="d5024-103">初学者指南</span><span class="sxs-lookup"><span data-stu-id="d5024-103">Beginner's guide</span></span>

<span data-ttu-id="d5024-104">想要开始构建自己的跨平台 Office 扩展？</span><span class="sxs-lookup"><span data-stu-id="d5024-104">Want to get started building your own cross-platform Office extensions?</span></span> <span data-ttu-id="d5024-105">以下步骤显示了需要先阅读的内容、要安装的工具以及要完成的推荐教程。</span><span class="sxs-lookup"><span data-stu-id="d5024-105">The following steps show you what to read first, what tools to install, and recommended tutorials to complete.</span></span>

> [!NOTE]
> <span data-ttu-id="d5024-106">如果你已熟知如何创建适用于 Office 的 VSTO 加载项，建议直接转到 [VSTO 加载项开发人员指南](learning-path-transition.md)（该文章是本文中信息的超集）。</span><span class="sxs-lookup"><span data-stu-id="d5024-106">If you're experienced in creating VSTO add-ins for Office, we recommend that you immediately turn to [VSTO add-in developer's guide](learning-path-transition.md), which is a superset of the information in this article.</span></span>

## <a name="step-0-prerequisites"></a><span data-ttu-id="d5024-107">步骤 0：先决条件</span><span class="sxs-lookup"><span data-stu-id="d5024-107">Step 0: Prerequisites</span></span>

- <span data-ttu-id="d5024-108">Office 加载项本质上是嵌入在 Office 中的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="d5024-108">Office Add-ins are essentially web applications embedded in Office.</span></span> <span data-ttu-id="d5024-109">因此，你首先应该对 Web 应用程序以及如何在 Web 上托管它们有基本的了解。</span><span class="sxs-lookup"><span data-stu-id="d5024-109">So, you should first have a basic understanding of web applications and how they are hosted on the web.</span></span> <span data-ttu-id="d5024-110">Internet、书籍和在线课程提供了有关它的大量信息。</span><span class="sxs-lookup"><span data-stu-id="d5024-110">There is an enormous amount of information about this on the Internet, in books, and in online courses.</span></span> <span data-ttu-id="d5024-111">如果你根本不了解 Web 应用程序，那么一个很好的开始方法是在</span><span class="sxs-lookup"><span data-stu-id="d5024-111">A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?"</span></span> <span data-ttu-id="d5024-112">必应上搜索“什么是 Web 应用程序？”。</span><span class="sxs-lookup"><span data-stu-id="d5024-112">on Bing.</span></span>
- <span data-ttu-id="d5024-113">创建 Office 加载项时将使用的主要编程语言是 JavaScript 或 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="d5024-113">The primary programming language you will use in creating Office Add-ins is JavaScript or TypeScript.</span></span> <span data-ttu-id="d5024-114">可将 TypeScript 视为 JavaScript 的强类型版本。</span><span class="sxs-lookup"><span data-stu-id="d5024-114">You can think of TypeScript as a strongly-typed version of JavaScript.</span></span> <span data-ttu-id="d5024-115">如果你不熟悉这两种语言，但是你有使用 VBA、VB.Net、C# 的经验，则你可能会发现 TypeScript 更容易学习。</span><span class="sxs-lookup"><span data-stu-id="d5024-115">If you are not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you will probably find TypeScript easier to learn.</span></span> <span data-ttu-id="d5024-116">此外，Internet、书籍和在线课程提供了有关这些语言的大量信息。</span><span class="sxs-lookup"><span data-stu-id="d5024-116">Again, there is a wealth of information about these languages on the Internet, in books, and in online courses.</span></span>

## <a name="step-1-begin-with-fundamentals"></a><span data-ttu-id="d5024-117">步骤 1：从基础知识开始</span><span class="sxs-lookup"><span data-stu-id="d5024-117">Step 1: Begin with fundamentals</span></span>

<span data-ttu-id="d5024-118">我们知道你渴望开始编码，但是在打开 IDE 或代码编辑器之前，你应该先阅读一些有关 Office 加载项的信息。</span><span class="sxs-lookup"><span data-stu-id="d5024-118">We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.</span></span>

- <span data-ttu-id="d5024-119">[Office 加载项平台概述](office-add-ins.md)：了解什么是 Office Web 加载项以及它们与扩展 Office（如 VSTO 加载项）的旧方法有何区别。</span><span class="sxs-lookup"><span data-stu-id="d5024-119">[Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.</span></span>
- <span data-ttu-id="d5024-120">[开发 Office 加载项](../develop/develop-overview.md)：概述 Office 加载项的开发和生命周期，包括工具、创建加载项 UI 以及使用 JavaScript API 与 Office 文档进行交互。</span><span class="sxs-lookup"><span data-stu-id="d5024-120">[Develop Office Add-ins](../develop/develop-overview.md): Get an overview of Office add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.</span></span>

<span data-ttu-id="d5024-121">这些文章中有许多链接，但是如果你是 Office 加载项的初学者，我们建议你在阅读完后返回此处并继续下一部分。</span><span class="sxs-lookup"><span data-stu-id="d5024-121">There are a lot of links in those articles, but if you're a beginner with Office Add-ins, we recommend that you come back here when you've read them and continue with the next section.</span></span>

## <a name="step-2-install-tools-and-create-your-first-add-in"></a><span data-ttu-id="d5024-122">步骤 2：安装工具并创建首个加载项</span><span class="sxs-lookup"><span data-stu-id="d5024-122">Step 2: Install tools and create your first add-in</span></span>

<span data-ttu-id="d5024-123">现在，你已有了大致的了解，下面需要深入了解其中一个快速入门。</span><span class="sxs-lookup"><span data-stu-id="d5024-123">You've got the big picture now, so dive in with one of our quick starts.</span></span> <span data-ttu-id="d5024-124">出于学习平台的目的，我们推荐使用 Excel 快速入门。</span><span class="sxs-lookup"><span data-stu-id="d5024-124">For purposes of learning the platform, we recommend the Excel quick start.</span></span> <span data-ttu-id="d5024-125">我们提供基于 Visual Studio 的版本以及基于 Node.js 和 Visual Studio Code 的版本。</span><span class="sxs-lookup"><span data-stu-id="d5024-125">There is a version that is based on Visual Studio and a version that is based in Node.js and Visual Studio Code.</span></span>

- [<span data-ttu-id="d5024-126">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="d5024-126">Visual Studio</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="d5024-127">Node.js 和 Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="d5024-127">Node.js and Visual Studio Code</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a><span data-ttu-id="d5024-128">步骤 3：代码</span><span class="sxs-lookup"><span data-stu-id="d5024-128">Step 3: Code</span></span>

<span data-ttu-id="d5024-129">你无法通过阅读车主手册学会开车，因此请从此 [Excel 教程](../tutorials/excel-tutorial.md)开始编码吧。</span><span class="sxs-lookup"><span data-stu-id="d5024-129">You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md).</span></span> <span data-ttu-id="d5024-130">你将使用 Office JavaScript 库和加载项清单中的一些 XML。</span><span class="sxs-lookup"><span data-stu-id="d5024-130">You'll be using the Office JavaScript library and some XML in the add-in's manifest.</span></span> <span data-ttu-id="d5024-131">无需记住任何内容，因为在后面的步骤中，你将获得关于这两者的更多背景知识。</span><span class="sxs-lookup"><span data-stu-id="d5024-131">There's no need to memorize anything, because you'll be getting more background about both in a later steps.</span></span>

## <a name="step-4-understand-the-javascript-library"></a><span data-ttu-id="d5024-132">步骤 4：了解 JavaScript 库</span><span class="sxs-lookup"><span data-stu-id="d5024-132">Step 4: Understand the JavaScript library</span></span>

<span data-ttu-id="d5024-133">首先，通过来自 Microsoft Learn 的本教程大致了解 Office JavaScript 库：[了解 Office JavaScript API](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index)。</span><span class="sxs-lookup"><span data-stu-id="d5024-133">First, get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).</span></span>

<span data-ttu-id="d5024-134">然后，使用我们的 [Script Lab 工具](explore-with-script-lab.md)（一种用于运行和探索 API 的沙箱）来探索 Office JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="d5024-134">Then explore the Office JavaScript APIs with our [the Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.</span></span>

## <a name="step-5-understand-the-manifest"></a><span data-ttu-id="d5024-135">步骤 5：了解清单</span><span class="sxs-lookup"><span data-stu-id="d5024-135">Step 5: Understand the manifest</span></span>

<span data-ttu-id="d5024-136">在 [Office 加载项 XML 清单](../develop/add-in-manifests.md)中了解加载项清单的用途以及有关其 XML 标记的简介。</span><span class="sxs-lookup"><span data-stu-id="d5024-136">Get an understanding of the purposes of the add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="d5024-137">后续步骤</span><span class="sxs-lookup"><span data-stu-id="d5024-137">Next Steps</span></span>

<span data-ttu-id="d5024-138">恭喜你完成了初学者的 Office 加载项学习之路！</span><span class="sxs-lookup"><span data-stu-id="d5024-138">Congratulations on finishing the beginner's learning path for Office Add-ins!</span></span> <span data-ttu-id="d5024-139">以下是进一步探索我们的文档的一些建议：</span><span class="sxs-lookup"><span data-stu-id="d5024-139">Here are some suggestions for further exploration of our documentation:</span></span>

- <span data-ttu-id="d5024-140">其他 Office 应用程序的教程或快速入门：</span><span class="sxs-lookup"><span data-stu-id="d5024-140">Tutorials or quick starts for other Office applications:</span></span>

  - [<span data-ttu-id="d5024-141">OneNote 快速入门</span><span class="sxs-lookup"><span data-stu-id="d5024-141">OneNote quick start</span></span>](../quickstarts/onenote-quickstart.md)
  - [<span data-ttu-id="d5024-142">Outlook 教程</span><span class="sxs-lookup"><span data-stu-id="d5024-142">Outlook tutorial</span></span>](/outlook/add-ins/addin-tutorial)
  - [<span data-ttu-id="d5024-143">PowerPoint 教程</span><span class="sxs-lookup"><span data-stu-id="d5024-143">PowerPoint tutorial</span></span>](../tutorials/powerpoint-tutorial.md)
  - [<span data-ttu-id="d5024-144">Project 快速入门</span><span class="sxs-lookup"><span data-stu-id="d5024-144">Project quick start</span></span>](../quickstarts/project-quickstart.md)
  - [<span data-ttu-id="d5024-145">Word 教程</span><span class="sxs-lookup"><span data-stu-id="d5024-145">Word tutorial</span></span>](../tutorials/word-tutorial.md)

- <span data-ttu-id="d5024-146">其他重要主题：</span><span class="sxs-lookup"><span data-stu-id="d5024-146">Other important subjects:</span></span>

  - [<span data-ttu-id="d5024-147">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="d5024-147">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
  - [<span data-ttu-id="d5024-148">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="d5024-148">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
  - [<span data-ttu-id="d5024-149">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="d5024-149">Design Office Add-ins</span></span>](../design/add-in-design.md)
  - [<span data-ttu-id="d5024-150">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="d5024-150">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
  - [<span data-ttu-id="d5024-151">部署和发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="d5024-151">Deploy and publish Office Add-ins</span></span>](../publish/publish.md)
  - [<span data-ttu-id="d5024-152">资源</span><span class="sxs-lookup"><span data-stu-id="d5024-152">Resources</span></span>](../resources/resources-links-help.md)
  - [<span data-ttu-id="d5024-153">了解 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="d5024-153">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)