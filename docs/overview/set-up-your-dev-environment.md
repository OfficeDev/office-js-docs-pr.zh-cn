---
title: 设置开发环境
description: 设置开发人员环境以生成 Office 加载项。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 1dd0cc6bb035a0274e36fe9916dcd2481bdf0b39
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234126"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="89ea8-103">设置开发环境</span><span class="sxs-lookup"><span data-stu-id="89ea8-103">Set up your development environment</span></span>

<span data-ttu-id="89ea8-104">本指南可帮助你设置工具，以便你可以按照我们的快速入门或教程创建 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="89ea8-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="89ea8-105">你需要从下面的列表中安装工具。</span><span class="sxs-lookup"><span data-stu-id="89ea8-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="89ea8-106">如果已安装这些，则已准备好开始快速入门，例如 [此 Excel React 快速入门](../quickstarts/excel-quickstart-react.md)。</span><span class="sxs-lookup"><span data-stu-id="89ea8-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="89ea8-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="89ea8-107">Node.js</span></span>
- <span data-ttu-id="89ea8-108">npm</span><span class="sxs-lookup"><span data-stu-id="89ea8-108">npm</span></span>
- <span data-ttu-id="89ea8-109">包含 Office 订阅版本的 Microsoft 365 帐户</span><span class="sxs-lookup"><span data-stu-id="89ea8-109">A Microsoft 365 account which includes the subscription version of Office</span></span>
- <span data-ttu-id="89ea8-110">你选择的代码编辑器</span><span class="sxs-lookup"><span data-stu-id="89ea8-110">A code editor of your choice</span></span>

<span data-ttu-id="89ea8-111">本指南假定你了解如何使用命令行工具。</span><span class="sxs-lookup"><span data-stu-id="89ea8-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="89ea8-112">安装 Node.js</span><span class="sxs-lookup"><span data-stu-id="89ea8-112">Install Node.js</span></span>

<span data-ttu-id="89ea8-113">Node.js JavaScript 运行时，你需要开发新式 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="89ea8-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="89ea8-114">通过Node.js[网站下载建议的最新版本来安装。](https://nodejs.org)</span><span class="sxs-lookup"><span data-stu-id="89ea8-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="89ea8-115">按照操作系统的安装说明操作。</span><span class="sxs-lookup"><span data-stu-id="89ea8-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="89ea8-116">安装 npm</span><span class="sxs-lookup"><span data-stu-id="89ea8-116">Install npm</span></span>

<span data-ttu-id="89ea8-117">npm 是一个开源软件注册表，可从中下载用于开发 Office 加载项的程序包。</span><span class="sxs-lookup"><span data-stu-id="89ea8-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="89ea8-118">若要安装 npm，请运行命令行中的以下命令：</span><span class="sxs-lookup"><span data-stu-id="89ea8-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="89ea8-119">若要检查是否已安装 npm 并查看安装的版本，请在命令行中运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="89ea8-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="89ea8-120">你可能希望使用节点版本管理器来允许你在多个版本的 Node.js 和 npm 之间切换，但这不是严格必需的。</span><span class="sxs-lookup"><span data-stu-id="89ea8-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="89ea8-121">有关如何执行此操作的详细信息，请参阅 [npm 的说明](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。</span><span class="sxs-lookup"><span data-stu-id="89ea8-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-microsoft-365"></a><span data-ttu-id="89ea8-122">获取 Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="89ea8-122">Get Microsoft 365</span></span>

<span data-ttu-id="89ea8-123">如果你还没有 Microsoft 365 帐户，可以通过加入 Microsoft 365 开发人员计划获取包含所有 Office 应用的 90 天免费可续订的 [Microsoft 365 订阅](https://developer.microsoft.com/office/dev-program)。</span><span class="sxs-lookup"><span data-stu-id="89ea8-123">If you don't already have a Microsoft 365 account, you can get a free, 90-day renewable Microsoft 365 subscription that includes all Office apps by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="89ea8-124">安装代码编辑器</span><span class="sxs-lookup"><span data-stu-id="89ea8-124">Install a code editor</span></span>

<span data-ttu-id="89ea8-125">若要生成 Web 部件，可以使用任何支持客户端开发的代码编辑器或 IDE，如：</span><span class="sxs-lookup"><span data-stu-id="89ea8-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="89ea8-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="89ea8-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="89ea8-127">Atom</span><span class="sxs-lookup"><span data-stu-id="89ea8-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="89ea8-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="89ea8-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="89ea8-129">后续步骤</span><span class="sxs-lookup"><span data-stu-id="89ea8-129">Next steps</span></span>

<span data-ttu-id="89ea8-130">尝试创建自己的外接程序或使用 Script Lab 试用内置示例。</span><span class="sxs-lookup"><span data-stu-id="89ea8-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="89ea8-131">创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="89ea8-131">Create an Office Add-in</span></span>

<span data-ttu-id="89ea8-132">可完成 [5 分钟快速入门](../index.yml)，快速创建适合 Excel、OneNote、Outlook、PowerPoint、Project 或 Word 的基本加载项。</span><span class="sxs-lookup"><span data-stu-id="89ea8-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.yml).</span></span> <span data-ttu-id="89ea8-133">如果你之前已完成快速入门，并且想要创建更复杂一些的加载项，请尝试本[教程](../index.yml)。</span><span class="sxs-lookup"><span data-stu-id="89ea8-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.yml).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="89ea8-134">使用 Script Lab 了解 API</span><span class="sxs-lookup"><span data-stu-id="89ea8-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="89ea8-135">了解 [Script Lab](explore-with-script-lab.md) 中的内置示例库，熟悉 Office JavaScript API 的功能。</span><span class="sxs-lookup"><span data-stu-id="89ea8-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="89ea8-136">另请参阅</span><span class="sxs-lookup"><span data-stu-id="89ea8-136">See also</span></span>

- [<span data-ttu-id="89ea8-137">Office 加载项的核心概念</span><span class="sxs-lookup"><span data-stu-id="89ea8-137">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="89ea8-138">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="89ea8-138">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="89ea8-139">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="89ea8-139">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="89ea8-140">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="89ea8-140">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="89ea8-141">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="89ea8-141">Publish Office Add-ins</span></span>](../publish/publish.md)
- [<span data-ttu-id="89ea8-142">了解 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="89ea8-142">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)