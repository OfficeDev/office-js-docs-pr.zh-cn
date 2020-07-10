---
title: 设置开发环境
description: 设置开发人员环境以生成 Office 外接程序
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: af59fb644d1001deb74615d6ced294ad77cbf4e6
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094006"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="df691-103">设置开发环境</span><span class="sxs-lookup"><span data-stu-id="df691-103">Set up your development environment</span></span>

<span data-ttu-id="df691-104">本指南可帮助您设置工具，以便您可以按照快速入门或教程创建 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="df691-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="df691-105">你将需要安装以下列表中的工具。</span><span class="sxs-lookup"><span data-stu-id="df691-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="df691-106">如果已安装了这些安装，则可以开始快速启动，例如此 Excel 会对[快速启动做出反应](../quickstarts/excel-quickstart-react.md)。</span><span class="sxs-lookup"><span data-stu-id="df691-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="df691-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="df691-107">Node.js</span></span>
- <span data-ttu-id="df691-108">npm</span><span class="sxs-lookup"><span data-stu-id="df691-108">npm</span></span>
- <span data-ttu-id="df691-109">包含 Office 订阅版本的 Microsoft 365 帐户</span><span class="sxs-lookup"><span data-stu-id="df691-109">A Microsoft 365 account which includes the subscription version of Office</span></span>
- <span data-ttu-id="df691-110">您选择的代码编辑器</span><span class="sxs-lookup"><span data-stu-id="df691-110">A code editor of your choice</span></span>

<span data-ttu-id="df691-111">本指南假定您知道如何使用命令行工具。</span><span class="sxs-lookup"><span data-stu-id="df691-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="df691-112">安装 Node.js</span><span class="sxs-lookup"><span data-stu-id="df691-112">Install Node.js</span></span>

<span data-ttu-id="df691-113">Node.js 是开发新式 Office 外接程序所需的 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="df691-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="df691-114">通过[从网站下载最新的推荐版本](https://nodejs.org)来安装 Node.js。</span><span class="sxs-lookup"><span data-stu-id="df691-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="df691-115">按照操作系统的安装说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="df691-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="df691-116">安装 npm</span><span class="sxs-lookup"><span data-stu-id="df691-116">Install npm</span></span>

<span data-ttu-id="df691-117">npm 是一个开放源代码软件注册表，可从中下载用于开发 Office 外接程序的程序包。</span><span class="sxs-lookup"><span data-stu-id="df691-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="df691-118">若要安装 npm，请在命令行中运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="df691-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="df691-119">若要检查是否已安装了 npm 并查看已安装的版本，请在命令行中运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="df691-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="df691-120">您可能希望使用节点版本管理器，以允许在多个版本的 Node.js 和 npm 之间进行切换，但这并不是绝对必要的。</span><span class="sxs-lookup"><span data-stu-id="df691-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="df691-121">有关如何执行此操作的详细信息，[请参阅 npm 的说明](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。</span><span class="sxs-lookup"><span data-stu-id="df691-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="df691-122">获取 Office 365</span><span class="sxs-lookup"><span data-stu-id="df691-122">Get Office 365</span></span>

<span data-ttu-id="df691-123">如果你还没有 Microsoft 365 帐户，则可以通过加入[microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取免费的90天 renewable microsoft 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="df691-123">If you don't already have a Microsoft 365 account, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="df691-124">安装代码编辑器</span><span class="sxs-lookup"><span data-stu-id="df691-124">Install a code editor</span></span>

<span data-ttu-id="df691-125">若要生成 Web 部件，可以使用任何支持客户端开发的代码编辑器或 IDE，如：</span><span class="sxs-lookup"><span data-stu-id="df691-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="df691-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="df691-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="df691-127">Atom</span><span class="sxs-lookup"><span data-stu-id="df691-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="df691-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="df691-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="df691-129">后续步骤</span><span class="sxs-lookup"><span data-stu-id="df691-129">Next steps</span></span>

<span data-ttu-id="df691-130">尝试创建您自己的外接程序，或使用脚本实验室来尝试内置的示例。</span><span class="sxs-lookup"><span data-stu-id="df691-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="df691-131">创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="df691-131">Create an Office add-in</span></span>

<span data-ttu-id="df691-132">可完成 [5 分钟快速入门](/office/dev/add-ins/)，快速创建适合 Excel、OneNote、Outlook、PowerPoint、Project 或 Word 的基本加载项。</span><span class="sxs-lookup"><span data-stu-id="df691-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](/office/dev/add-ins/).</span></span> <span data-ttu-id="df691-133">如果你之前已完成快速入门，并且想要创建更复杂一些的加载项，请尝试本[教程](/office/dev/add-ins/)。</span><span class="sxs-lookup"><span data-stu-id="df691-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](/office/dev/add-ins/).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="df691-134">使用 Script Lab 了解 API</span><span class="sxs-lookup"><span data-stu-id="df691-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="df691-135">了解 [Script Lab](explore-with-script-lab.md) 中的内置示例库，熟悉 Office JavaScript API 的功能。</span><span class="sxs-lookup"><span data-stu-id="df691-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="df691-136">另请参阅</span><span class="sxs-lookup"><span data-stu-id="df691-136">See also</span></span>

- [<span data-ttu-id="df691-137">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="df691-137">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="df691-138">Office 加载项的核心概念</span><span class="sxs-lookup"><span data-stu-id="df691-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="df691-139">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="df691-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="df691-140">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="df691-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="df691-141">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="df691-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="df691-142">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="df691-142">Publish Office Add-ins</span></span>](../publish/publish.md)
