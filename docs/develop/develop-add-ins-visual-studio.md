---
title: 使用 Visual Studio 开发 Office 加载项
description: 如何使用 Visual Studio 开发 Office 加载项
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: a0ad1e347f0b1013deb6a431d4a9fdcf1133d4d5
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719033"
---
# <a name="develop-office-add-ins-with-visual-studio"></a><span data-ttu-id="33239-103">使用 Visual Studio 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="33239-103">Develop Office Add-ins with Visual Studio</span></span>

<span data-ttu-id="33239-104">本文介绍如何使用 Visual Studio 开发 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="33239-104">This article describes how to use Visual Studio to develop an Office Add-in.</span></span> <span data-ttu-id="33239-105">如果你已创建加载项，则可以跳至[使用 Visual Studio 开发加载项](#develop-the-add-in-using-visual-studio)部分。</span><span class="sxs-lookup"><span data-stu-id="33239-105">If you've already created your add-in, you can skip ahead to the [Develop the add-in using Visual Studio](#develop-the-add-in-using-visual-studio) section.</span></span>

> [!NOTE]
> <span data-ttu-id="33239-106">作为使用 Visual Studio 的替代方法，你可以选择使用适用于 Office 加载项和 VS Code 的 Yeoman 生成器来创建 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="33239-106">As an alternative to using Visual Studio, you may choose to use the Yeoman generator for Office Add-ins and VS Code to create an Office Add-in.</span></span> <span data-ttu-id="33239-107">有关此选项的详细信息，请参阅[创建 Office 加载项](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in)。</span><span class="sxs-lookup"><span data-stu-id="33239-107">For more information about this choice, see [Creating an Office Add-in](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in).</span></span>

## <a name="create-the-add-in-project-using-visual-studio"></a><span data-ttu-id="33239-108">使用 Visual Studio 创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="33239-108">Create the add-in project using Visual Studio</span></span>

<span data-ttu-id="33239-109">Visual Studio 可用于创建适用于 Excel、Outlook、Word 和 PowerPoint 的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="33239-109">Visual Studio can be used to create Office Add-ins for Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="33239-110">Office 加载项项目是作为 Visual Studio 解决方案的一部分创建的，它使用 HTML、CSS 和 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="33239-110">An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript.</span></span> <span data-ttu-id="33239-111">要使用 Visual Studio 创建 Office 加载项，请按照快速入门中与你要创建的加载项相对应的说明进行操作：</span><span class="sxs-lookup"><span data-stu-id="33239-111">To create an Office Add-in with Visual Studio, follow instructions in the quick start that corresponds to the add-in you'd like to create:</span></span>

- [<span data-ttu-id="33239-112">Excel 快速入门</span><span class="sxs-lookup"><span data-stu-id="33239-112">Excel quick start</span></span>](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [<span data-ttu-id="33239-113">Outlook 快速入门</span><span class="sxs-lookup"><span data-stu-id="33239-113">Outlook quick start</span></span>](../quickstarts/outlook-quickstart.md?tabs=visualstudio)
- [<span data-ttu-id="33239-114">Word 快速入门</span><span class="sxs-lookup"><span data-stu-id="33239-114">Word quick start</span></span>](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [<span data-ttu-id="33239-115">PowerPoint 快速入门</span><span class="sxs-lookup"><span data-stu-id="33239-115">PowerPoint quick start</span></span>](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

<span data-ttu-id="33239-116">Visual Studio 不支持创建适用于 OneNote 或 Project 的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="33239-116">Visual Studio doesn't support creating Office Add-ins for OneNote or Project.</span></span> <span data-ttu-id="33239-117">要为这些主机中的任何主机创建 Office 加载项，你需要使用适用于 Office 加载项的 Yeoman 生成器，如 [OneNote 快速入门](../quickstarts/onenote-quickstart.md)或 [Project 快速入门](../quickstarts/project-quickstart.md)中所述。</span><span class="sxs-lookup"><span data-stu-id="33239-117">To create Office Add-ins for either of these hosts, you'll need to use the Yeoman generator for Office Add-ins, as described in the [OneNote quick start](../quickstarts/onenote-quickstart.md) or the [Project quick start](../quickstarts/project-quickstart.md).</span></span>

## <a name="develop-the-add-in-using-visual-studio"></a><span data-ttu-id="33239-118">使用 Visual Studio 开发加载项</span><span class="sxs-lookup"><span data-stu-id="33239-118">Develop the add-in using Visual Studio</span></span>

<span data-ttu-id="33239-119">Visual Studio 会创建一个功能受限的基本加载项。</span><span class="sxs-lookup"><span data-stu-id="33239-119">Visual Studio creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="33239-120">你可通过在 Visual Studio 中编辑[清单](add-in-manifests.md)、HTML、JavaScript 和 CSS 文件来自定义加载项。</span><span class="sxs-lookup"><span data-stu-id="33239-120">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript, and CSS files in Visual Studio.</span></span> <span data-ttu-id="33239-121">有关 Visual Studio 创建的加载项项目中的项目结构和文件的高级说明，请参阅用于指导创建加载项的快速入门中的 Visual Studio 指南。</span><span class="sxs-lookup"><span data-stu-id="33239-121">For a high-level description of the project structure and files in the add-in project that Visual Studio creates, see the Visual Studio guidance within the quick start that you completed to create your add-in.</span></span> 

> [!TIP]
> <span data-ttu-id="33239-122">由于 Office 加载项是一种 Web 应用程序，因此你至少需要具备基本的 Web 开发技能才能自定义加载项。</span><span class="sxs-lookup"><span data-stu-id="33239-122">Because an Office Add-in is a web application, you'll need at least basic web development skills to customize your add-in.</span></span> <span data-ttu-id="33239-123">如果你不熟悉 JavaScript，建议查看 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)。</span><span class="sxs-lookup"><span data-stu-id="33239-123">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

<span data-ttu-id="33239-124">要自定义加载项，你需要了解本文档的“[核心概念 > 开发](develop-overview.md)”区域中描述的概念，以及与要构建的加载项相对应的文档主机特定区域中描述的概念（例如，[Excel](../excel/index.md)）。</span><span class="sxs-lookup"><span data-stu-id="33239-124">To customize your add-in, you'll need to understand concepts described in the [Core concepts > Develop](develop-overview.md) area of this documentation, as well as concepts described in the host-specific area of documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.md)).</span></span> 

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="33239-125">测试和调试加载项</span><span class="sxs-lookup"><span data-stu-id="33239-125">Test and debug the add-in</span></span>

<span data-ttu-id="33239-126">用于测试、调试和故障排除 Office 加载项的方法因平台而异。</span><span class="sxs-lookup"><span data-stu-id="33239-126">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="33239-127">有关详细信息，请参阅[在 Visual Studio 中调试 Office 加载项](debug-office-add-ins-in-visual-studio.md)和[测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="33239-127">For more information, see [Debug Office Add-ins in Visual Studio](debug-office-add-ins-in-visual-studio.md) and [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="33239-128">发布加载项</span><span class="sxs-lookup"><span data-stu-id="33239-128">Publish the add-in</span></span>

<span data-ttu-id="33239-129">Office 加载项由一个 Web 应用程序和一个清单文件构成。</span><span class="sxs-lookup"><span data-stu-id="33239-129">An Office Add-in consists of a web application and a manifest file.</span></span> <span data-ttu-id="33239-130">Web 应用程序定义加载项的用户界面和功能，清单指定 Web 应用程序的位置并定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="33239-130">The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.</span></span>

<span data-ttu-id="33239-131">在 Visual Studio 中开发加载项时，该加载项将在本地 Web 服务器 (`localhost`) 上运行。</span><span class="sxs-lookup"><span data-stu-id="33239-131">While you're developing your add-in in Visual Studio, your add-in runs on your local web server (`localhost`).</span></span> <span data-ttu-id="33239-132">如果加载项如期工作且你已准备好发布它供其他用户访问，你需要完成以下步骤：</span><span class="sxs-lookup"><span data-stu-id="33239-132">When your add-in is working as desired and you're ready to publish it for other users to access, you'll need to complete the following steps:</span></span>

1. <span data-ttu-id="33239-133">将 Web 应用程序部署到 Web 服务器或 Web 托管服务（例如 Microsoft Azure）。</span><span class="sxs-lookup"><span data-stu-id="33239-133">Deploy the web application to a web server or web hosting service (for example, Microsoft Azure).</span></span>
2. <span data-ttu-id="33239-134">更新清单以指定已部署应用程序的 URL。</span><span class="sxs-lookup"><span data-stu-id="33239-134">Update the manifest to specify the URL of the deployed application.</span></span> 
3. <span data-ttu-id="33239-135">选择要用来[部署 Office 加载项](../publish/publish.md)的方法，再按照说明发布清单文件。</span><span class="sxs-lookup"><span data-stu-id="33239-135">Choose the method you'd like to use to [deploy your Office Add-in](../publish/publish.md), and follow the instructions to publish the manifest file.</span></span>

## <a name="see-also"></a><span data-ttu-id="33239-136">另请参阅</span><span class="sxs-lookup"><span data-stu-id="33239-136">See also</span></span>

- [<span data-ttu-id="33239-137">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="33239-137">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="33239-138">Office 加载项的核心概念</span><span class="sxs-lookup"><span data-stu-id="33239-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="33239-139">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="33239-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="33239-140">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="33239-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="33239-141">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="33239-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="33239-142">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="33239-142">Publish Office Add-ins</span></span>](../publish/publish.md)