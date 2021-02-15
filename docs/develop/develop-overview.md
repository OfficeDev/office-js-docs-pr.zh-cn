---
title: 开发 Office 加载项
description: Office 加载项开发简介。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 0d19ec8203e7141b6667713786d790eb0a12bba2
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237887"
---
# <a name="develop-office-add-ins"></a><span data-ttu-id="b1ba0-103">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b1ba0-103">Develop Office Add-ins</span></span>

> [!TIP]
> <span data-ttu-id="b1ba0-104">阅读本文之前，请查看 [Office 加载项平台概述](../overview/office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-104">Please review [Office Add-ins platform overview](../overview/office-add-ins.md) before reading this article.</span></span>

<span data-ttu-id="b1ba0-105">所有 Office 加载项均基于 Office 加载项平台构建。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-105">All Office Add-ins are built upon the Office Add-ins platform.</span></span> <span data-ttu-id="b1ba0-106">无论构建任何加载项，你都需要了解应用程序和平台可用性、Office JavaScript API 编程模式、如何在清单文件中指定加载项的设置和功能、如何设计 UI 和用户体验等重要概念。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-106">For any add-in you build, you'll need to understand important concepts like application and platform availability, Office JavaScript API programming patterns, how to specify an add-in's settings and capabilities in the manifest file, how to design the UI and user experience, and more.</span></span> <span data-ttu-id="b1ba0-107">本文档的“**开发生命周期**” > “**开发**”部分在此介绍了这类核心开发概念。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-107">Core development concepts like these are covered here in the **Development lifecycle** > **Develop** section of the documentation.</span></span> <span data-ttu-id="b1ba0-108">在浏览与所构建的加载项（例如 [Excel](../excel/index.yml)）相对应的应用程序特定文档之前，请先查看此处的信息。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-108">Review the information here before exploring the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).</span></span>

## <a name="creating-an-office-add-in"></a><span data-ttu-id="b1ba0-109">创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b1ba0-109">Creating an Office Add-in</span></span>

<span data-ttu-id="b1ba0-110">你可通过适用于 Office 加载项的 Yeoman 生成器或 Visual Studio 来创建 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-110">You can create an Office Add-in by using the Yeoman generator for Office Add-ins or Visual Studio.</span></span>

### <a name="yeoman-generator-for-office-add-ins"></a><span data-ttu-id="b1ba0-111">适用于 Office 加载项的 Yeoman 生成器</span><span class="sxs-lookup"><span data-stu-id="b1ba0-111">Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="b1ba0-112">[](https://github.com/officedev/generator-office)可用来创建 Node.js Office 加载项项目，而后者可通过 Visual Studio Code 或任何其他编辑器进行管理。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-112">The [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office) can be used to create a Node.js Office Add-in project that can be managed with Visual Studio Code or any other editor.</span></span> <span data-ttu-id="b1ba0-113">该生成器可创建适合下述任一应用的 Office 加载项：</span><span class="sxs-lookup"><span data-stu-id="b1ba0-113">The generator can create Office Add-ins for any of the following:</span></span>

- <span data-ttu-id="b1ba0-114">Excel</span><span class="sxs-lookup"><span data-stu-id="b1ba0-114">Excel</span></span>
- <span data-ttu-id="b1ba0-115">OneNote</span><span class="sxs-lookup"><span data-stu-id="b1ba0-115">OneNote</span></span>
- <span data-ttu-id="b1ba0-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="b1ba0-116">Outlook</span></span>
- <span data-ttu-id="b1ba0-117">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b1ba0-117">PowerPoint</span></span>
- <span data-ttu-id="b1ba0-118">Project</span><span class="sxs-lookup"><span data-stu-id="b1ba0-118">Project</span></span>
- <span data-ttu-id="b1ba0-119">Word</span><span class="sxs-lookup"><span data-stu-id="b1ba0-119">Word</span></span>
- <span data-ttu-id="b1ba0-120">Excel 自定义函数</span><span class="sxs-lookup"><span data-stu-id="b1ba0-120">Excel custom functions</span></span>

<span data-ttu-id="b1ba0-121">你可选择使用 HTML、CSS 和 JavaScript 创建该项目，也可使用 Angular 或 React 进行创建。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-121">You can choose to create the project using HTML, CSS and JavaScript, or using Angular or React.</span></span> <span data-ttu-id="b1ba0-122">此外，无论选择哪种框架，都可在 JavaScript 和 Typescript 之间进行选择。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-122">For whichever framework you choose, you can choose between JavaScript and Typescript as well.</span></span> <span data-ttu-id="b1ba0-123">有关使用 Yeoman 生成器创建加载项的详细信息，请参阅[使用 Visual Studio Code 开发 Office 加载项](../develop/develop-add-ins-vscode.md)。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-123">For more information about creating add-ins with the Yeoman generator, see [Develop Office Add-ins with Visual Studio Code](../develop/develop-add-ins-vscode.md).</span></span>

### <a name="visual-studio"></a><span data-ttu-id="b1ba0-124">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="b1ba0-124">Visual Studio</span></span>

<span data-ttu-id="b1ba0-125">Visual Studio 可用于创建适用于 Excel、Outlook、Word 和 PowerPoint 的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-125">Visual Studio can be used to create Office Add-ins for Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="b1ba0-126">Office 加载项项目是作为 Visual Studio 解决方案的一部分创建的，它使用 HTML、CSS 和 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-126">An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript.</span></span> <span data-ttu-id="b1ba0-127">有关使用 Visual Studio 创建加载项的详细信息，请参阅[使用 Visual Studio 开发 Office 加载项](../develop/develop-add-ins-visual-studio.md)。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-127">For more information about creating add-ins with Visual Studio, see [Develop Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## <a name="understanding-the-two-parts-of-an-office-add-in"></a><span data-ttu-id="b1ba0-128">了解 Office 加载项的两个部分</span><span class="sxs-lookup"><span data-stu-id="b1ba0-128">Understanding the two parts of an Office Add-in</span></span>

<span data-ttu-id="b1ba0-129">Office 加载项由两部分组成：</span><span class="sxs-lookup"><span data-stu-id="b1ba0-129">An Office Add-in consists of two parts:</span></span>

- <span data-ttu-id="b1ba0-130">加载项清单（XML 文件），它定义了加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-130">The add-in manifest (an XML file) that defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="b1ba0-131">Web 应用程序，它定义了加载项组件的 UI 和功能，例如任务窗格、内容加载项和对话框。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-131">The web application that defines the UI and functionality of add-in components such as task panes, content add-ins, and dialog boxes.</span></span>

<span data-ttu-id="b1ba0-132">Web 应用程序使用 Office JavaScript API 来与其中在运行加载项的 Office 文档中的内容进行交互。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-132">The web application uses the Office JavaScript API to interact with content in the Office document where the add-in is running.</span></span> <span data-ttu-id="b1ba0-133">你的加载项还可执行 Web 应用程序通常可实现的其他操作，例如调用外部 Web 服务和简化用户身份验证等等。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-133">Your add-in can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.</span></span>

### <a name="defining-an-add-ins-settings-and-capabilities"></a><span data-ttu-id="b1ba0-134">定义加载项的设置和功能</span><span class="sxs-lookup"><span data-stu-id="b1ba0-134">Defining an add-in's settings and capabilities</span></span>

<span data-ttu-id="b1ba0-135">Office 加载项的清单是一个 XML 文件，它定义了加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-135">An Office Add-in's manifest (an XML file) defines the settings and capabilities of the add-in.</span></span> <span data-ttu-id="b1ba0-136">你需配置清单来指定如下内容：</span><span class="sxs-lookup"><span data-stu-id="b1ba0-136">You'll configure the manifest to specify things such as:</span></span>

- <span data-ttu-id="b1ba0-137">描述加载项的元数据（例如 ID、版本、说明、显示名称和默认区域设置）。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-137">Metadata that describes the add-in (for example, ID, version, description, display name, default locale).</span></span>
- <span data-ttu-id="b1ba0-138">将在其中运行加载项的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-138">Office applications where the add-in will run.</span></span>
- <span data-ttu-id="b1ba0-139">加载项所需的权限。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-139">Permissions that the add-in requires.</span></span>
- <span data-ttu-id="b1ba0-140">加载项与 Office 集成的方式，包括与加载项创建的自定义选项卡和功能区按钮等自定义 UI 的集成。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-140">How the add-in integrates with Office, including any custom UI that the add-in creates (for example, custom tabs, ribbon buttons).</span></span>
- <span data-ttu-id="b1ba0-141">加载项对品牌和命令图标使用的图像的位置。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-141">Location of images that the add-in uses for branding and command iconography.</span></span>
- <span data-ttu-id="b1ba0-142">加载项的尺寸（例如内容加载项的尺寸、Outlook 加载项请求的高度）。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-142">Dimensions of the add-in (for example, dimensions for content add-ins, requested height for Outlook add-ins).</span></span>
- <span data-ttu-id="b1ba0-143">指定何时在消息或约会上下文中激活加载项的规则（仅限 Outlook 加载项）。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-143">Rules that specify when the add-in activates in the context of a message or appointment (for Outlook add-ins only).</span></span>

<span data-ttu-id="b1ba0-144">有关清单的详细信息，请参阅 [Office 加载项 XML 清单](add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-144">For detailed information about the manifest, see [Office Add-ins XML manifest](add-in-manifests.md).</span></span>

### <a name="interacting-with-content-in-an-office-document"></a><span data-ttu-id="b1ba0-145">与 Office 文档中的内容交互</span><span class="sxs-lookup"><span data-stu-id="b1ba0-145">Interacting with content in an Office document</span></span>

<span data-ttu-id="b1ba0-146">Office 加载项可使用 Office JavaScript API 来与其中在运行加载项的 Office 文档中的内容进行交互。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-146">An Office Add-in can use the Office JavaScript APIs to interact with content in the Office document where the add-in is running.</span></span>

#### <a name="accessing-the-office-javascript-api-library"></a><span data-ttu-id="b1ba0-147">访问 Office JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="b1ba0-147">Accessing the Office JavaScript API library</span></span>

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a><span data-ttu-id="b1ba0-148">API 模型</span><span class="sxs-lookup"><span data-stu-id="b1ba0-148">API models</span></span>

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a><span data-ttu-id="b1ba0-149">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b1ba0-149">API requirement sets</span></span>

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

#### <a name="exploring-apis-with-script-lab"></a><span data-ttu-id="b1ba0-150">使用 Script Lab 了解 API</span><span class="sxs-lookup"><span data-stu-id="b1ba0-150">Exploring APIs with Script Lab</span></span>

<span data-ttu-id="b1ba0-151">Script Lab 是一款加载项，在 Excel 或 Word 等 Office 程序中工作时，你可用它来了解 Office JavaScript API 和运行代码片段。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-151">Script Lab is an add-in that enables you to explore the Office JavaScript API and run code snippets while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="b1ba0-152">该工具通过 [AppSource](https://appsource.microsoft.com/product/office/WA104380862) 免费提供，随附在你的开发工具包中，在你建立希望加载项中拥有的功能原型和验证该功能时非常有用。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-152">It's available for free via [AppSource](https://appsource.microsoft.com/product/office/WA104380862) and is a useful tool to include in your development toolkit as you prototype and verify the functionality you want in your add-in.</span></span> <span data-ttu-id="b1ba0-153">在 Script Lab 中，你可访问内置示例库以快速试用 API，甚至还可将示例用作你自己的代码的起点。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-153">In Script Lab, you can access a library of built-in samples to quickly try out APIs or even use a sample as the starting point for your own code.</span></span>

<span data-ttu-id="b1ba0-154">下面时长一分钟的视频展示了 Script Lab 的实际运行情况。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-154">The following one-minute video shows Script Lab in action.</span></span>

<span data-ttu-id="b1ba0-155">[![显示 Script Lab 在 Excel、Word 和 PowerPoint 中运行的短视频](../images/screenshot-wide-youtube.png 'Script Lab 预览视频')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="b1ba0-155">[![Short video that shows Script Lab running in Excel, Word, and PowerPoint](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

<span data-ttu-id="b1ba0-156">有关 Script Lab 的详细信息，请参阅[使用 Script Lab 了解 Office JavaScript API](../overview/explore-with-script-lab.md)。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-156">For more information about Script Lab, see [Explore Office JavaScript APIs using Script Lab](../overview/explore-with-script-lab.md).</span></span>

## <a name="extending-the-office-ui"></a><span data-ttu-id="b1ba0-157">扩展 Office UI</span><span class="sxs-lookup"><span data-stu-id="b1ba0-157">Extending the Office UI</span></span>

<span data-ttu-id="b1ba0-158">Office 加载项可使用加载项命令和 HTML 容器（如任务窗格、内容加载项或对话框）来扩展 Office UI。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-158">An Office Add-in can extend the Office UI by using add-in commands and HTML containers such as task panes, content add-ins, or dialog boxes.</span></span>

- <span data-ttu-id="b1ba0-159">[加载项命令](../design/add-in-commands.md)可用于向 Office 中的默认功能区添加自定义选项卡、按钮和菜单，或者扩展当用户右键单击 Office 文档中的文本或 Excel 中的对象时显示的默认上下文菜单。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-159">[Add-in commands](../design/add-in-commands.md) can be used to add custom tabs, buttons, and menus to the default ribbon in Office, or to extend the default context menu that appears when users right-click text in an Office document or an object in Excel.</span></span> <span data-ttu-id="b1ba0-160">当用户选择加载项命令时，他们将启动该加载项命令指定的任务，例如运行 JavaScript 代码、打开任务窗格或启动对话框。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-160">When users select an add-in command, they initiate the task that the add-in command specifies, such as running JavaScript code, opening a task pane, or launching a dialog box.</span></span>

- <span data-ttu-id="b1ba0-161">[任务窗格](../design/task-pane-add-ins.md)、[内容加载项](../design/content-add-ins.md)和[对话框](../design/dialog-boxes.md)等 HTML 容器可用于显示自定义 UI 和探索 Office 应用程序中的附加功能。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-161">HTML containers like [task panes](../design/task-pane-add-ins.md), [content add-ins](../design/content-add-ins.md), and [dialog boxes](../design/dialog-boxes.md) can be used to display custom UI and expose additional functionality within an Office application.</span></span> <span data-ttu-id="b1ba0-162">每个任务窗格、内容加载项或对话框的内容和功能派生自你指定的网页。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-162">The content and functionality of each task pane, content add-in, or dialog box derives from a web page that you specify.</span></span> <span data-ttu-id="b1ba0-163">这些网页可使用 Office JavaScript API 来与其中正在运行加载项的 Office 文档中的内容进行交互，还可执行网页通常可实现的其他操作，例如调用外部 Web 服务和简化用户身份验证等等。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-163">Those web pages can use the Office JavaScript API to interact with content in the Office document where the add-in is running, and can also do other things that web pages typically do, like call external web services, facilitate user authentication, and more.</span></span>

<span data-ttu-id="b1ba0-164">下图显示功能区中有一个加载项命令、文档右侧有一个任务窗格，且文档上方有一个对话框或内容加载项。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-164">The following image shows an add-in command in the ribbon, a task pane to the right of the document, and a dialog box or content add-in over the document.</span></span>

![显示 Office 文档中的功能区内加载项命令、任务窗格、对话框/内容加载项的图表](../images/add-in-ui-elements.png)

<span data-ttu-id="b1ba0-166">要详细了解如何扩展 Office UI 和设计加载项的 UX，请参阅 [Office 加载项的 Office UI 元素](../design/interface-elements.md)。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-166">For more information about extending the Office UI and designing the add-in's UX, see [Office UI elements for Office Add-ins](../design/interface-elements.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="b1ba0-167">后续步骤</span><span class="sxs-lookup"><span data-stu-id="b1ba0-167">Next steps</span></span>

<span data-ttu-id="b1ba0-168">本文概述了创建 Office 加载项的不同方法、介绍了外接程序扩展 Office UI 的方法，描述了 API 集,介绍了 Script Lab（一种用来了解 Office JavaScript API 和建立加载项功能原型的宝贵工具）。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-168">This article has outlined the different ways to create Office Add-ins, introduced the ways that an add-in can extend the Office UI, described the API sets, and introduced Script Lab as a valuable tool for exploring Office JavaScript APIs and prototyping add-in functionality.</span></span> <span data-ttu-id="b1ba0-169">现在，你了解这一介绍性信息，请考虑沿着以下学习路径继续你的 Office 加载项之旅。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-169">Now that you've explored this introductory information, consider continuing your Office Add-ins journey along the following paths.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="b1ba0-170">创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b1ba0-170">Create an Office Add-in</span></span>

<span data-ttu-id="b1ba0-171">可完成 [5 分钟快速入门](../index.yml)，快速创建适合 Excel、OneNote、Outlook、PowerPoint、Project 或 Word 的基本加载项。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-171">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.yml).</span></span> <span data-ttu-id="b1ba0-172">如果你之前已完成快速入门，并且想要创建更复杂一些的加载项，请尝试本[教程](../index.yml)。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-172">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.yml).</span></span>

### <a name="learn-more"></a><span data-ttu-id="b1ba0-173">了解详细信息</span><span class="sxs-lookup"><span data-stu-id="b1ba0-173">Learn more</span></span>

<span data-ttu-id="b1ba0-174">查看此文档，详细了解如何开发、测试和发布 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-174">Learn more about developing, testing, and publishing Office Add-ins by exploring this documentation.</span></span>

> [!TIP]
> <span data-ttu-id="b1ba0-175">对于你构建的任何加载项，都可查看本文档的[开发生命周期](../overview/core-concepts-office-add-ins.md)部分中的信息，还可查看与你要构建的加载项类型（例如 [Excel](../excel/index.yml)）相对应的应用程序特定部分中的信息。</span><span class="sxs-lookup"><span data-stu-id="b1ba0-175">For any add-in that you build, you'll use information in the [Development lifecycle](../overview/core-concepts-office-add-ins.md) section of this documentation, along with information in the application-specific section that corresponds to the type of add-in you're building (for example, [Excel](../excel/index.yml)).</span></span>

## <a name="see-also"></a><span data-ttu-id="b1ba0-176">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b1ba0-176">See also</span></span>

- [<span data-ttu-id="b1ba0-177">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="b1ba0-177">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="b1ba0-178">了解 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="b1ba0-178">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
- [<span data-ttu-id="b1ba0-179">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b1ba0-179">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="b1ba0-180">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b1ba0-180">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="b1ba0-181">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b1ba0-181">Publish Office Add-ins</span></span>](../publish/publish.md)