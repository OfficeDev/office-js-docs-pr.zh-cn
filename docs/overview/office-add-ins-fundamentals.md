---
title: 构建 Office 加载项
description: Office 加载项开发简介。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: e0deeebb3a1c8761217a9fe33a3ef04a945b2cff
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915019"
---
# <a name="building-office-add-ins"></a><span data-ttu-id="4d4fe-103">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-103">Building Office Add-ins</span></span>

> [!TIP]
> <span data-ttu-id="4d4fe-104">阅读本文之前，请查看 [Office 加载项平台概述](office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-104">Please review [Office Add-ins platform overview](office-add-ins.md) before reading this article.</span></span>

<span data-ttu-id="4d4fe-105">Office 加载项可扩展 Office 应用程序的 UI 和功能，并与 Office 文档中的内容交互。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-105">Office Add-ins extend the UI and functionality of Office applications and interact with content in Office documents.</span></span> <span data-ttu-id="4d4fe-106">你将使用熟悉的 Web 技术创建 Office 加载项来扩展 Word、Excel、PowerPoint、OneNote、Project 或 Outlook 并与之交互。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-106">You'll use familiar web technologies to create Office Add-ins that extend and interact with Word, Excel, PowerPoint, OneNote, Project, or Outlook.</span></span> <span data-ttu-id="4d4fe-107">你构建的加载项可跨多个平台在 Office 中运行，包括 Windows、Mac、iPad 和在浏览器中。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-107">The add-ins you build can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span> <span data-ttu-id="4d4fe-108">本文简要介绍了如何开发 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-108">This article provides an introduction to developing Office Add-ins.</span></span>

## <a name="creating-an-office-add-in"></a><span data-ttu-id="4d4fe-109">创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-109">Creating an Office Add-in</span></span> 

<span data-ttu-id="4d4fe-110">你可通过适用于 Office 加载项的 Yeoman 生成器或 Visual Studio 来创建 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-110">You can create an Office Add-in by using the Yeoman generator for Office Add-ins or Visual Studio.</span></span>

### <a name="yeoman-generator-for-office-add-ins"></a><span data-ttu-id="4d4fe-111">适用于 Office 加载项的 Yeoman 生成器</span><span class="sxs-lookup"><span data-stu-id="4d4fe-111">Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="4d4fe-112">[](https://github.com/officedev/generator-office)可用来创建 Node.js Office 加载项项目，而后者可通过 Visual Studio Code 或任何其他编辑器进行管理。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-112">The [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office) can be used to create a Node.js Office Add-in project that can be managed with Visual Studio Code or any other editor.</span></span> <span data-ttu-id="4d4fe-113">该生成器可创建适合下述任一应用的 Office 加载项：</span><span class="sxs-lookup"><span data-stu-id="4d4fe-113">The generator can create Office Add-ins for any of the following:</span></span>

- <span data-ttu-id="4d4fe-114">Excel</span><span class="sxs-lookup"><span data-stu-id="4d4fe-114">Excel</span></span>
- <span data-ttu-id="4d4fe-115">OneNote</span><span class="sxs-lookup"><span data-stu-id="4d4fe-115">OneNote</span></span>
- <span data-ttu-id="4d4fe-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="4d4fe-116">Outlook</span></span>
- <span data-ttu-id="4d4fe-117">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4d4fe-117">PowerPoint</span></span>
- <span data-ttu-id="4d4fe-118">Project</span><span class="sxs-lookup"><span data-stu-id="4d4fe-118">Project</span></span>
- <span data-ttu-id="4d4fe-119">Word</span><span class="sxs-lookup"><span data-stu-id="4d4fe-119">Word</span></span>
- <span data-ttu-id="4d4fe-120">Excel 自定义函数</span><span class="sxs-lookup"><span data-stu-id="4d4fe-120">Excel custom functions</span></span>

<span data-ttu-id="4d4fe-121">你可选择使用 HTML、CSS 和 JavaScript 创建该项目，也可使用 Angular 或 React 进行创建。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-121">You can choose to create the project using HTML, CSS and JavaScript, or using Angular or React.</span></span> <span data-ttu-id="4d4fe-122">此外，无论选择哪种框架，都可在 JavaScript 和 Typescript 之间进行选择。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-122">For whichever framework you choose, you can choose between JavaScript and Typescript as well.</span></span> <span data-ttu-id="4d4fe-123">有关使用 Yeoman 生成器创建加载项的详细信息，请参阅[使用 Visual Studio Code 开发 Office 加载项](../develop/develop-add-ins-vscode.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-123">For more information about creating add-ins with the Yeoman generator, see [Develop Office Add-ins with Visual Studio Code](../develop/develop-add-ins-vscode.md).</span></span>

### <a name="visual-studio"></a><span data-ttu-id="4d4fe-124">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="4d4fe-124">Visual Studio</span></span>

<span data-ttu-id="4d4fe-125">Visual Studio 可用于创建适用于 Excel、Outlook、Word 和 PowerPoint 的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-125">Visual Studio can be used to create Office Add-ins for Excel, Word, PowerPoint, or Outlook.</span></span> <span data-ttu-id="4d4fe-126">Office 加载项项目是作为 Visual Studio 解决方案的一部分创建的，它使用 HTML、CSS 和 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-126">An Office Add-in project gets created as part of a Visual Studio solution and uses HTML, CSS, and JavaScript.</span></span> <span data-ttu-id="4d4fe-127">有关使用 Visual Studio 创建加载项的详细信息，请参阅[使用 Visual Studio 开发 Office 加载项](../develop/develop-add-ins-visual-studio.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-127">For more information about creating add-ins with Visual Studio, see [Develop Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

[!include[Yeoman vs Visual Studio comparision](../includes/yeoman-generator-recommendation.md)]

## <a name="exploring-apis-with-script-lab"></a><span data-ttu-id="4d4fe-128">使用 Script Lab 了解 API</span><span class="sxs-lookup"><span data-stu-id="4d4fe-128">Exploring APIs with Script Lab</span></span>

<span data-ttu-id="4d4fe-129">Script Lab 是一款加载项，在 Excel 或 Word 等 Office 程序中工作时，你可用它来了解 Office JavaScript API 和运行代码片段。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-129">Script Lab is an add-in that enables you to explore the Office JavaScript API and run code snippets while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="4d4fe-130">该工具通过 [AppSource](https://appsource.microsoft.com/product/office/WA104380862) 免费提供，随附在你的开发工具包中，在你建立希望加载项中拥有的功能原型和验证该功能时非常有用。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-130">It's available for free via [AppSource](https://appsource.microsoft.com/product/office/WA104380862) and is a useful tool to include in your development toolkit as you prototype and verify the functionality you want in your add-in.</span></span> <span data-ttu-id="4d4fe-131">在 Script Lab 中，你可访问内置示例库以快速试用 API，甚至还可将示例用作你自己的代码的起点。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-131">In Script Lab, you can access a library of built-in samples to quickly try out APIs or even use a sample as the starting point for your own code.</span></span> 

<span data-ttu-id="4d4fe-132">下面时长一分钟的视频展示了 Script Lab 的实际运行情况。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-132">The following one-minute video shows Script Lab in action.</span></span>

<span data-ttu-id="4d4fe-133">[![展示 Script Lab 在 Excel、Word 和 PowerPoint 中运行的预览视频。](../images/screenshot-wide-youtube.png 'Script Lab 预览视频')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="4d4fe-133">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

<span data-ttu-id="4d4fe-134">有关 Script Lab 的详细信息，请参阅[使用 Script Lab 了解 Office JavaScript API](../overview/explore-with-script-lab.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-134">For more information about Script Lab, see [Explore Office JavaScript APIs using Script Lab](../overview/explore-with-script-lab.md).</span></span>

## <a name="extending-the-office-ui"></a><span data-ttu-id="4d4fe-135">扩展 Office UI</span><span class="sxs-lookup"><span data-stu-id="4d4fe-135">Extending the Office UI</span></span>

<span data-ttu-id="4d4fe-136">Office 加载项可使用加载项命令和 HTML 容器（如任务窗格、内容加载项或对话框）来扩展 Office UI。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-136">An Office Add-in can extend the Office UI by using add-in commands and HTML containers such as task panes, content add-ins, or dialog boxes.</span></span>

- <span data-ttu-id="4d4fe-137">[加载项命令](../design/add-in-commands.md)可用于向 Office 中的默认功能区添加自定义选项卡、按钮和菜单，或者扩展当用户右键单击 Office 文档中的文本或 Excel 中的对象时显示的默认上下文菜单。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-137">[Add-in commands](../design/add-in-commands.md) can be used to add custom tabs, buttons, and menus to the default ribbon in Office, or to extend the default context menu that appears when users right-click text in an Office document or an object in Excel.</span></span> <span data-ttu-id="4d4fe-138">当用户选择加载项命令时，他们将启动该加载项命令指定的任务，例如运行 JavaScript 代码、打开任务窗格或启动对话框。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-138">When users select an add-in command, they initiate the task that the add-in command specifies, such as running JavaScript code, opening a task pane, or launching a dialog box.</span></span>

- <span data-ttu-id="4d4fe-139">[任务窗格](../design/task-pane-add-ins.md)、[内容加载项](../design/content-add-ins.md)和[对话框](../design/dialog-boxes.md)等 HTML 容器可用于显示自定义 UI 和探索 Office 应用程序中的附加功能。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-139">HTML containers like [task panes](../design/task-pane-add-ins.md), [content add-ins](../design/content-add-ins.md), and [dialog boxes](../design/dialog-boxes.md) can be used to display custom UI and expose additional functionality within an Office application.</span></span> <span data-ttu-id="4d4fe-140">每个任务窗格、内容加载项或对话框的内容和功能派生自你指定的网页。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-140">The content and functionality of each task pane, content add-in, or dialog box derives from a web page that you specify.</span></span> <span data-ttu-id="4d4fe-141">这些网页可使用 Office JavaScript API 来与其中正在运行加载项的 Office 文档中的内容进行交互，还可执行网页通常可实现的其他操作，例如调用外部 Web 服务和简化用户身份验证等等。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-141">Those web pages can use the Office JavaScript API to interact with content in the Office document where the add-in is running, and can also do other things that web pages typically do, like call external web services, facilitate user authentication, and more.</span></span>

<span data-ttu-id="4d4fe-142">下图显示功能区中有一个加载项命令、文档右侧有一个任务窗格，且文档上方有一个对话框或内容加载项。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-142">The following image shows an add-in command in the ribbon, a task pane to the right of the document, and a dialog box or content add-in over the document.</span></span>

![显示 Office 文档中的功能区内加载项命令、任务窗格和对话框的图像](../images/add-in-ui-elements.png)

<span data-ttu-id="4d4fe-144">要详细了解如何扩展 Office UI，请参阅 [Office 加载项的 Office UI 元素](../design/interface-elements.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-144">For more information about extending the Office UI, see [Office UI elements for Office Add-ins](../design/interface-elements.md).</span></span>

## <a name="core-development-concepts"></a><span data-ttu-id="4d4fe-145">核心开发概念</span><span class="sxs-lookup"><span data-stu-id="4d4fe-145">Core development concepts</span></span> 

<span data-ttu-id="4d4fe-146">Office 加载项由两部分组成：</span><span class="sxs-lookup"><span data-stu-id="4d4fe-146">An Office Add-in consists of two parts:</span></span>

- <span data-ttu-id="4d4fe-147">加载项清单（XML 文件），它定义了加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-147">The add-in manifest (an XML file) that defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="4d4fe-148">Web 应用程序，它定义了加载项组件的 UI 和功能，例如任务窗格、内容加载项和对话框。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-148">The web application that defines the UI and functionality of add-in components such as task panes, content add-ins, and dialog boxes.</span></span>

<span data-ttu-id="4d4fe-149">Web 应用程序使用 Office JavaScript API 来与其中在运行加载项的 Office 文档中的内容进行交互。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-149">The web application uses the Office JavaScript API to interact with content in the Office document where the add-in is running.</span></span> <span data-ttu-id="4d4fe-150">你的加载项还可执行 Web 应用程序通常可实现的其他操作，例如调用外部 Web 服务和简化用户身份验证等等。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-150">Your add-in can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.</span></span>

### <a name="defining-an-add-ins-settings-and-capabilities"></a><span data-ttu-id="4d4fe-151">定义加载项的设置和功能</span><span class="sxs-lookup"><span data-stu-id="4d4fe-151">Defining an add-in's settings and capabilities</span></span>

<span data-ttu-id="4d4fe-152">Office 加载项的清单是一个 XML 文件，它定义了加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-152">An Office Add-in's manifest (an XML file) defines the settings and capabilities of the add-in.</span></span> <span data-ttu-id="4d4fe-153">你需配置清单来指定如下内容：</span><span class="sxs-lookup"><span data-stu-id="4d4fe-153">You'll configure the manifest to specify things such as:</span></span>

- <span data-ttu-id="4d4fe-154">描述加载项的元数据（例如 ID、版本、说明、显示名称和默认区域设置）。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-154">Metadata that describes the add-in (for example, ID, version, description, display name, default locale).</span></span>
- <span data-ttu-id="4d4fe-155">将在其中运行加载项的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-155">Office applications where the add-in will run.</span></span>
- <span data-ttu-id="4d4fe-156">加载项所需的权限。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-156">Permissions that the add-in requires.</span></span>
- <span data-ttu-id="4d4fe-157">加载项与 Office 集成的方式，包括与加载项创建的自定义选项卡和功能区按钮等自定义 UI 的集成。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-157">How the add-in integrates with Office, including any custom UI that the add-in creates (for example, custom tabs, ribbon buttons).</span></span>
- <span data-ttu-id="4d4fe-158">加载项对品牌和命令图标使用的图像的位置。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-158">Location of images that the add-in uses for branding and command iconography.</span></span>
- <span data-ttu-id="4d4fe-159">加载项的尺寸（例如内容加载项的尺寸、Outlook 加载项请求的高度）。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-159">Dimensions of the add-in (for example, dimensions for content add-ins, requested height for Outlook add-ins).</span></span>
- <span data-ttu-id="4d4fe-160">指定何时在消息或约会上下文中激活加载项的规则（仅限 Outlook 加载项）。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-160">Rules that specify when the add-in activates in the context of a message or appointment (for Outlook add-ins only).</span></span>

<span data-ttu-id="4d4fe-161">有关清单的详细信息，请参阅 [Office 加载项 XML 清单](add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-161">For detailed information about the manifest, see [Office Add-ins XML manifest](add-in-manifests.md).</span></span>

### <a name="interacting-with-content-in-an-office-document"></a><span data-ttu-id="4d4fe-162">与 Office 文档中的内容交互</span><span class="sxs-lookup"><span data-stu-id="4d4fe-162">Interacting with content in an Office document</span></span>

<span data-ttu-id="4d4fe-163">Office 加载项可使用 Office JavaScript API 来与其中在运行加载项的 Office 文档中的内容进行交互。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-163">An Office Add-in can use the Office JavaScript APIs to interact with content in the Office document where the add-in is running.</span></span> 

#### <a name="accessing-the-office-javascript-library"></a><span data-ttu-id="4d4fe-164">访问 Office JavaScript 库</span><span class="sxs-lookup"><span data-stu-id="4d4fe-164">Accessing the Office JavaScript library</span></span>

<span data-ttu-id="4d4fe-165">可通过 Office JS 内容交付网络 (CDN) 访问 Office JavaScript 库：`https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`</span><span class="sxs-lookup"><span data-stu-id="4d4fe-165">The Office JavaScript library can be accessed via the Office JS content delivery network (CDN) at: `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span></span> <span data-ttu-id="4d4fe-166">要在任何加载项的网页中使用 Office JavaScript API，必须在页面的 `<head>` 标记中的 `<script>` 标记内引用 CDN。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-166">To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> <span data-ttu-id="4d4fe-167">要使用预览版 API，请参考 CDN 上的 Office JavaScript 库预览版：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-167">To use preview APIs, reference the preview version of the Office JavaScript library on the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

<span data-ttu-id="4d4fe-168">要详细了解如何访问 Office JavaScript 库（包括如何获取 IntelliSense），请参阅[通过 JavaScript API for Office 的内容交付网络 (CDN) 引用该库](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-168">For more information about accessing the Office JavaScript library, including how to get IntelliSense, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

#### <a name="api-models"></a><span data-ttu-id="4d4fe-169">API 模型</span><span class="sxs-lookup"><span data-stu-id="4d4fe-169">API models</span></span>

<span data-ttu-id="4d4fe-170">Office JavaScript API 包含两种不同的模型：</span><span class="sxs-lookup"><span data-stu-id="4d4fe-170">The Office JavaScript APIs include two distinct models:</span></span>

- <span data-ttu-id="4d4fe-171">**主机特定的** API 提供了强类型对象，它可用于与特定 Office 应用程序的本机对象进行交互。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-171">**Host-specific** APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application.</span></span> <span data-ttu-id="4d4fe-172">例如，可使用 Excel JavaScript API 来访问工作表、区域、表格和图表等。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-172">For example, you can use the Excel JavaScript APIs to access worksheets, ranges, tables, charts, and more.</span></span> <span data-ttu-id="4d4fe-173">主机特定的 API 当前可用于 [Excel](../reference/overview/excel-add-ins-reference-overview.md)、[Word](../reference/overview/word-add-ins-reference-overview.md) 和 [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-173">Host-specific APIs are currently available for [Excel](../reference/overview/excel-add-ins-reference-overview.md), [Word](../reference/overview/word-add-ins-reference-overview.md), and [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md).</span></span> <span data-ttu-id="4d4fe-174">此 API 模型使用的是[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，你可用它在你发送给 Office 主机的每个请求中指定多个操作。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-174">This API model uses [promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) and allows you to specify multiple operations in each request you send to the Office host.</span></span> <span data-ttu-id="4d4fe-175">通过此方式批量处理操作，可大幅提升 Web 应用程序上的 Office 中的性能。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-175">Batching operations in this manner can significantly improve add-in performance in Office on the web applications.</span></span> <span data-ttu-id="4d4fe-176">主机特定的 API 是随 Office 2016 引入的，不可用于与 Office 2013 进行交互。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-176">Host-specific APIs were introduced with Office 2016 and cannot be used to interact with Office 2013.</span></span>

- <span data-ttu-id="4d4fe-177">**通用** API 可用于访问在多种类型的 Office 应用程序中都很常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-177">**Common** APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span> <span data-ttu-id="4d4fe-178">此 API 模型使用的是[回调](https://developer.mozilla.org/docs/Glossary/Callback_function)，其中你仅可在发送给 Office 主机的每个请求中指定一个操作。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-178">This API model uses [callbacks](https://developer.mozilla.org/docs/Glossary/Callback_function), where you can only specify one operation in each request you send to the Office host.</span></span> <span data-ttu-id="4d4fe-179">通用 API 是随 Office 2013 引入的，可用于与 Office 2013 或更高版本进行交互。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-179">Common APIs were introduced with Office 2013 and can be used to interact with Office 2013 or later.</span></span> <span data-ttu-id="4d4fe-180">要详细了解通用 API 对象模型（其中包括用于与 Outlook 和 PowerPoint 交互的 API），请参阅 [Office JavaScript API 对象模型](../develop/office-javascript-api-object-model.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-180">For details about the Common API object model, which includes APIs for interacting with Outlook and PowerPoint, see [Office JavaScript API object model](../develop/office-javascript-api-object-model.md).</span></span>

> [!NOTE]
> <span data-ttu-id="4d4fe-181">Excel 自定义函数在排列了计算执行优先级的唯一运行时中运行，因此使用的编程模型略有不同。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-181">Excel Custom functions run within a unique runtime that prioritizes execution of calculations, and therefore uses a slightly different programming model.</span></span> <span data-ttu-id="4d4fe-182">有关详细信息，请参阅[自定义函数体系结构](../excel/custom-functions-architecture.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-182">For details, see [Custom functions architecture](../excel/custom-functions-architecture.md).</span></span>

<span data-ttu-id="4d4fe-183">有关 Office JavaScript API 的详细信息，请参阅[了解 JavaScript API for Office](../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-183">For additional information about the Office JavaScript APIs, see [Understanding the JavaScript API for Office](../develop/understanding-the-javascript-api-for-office.md).</span></span>

#### <a name="api-requirement-sets"></a><span data-ttu-id="4d4fe-184">API 要求集</span><span class="sxs-lookup"><span data-stu-id="4d4fe-184">API requirement sets</span></span>

<span data-ttu-id="4d4fe-185">[要求集](../develop/office-versions-and-requirement-sets.md)是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-185">[Requirement sets](../develop/office-versions-and-requirement-sets.md) are named groups of API members.</span></span> <span data-ttu-id="4d4fe-186">要求集可特定于 Office 主机，例如 `ExcelApi 1.7` 要求集（一组仅可在 Excel 中使用的 API），也可常用于多台主机，例如 `DialogApi 1.1` 要求集（一组可在支持对话框 API 的任何 Office 应用程序中使用的 API）。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-186">Requirement sets can be specific to Office hosts, such as the `ExcelApi 1.7` requirement set (a set of APIs that can only be used in Excel), or common to multiple hosts, such as the `DialogApi 1.1` requirement set (a set of APIs that can be used in any Office application that supports the Dialog API).</span></span>

<span data-ttu-id="4d4fe-187">加载项可使用要求集来确定 Office 主机是否支持需要使用的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-187">Your add-in can use requirement sets to determine whether the Office host supports the API members that it needs to use.</span></span> <span data-ttu-id="4d4fe-188">有关详细信息，请参阅[指定 Office 主机和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-188">For more information about this, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>

<span data-ttu-id="4d4fe-189">要求集支持因 Office 主机、版本和平台而异。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-189">Requirement set support varies by Office host, version, and platform.</span></span> <span data-ttu-id="4d4fe-190">要详细了解每个 Office 应用程序支持的平台、要求集和通用 API，请参阅 [Office 加载项主机和平台可用性](office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-190">For detailed information about the platforms, requirement sets, and Common APIs that each Office application supports, see [Office Add-in host and platform availability](office-add-in-availability.md).</span></span>

## <a name="testing-and-debugging-an-office-add-in"></a><span data-ttu-id="4d4fe-191">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-191">Testing and debugging an Office Add-in</span></span>

<span data-ttu-id="4d4fe-192">开发加载项时，可使用一种名为_旁加载_的技术在本地测试它。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-192">As you develop your add-in, you can test it locally by using a technique known as _sideloading_.</span></span> <span data-ttu-id="4d4fe-193">加载项的旁加载过程因平台而异，在某些情况下，也因产品而异。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-193">The procedure for sideloading an add-in varies by platform, and in some cases, by product as well.</span></span> <span data-ttu-id="4d4fe-194">同样地，加载项的调试流程也因平台和产品而异。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-194">Likewise, the procedure for debugging an add-in can also vary by platform and product.</span></span> <span data-ttu-id="4d4fe-195">有关测试和调试的详细信息，请参阅[测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-195">For more information about testing and debugging, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publishing-an-office-add-in"></a><span data-ttu-id="4d4fe-196">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-196">Publishing an Office Add-in</span></span>

<span data-ttu-id="4d4fe-197">当准备好与他人共享加载项时，可使用最符合你的目标的部署方法实现这一点。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-197">When you're ready to share your add-in with others, you'll do so by using the deployment method that best meets your objectives.</span></span> <span data-ttu-id="4d4fe-198">例如，若要将加载项部署给组织内部用户，可使用集中式部署或在 SharePoint 应用目录中发布加载项。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-198">For example, to deploy an add-in to users within your organization, you might use centralized deployment or publish the add-in to a SharePoint app catalog.</span></span> <span data-ttu-id="4d4fe-199">如果想要公开共享加载项供任何人获取，可在 AppSource 中发布加载项。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-199">If you want to share your add-in publicly for anyone to obtain, you can publish the add-in to AppSource.</span></span> <span data-ttu-id="4d4fe-200">有关发布的详细信息，请参阅[部署和发布 Office 加载项](../publish/publish.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-200">For more information about publishing, see [Deploy and publish Office Add-ins](../publish/publish.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="4d4fe-201">后续步骤</span><span class="sxs-lookup"><span data-stu-id="4d4fe-201">Next steps</span></span>

<span data-ttu-id="4d4fe-202">本文概述了创建 Office 加载项的不同方法、介绍了 Script Lab（一种用来了解 Office JavaScript API 和建立加载项功能原型的宝贵工具），还描述了重要的 Office 加载项开发、测试和发布概念。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-202">This article has outlined the different ways to create Office Add-ins, introduced Script Lab as a valuable tool for exploring Office JavaScript APIs and prototyping add-in functionality, and described important Office Add-ins development, testing, and publishing concepts.</span></span> <span data-ttu-id="4d4fe-203">现在，你了解这一介绍性信息，请考虑沿着以下学习路径继续你的 Office 加载项之旅。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-203">Now that you've explored this introductory information, consider continuing your Office Add-ins journey along the following paths.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="4d4fe-204">创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-204">Create an Office add-in</span></span>

<span data-ttu-id="4d4fe-205">可完成 [5 分钟快速入门](../index.md)，快速创建适合 Excel、OneNote、Outlook、PowerPoint、Project 或 Word 的基本加载项。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-205">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.md).</span></span> <span data-ttu-id="4d4fe-206">如果你之前已完成快速入门，并且想要创建更复杂一些的加载项，请尝试本[教程](../index.md)。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-206">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.md).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="4d4fe-207">使用 Script Lab 了解 API</span><span class="sxs-lookup"><span data-stu-id="4d4fe-207">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="4d4fe-208">了解 [Script Lab](explore-with-script-lab.md) 中的内置示例库，熟悉 Office JavaScript API 的功能。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-208">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

### <a name="learn-more"></a><span data-ttu-id="4d4fe-209">了解详细信息</span><span class="sxs-lookup"><span data-stu-id="4d4fe-209">Learn more</span></span>

<span data-ttu-id="4d4fe-210">查看此文档，详细了解如何开发、测试和发布 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-210">Learn more about developing, testing, and publishing Office Add-ins by exploring this documentation.</span></span>

> [!TIP]
> <span data-ttu-id="4d4fe-211">对于你构建的任何加载项，都可查看本文档的[核心概念](core-concepts-office-add-ins.md)部分中的信息，还可查看与你要构建的加载项类型（例如 [Excel](../excel/index.md)）相对应的主机特定部分中的信息。</span><span class="sxs-lookup"><span data-stu-id="4d4fe-211">For any add-in that you build, you'll use information in the [Core concepts](core-concepts-office-add-ins.md) section of this documentation, along with information in the host-specific section that corresponds to the type of add-in you're building (for example, [Excel](../excel/index.md)).</span></span>
>
> ![显示目录的图像](../images/top-level-toc.png)

## <a name="see-also"></a><span data-ttu-id="4d4fe-213">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4d4fe-213">See also</span></span> 

- [<span data-ttu-id="4d4fe-214">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="4d4fe-214">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4d4fe-215">Office 加载项的核心概念</span><span class="sxs-lookup"><span data-stu-id="4d4fe-215">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="4d4fe-216">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-216">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="4d4fe-217">使用 Visual Studio Code 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-217">Develop Office Add-ins with Visual Studio Code</span></span>](../develop/develop-add-ins-vscode.md)
- [<span data-ttu-id="4d4fe-218">使用 Visual Studio 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-218">Develop Office Add-ins with Visual Studio Code</span></span>](../develop/develop-add-ins-visual-studio.md)
- [<span data-ttu-id="4d4fe-219">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-219">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="4d4fe-220">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-220">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="4d4fe-221">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4d4fe-221">Publish Office Add-ins</span></span>](../publish/publish.md)