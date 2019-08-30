---
title: 使用脚本实验室浏览 Office JavaScript API
description: 使用脚本实验室浏览 Office JS API 并建立原型功能。
ms.date: 07/05/2019
ms.topic: overview
scenarios: getting-started
localization_priority: Normal
ms.openlocfilehash: 908d27cdb5c8a7d4bc080c266cdb4d604114c42f
ms.sourcegitcommit: 49af31060aa56c1e1ec1e08682914d3cbefc3f1c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/29/2019
ms.locfileid: "36672836"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="f4e90-103">使用脚本实验室浏览 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="f4e90-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="f4e90-104">[脚本实验室外接程序](https://appsource.microsoft.com/product/office/WA104380862)(从 AppSource 中免费获取) 使您能够在使用 office 程序 (如 Excel 或 Word) 时浏览 OFFICE JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="f4e90-104">The [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862), which is available free from AppSource, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="f4e90-105">当您在外接程序中原型和验证所需功能时, 脚本实验室是一个方便的工具, 可将其添加到开发工具包中。</span><span class="sxs-lookup"><span data-stu-id="f4e90-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="f4e90-106">什么是脚本实验室？</span><span class="sxs-lookup"><span data-stu-id="f4e90-106">What is Script Lab?</span></span>

<span data-ttu-id="f4e90-107">脚本实验室是任何希望了解如何使用 Excel、Word 或 PowerPoint 中的 Office JavaScript API 开发 Office 外接程序的工具。</span><span class="sxs-lookup"><span data-stu-id="f4e90-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="f4e90-108">它提供了智能感知功能, 以便您可以查看在摩纳哥框架 (由 Visual Studio Code 使用的相同框架) 中构建的可用功能。</span><span class="sxs-lookup"><span data-stu-id="f4e90-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="f4e90-109">通过脚本实验室, 您可以访问示例库以快速试用功能, 也可以将示例用作您自己的代码的起始点。</span><span class="sxs-lookup"><span data-stu-id="f4e90-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="f4e90-110">您甚至可以使用脚本实验室尝试预览 Api。</span><span class="sxs-lookup"><span data-stu-id="f4e90-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="f4e90-111">我到目前为止听起来正常吗？</span><span class="sxs-lookup"><span data-stu-id="f4e90-111">Sounds good so far?</span></span> <span data-ttu-id="f4e90-112">查看此一分钟视频可查看脚本实验室的实际效果。</span><span class="sxs-lookup"><span data-stu-id="f4e90-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="f4e90-113">[![显示在 Excel、Word 和 PowerPoint 中运行的脚本实验室的预览视频。](../images/screenshot-wide-youtube.png '脚本实验室预览视频')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="f4e90-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="f4e90-114">关键功能</span><span class="sxs-lookup"><span data-stu-id="f4e90-114">Key features</span></span>

<span data-ttu-id="f4e90-115">脚本实验室提供了许多功能, 可帮助您探索 Office JavaScript API 和原型加载项功能。</span><span class="sxs-lookup"><span data-stu-id="f4e90-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="f4e90-116">浏览示例</span><span class="sxs-lookup"><span data-stu-id="f4e90-116">Explore samples</span></span>

<span data-ttu-id="f4e90-117">使用内置示例代码段集合快速入门, 其中展示了如何使用 API 完成任务。</span><span class="sxs-lookup"><span data-stu-id="f4e90-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="f4e90-118">您可以运行示例来即时查看任务窗格或文档中的结果, 检查示例以了解 API 的工作原理, 甚至使用示例来原型自己的外接程序。</span><span class="sxs-lookup"><span data-stu-id="f4e90-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![示例](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="f4e90-120">代码和样式</span><span class="sxs-lookup"><span data-stu-id="f4e90-120">Code and style</span></span>

<span data-ttu-id="f4e90-121">除了调用 Office JS API 的 JavaScript 或 TypeScript 代码外, 每个代码段还包含用于定义任务窗格外观的任务窗格和 CSS 内容的 HTML 标记。</span><span class="sxs-lookup"><span data-stu-id="f4e90-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="f4e90-122">您可以自定义 HTML 标记和 CSS 以在为自己的外接程序设置任务窗格设计原型时体验元素的放置和样式。</span><span class="sxs-lookup"><span data-stu-id="f4e90-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="f4e90-123">若要在代码段中调用预览 Api, 您需要更新代码段的库以使用 beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) 和预览类型定义。 `@types/office-js-preview`</span><span class="sxs-lookup"><span data-stu-id="f4e90-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="f4e90-124">此外, 某些预览 Api 仅当你注册[Office 预览体验计划](https://products.office.com/office-insider)并运行内部版本的 office 时才可访问。</span><span class="sxs-lookup"><span data-stu-id="f4e90-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://products.office.com/office-insider) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="f4e90-125">保存和共享代码段</span><span class="sxs-lookup"><span data-stu-id="f4e90-125">Save and share snippets</span></span>

<span data-ttu-id="f4e90-126">默认情况下, 在脚本实验室中打开的代码段将保存到您的浏览器缓存中。</span><span class="sxs-lookup"><span data-stu-id="f4e90-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="f4e90-127">若要永久保存代码段, 可以将其导出到[GitHub gist](https://gist.github.com)。</span><span class="sxs-lookup"><span data-stu-id="f4e90-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="f4e90-128">创建一个机密 gist 以仅用于您自己使用的代码段, 或者创建一个公用 gist (如果您计划与其他人共享它)。</span><span class="sxs-lookup"><span data-stu-id="f4e90-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![共享选项](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="f4e90-130">导入代码段</span><span class="sxs-lookup"><span data-stu-id="f4e90-130">Import snippets</span></span>

<span data-ttu-id="f4e90-131">您可以通过指定存储代码段 YAML 的公共[GitHub gist](https://gist.github.com)的 URL 或在代码段的完整 YAML 中粘贴, 将代码段导入脚本实验室。</span><span class="sxs-lookup"><span data-stu-id="f4e90-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="f4e90-132">如果其他人已通过将代码段发布到 GitHub gist 或提供代码段的 YAML, 则此功能可能对您共享其代码段的方案有用。</span><span class="sxs-lookup"><span data-stu-id="f4e90-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![导入代码段选项](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="f4e90-134">支持的客户端</span><span class="sxs-lookup"><span data-stu-id="f4e90-134">Supported clients</span></span>

<span data-ttu-id="f4e90-135">以下客户端上的 Excel、Word 和 PowerPoint 支持脚本实验室。</span><span class="sxs-lookup"><span data-stu-id="f4e90-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="f4e90-136">Windows 上的 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="f4e90-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="f4e90-137">Mac 上的 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="f4e90-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="f4e90-138">网上的 Office</span><span class="sxs-lookup"><span data-stu-id="f4e90-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="f4e90-139">后续步骤</span><span class="sxs-lookup"><span data-stu-id="f4e90-139">Next steps</span></span>

<span data-ttu-id="f4e90-140">若要在 Excel、Word 或 PowerPoint 中使用脚本实验室, 请从 AppSource 安装[脚本实验室加载项](https://appsource.microsoft.com/product/office/WA104380862)。</span><span class="sxs-lookup"><span data-stu-id="f4e90-140">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="f4e90-141">欢迎您通过将新代码片段发布到[office js](https://github.com/OfficeDev/office-js-snippets#office-js-snippets)的 GitHub 存储库来扩展脚本实验室中的示例库。</span><span class="sxs-lookup"><span data-stu-id="f4e90-141">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="f4e90-142">当您准备好创建第一个 Office 加载项时, 请试用[Excel](../quickstarts/excel-quickstart-jquery.md)、 [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context)、 [Word](../quickstarts/word-quickstart.md)、 [OneNote](../quickstarts/onenote-quickstart.md)、 [PowerPoint](../quickstarts/powerpoint-quickstart.md)或[Project](../quickstarts/project-quickstart.md)的快速入门。</span><span class="sxs-lookup"><span data-stu-id="f4e90-142">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f4e90-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f4e90-143">See also</span></span>

- [<span data-ttu-id="f4e90-144">获取脚本实验室</span><span class="sxs-lookup"><span data-stu-id="f4e90-144">Get Script Lab</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="f4e90-145">了解有关脚本实验室的详细信息</span><span class="sxs-lookup"><span data-stu-id="f4e90-145">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="f4e90-146">注册开发计划</span><span class="sxs-lookup"><span data-stu-id="f4e90-146">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)
