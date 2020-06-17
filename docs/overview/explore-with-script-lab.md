---
title: 使用 Script Lab 探索 Office JavaScript API
description: 使用 Script Lab 探索 Office JS API 和原型功能。
ms.date: 04/16/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 88c57e163e8fc59e31fec80f5faa0bfbfd96402b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604550"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="06c03-103">使用 Script Lab 探索 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="06c03-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="06c03-104">可从 AppSource 免费获取 [Script Lab 加载项](https://appsource.microsoft.com/product/office/WA104380862)，使用 Excel 或 Word 等 Office 程序时可通过它探索 Office JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="06c03-104">The [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862), which is available free from AppSource, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="06c03-105">Script Lab 是一项方便的工具，可将其作为原型添加到开发工具包，并在加载项中验证你想使用的功能。</span><span class="sxs-lookup"><span data-stu-id="06c03-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="06c03-106">什么是 Script Lab？</span><span class="sxs-lookup"><span data-stu-id="06c03-106">What is Script Lab?</span></span>

<span data-ttu-id="06c03-107">任何人都可以使用 Script Lab 工具，了解如何在 Excel、Word 或 PowerPoint 中编写使用 Office JavaScript API 的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="06c03-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="06c03-108">它提供 IntelliSense，让你可以看到可用的内容；并且它是基于 Monaco 框架构建的（Visual Studio Code 也使用该框架）。</span><span class="sxs-lookup"><span data-stu-id="06c03-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="06c03-109">通过 Script Lab，可访问示例库以快速试用各项功能，也由示例开始编写自己的代码。</span><span class="sxs-lookup"><span data-stu-id="06c03-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="06c03-110">甚至可以通过 Script Lab 试用预览 API。</span><span class="sxs-lookup"><span data-stu-id="06c03-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="06c03-111">听起来还不错吧？</span><span class="sxs-lookup"><span data-stu-id="06c03-111">Sounds good so far?</span></span> <span data-ttu-id="06c03-112">观看以下片长一分钟的视频，在操作中了解 Script Lab。</span><span class="sxs-lookup"><span data-stu-id="06c03-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="06c03-113">[![展示 Script Lab 在 Excel、Word 和 PowerPoint 中运行的预览视频。](../images/screenshot-wide-youtube.png 'Script Lab 预览视频')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="06c03-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="06c03-114">关键功能</span><span class="sxs-lookup"><span data-stu-id="06c03-114">Key features</span></span>

<span data-ttu-id="06c03-115">Script Lab 提供许多功能，可帮助你探索 Office JavaScript API 和原型加载项功能。</span><span class="sxs-lookup"><span data-stu-id="06c03-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="06c03-116">浏览示例</span><span class="sxs-lookup"><span data-stu-id="06c03-116">Explore samples</span></span>

<span data-ttu-id="06c03-117">通过一系列展示如何使用 API 完成任务的内置示例快速入门。</span><span class="sxs-lookup"><span data-stu-id="06c03-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="06c03-118">可以运行示例来立即查看任务窗格或文档中的结果，检查示例来了解 API 的工作原理，甚至可以使用示例来构建自己的加载项的原型。</span><span class="sxs-lookup"><span data-stu-id="06c03-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![示例](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="06c03-120">代码和样式</span><span class="sxs-lookup"><span data-stu-id="06c03-120">Code and style</span></span>

<span data-ttu-id="06c03-121">除了用于调用 Office JS API 的 JavaScript 或 TypeScript 代码之外，每个代码段还包含用于定义任务窗格内容的 HTML 标记和用于定义任务窗格外观的 CSS。</span><span class="sxs-lookup"><span data-stu-id="06c03-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="06c03-122">在为自己的加载项确定任务窗格设计原型时，可以自定义该 HTML 标记 和 CSS，对元素放置和样式设计进行试验。</span><span class="sxs-lookup"><span data-stu-id="06c03-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="06c03-123">若要在代码段中调用预览 API，需更新该代码段的库，令其使用 beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) 和预览类型定义 `@types/office-js-preview`。</span><span class="sxs-lookup"><span data-stu-id="06c03-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="06c03-124">此外，仅当注册 [Office 预览体验计划](https://insider.office.com)后、运行 Office 预览体验计划版本时，才能访问某些预览 API。</span><span class="sxs-lookup"><span data-stu-id="06c03-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://insider.office.com) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="06c03-125">保存和共享代码段</span><span class="sxs-lookup"><span data-stu-id="06c03-125">Save and share snippets</span></span>

<span data-ttu-id="06c03-126">默认情况下，在 Script Lab 中打开的代码段将保存到浏览器缓存中。</span><span class="sxs-lookup"><span data-stu-id="06c03-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="06c03-127">若要永久保存代码段，可将其导出到 [GitHub gist](https://gist.github.com)。</span><span class="sxs-lookup"><span data-stu-id="06c03-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="06c03-128">可创建机密 gist 来保存自己专用的代码段，或创建公用 gist 以便与他人共享。</span><span class="sxs-lookup"><span data-stu-id="06c03-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![共享选项](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="06c03-130">导入代码段</span><span class="sxs-lookup"><span data-stu-id="06c03-130">Import snippets</span></span>

<span data-ttu-id="06c03-131">可通过指定存用于储代码段 YAML 的公共 [GitHub gist](https://gist.github.com) URL，或通过在代码段的完整 YAML 中粘贴，将代码段导入到 Script Lab。</span><span class="sxs-lookup"><span data-stu-id="06c03-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="06c03-132">当其他人通过发布到 GitHub gist 或提供 YAML 来与你共享其代码段时，此功能可能很有用。</span><span class="sxs-lookup"><span data-stu-id="06c03-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![导入代码段选项](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="06c03-134">支持的客户端</span><span class="sxs-lookup"><span data-stu-id="06c03-134">Supported clients</span></span>

<span data-ttu-id="06c03-135">以下客户端上的 Excel、Word 和 PowerPoint 支持 Script Lab。</span><span class="sxs-lookup"><span data-stu-id="06c03-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="06c03-136">Windows 上的 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="06c03-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="06c03-137">Mac 上的 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="06c03-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="06c03-138">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="06c03-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="06c03-139">后续步骤</span><span class="sxs-lookup"><span data-stu-id="06c03-139">Next steps</span></span>

<span data-ttu-id="06c03-140">若要在 Excel、Word 或 PowerPoint 中使用 Script Lab，请从 AppSource 安装 [Script Lab 加载项](https://appsource.microsoft.com/product/office/WA104380862)。</span><span class="sxs-lookup"><span data-stu-id="06c03-140">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="06c03-141">欢迎将新代码段发布到 [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub 存储库，以扩充 Script Lab 中的示例库。</span><span class="sxs-lookup"><span data-stu-id="06c03-141">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="06c03-142">准备好创建你的首个 Office 加载项时，请尝试使用 [Excel](../quickstarts/excel-quickstart-jquery.md)、[Outlook](../quickstarts/outlook-quickstart.md)、[Word](../quickstarts/word-quickstart.md)、[OneNote](../quickstarts/onenote-quickstart.md)、[PowerPoint](../quickstarts/powerpoint-quickstart.md) 或 [Project](../quickstarts/project-quickstart.md) 快速入门。</span><span class="sxs-lookup"><span data-stu-id="06c03-142">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="06c03-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="06c03-143">See also</span></span>

- [<span data-ttu-id="06c03-144">获取 Script Lab</span><span class="sxs-lookup"><span data-stu-id="06c03-144">Get Script Lab</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="06c03-145">详细了解 Script Lab</span><span class="sxs-lookup"><span data-stu-id="06c03-145">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="06c03-146">加入 Office 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="06c03-146">Join the Office 365 Developer Program</span></span>](https://developer.microsoft.com/office/dev-program)
- [<span data-ttu-id="06c03-147">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="06c03-147">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
