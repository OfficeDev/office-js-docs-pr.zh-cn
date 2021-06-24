---
title: 使用 Script Lab 探索 Office JavaScript API
description: 使用 Script Lab 探索 Office JS API 和原型功能。
ms.date: 06/18/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 7f4b67dd2369181e5d7b2b92496c8259ffd5c120
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077006"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="498b3-103">使用 Script Lab 探索 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="498b3-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="498b3-104">借助 [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) 和[适用于 Outlook 的 Script Lab](https://appsource.microsoft.com/product/office/wa200001603) 加载项（可从 AppSource 免费获取），你可以在使用 Excel 或 Outlook 等 Office 程序时探索 Office JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="498b3-104">The [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) and [Script Lab for Outlook](https://appsource.microsoft.com/product/office/wa200001603) add-ins, available free from AppSource, enable you to explore the Office JavaScript API while you're working in an Office program such as Excel or Outlook.</span></span> <span data-ttu-id="498b3-105">Script Lab 是一项方便的工具，可将其作为原型添加到开发工具包，并在你自己的加载项中验证你想使用的功能。</span><span class="sxs-lookup"><span data-stu-id="498b3-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your own add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="498b3-106">什么是 Script Lab？</span><span class="sxs-lookup"><span data-stu-id="498b3-106">What is Script Lab?</span></span>

<span data-ttu-id="498b3-107">Script Lab 是一款面向具有以下需求的用户的工具：希望了解如何在 Excel、Outlook、Word 和 PowerPoint 中开发使用 Office JavaScript API 的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="498b3-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Outlook, Word, and PowerPoint.</span></span> <span data-ttu-id="498b3-108">它提供 IntelliSense，让你可以看到可用的内容；并且它是基于 Monaco 框架构建的（Visual Studio Code 也使用该框架）。</span><span class="sxs-lookup"><span data-stu-id="498b3-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="498b3-109">通过 Script Lab，可访问示例库以快速试用各项功能，也由示例开始编写自己的代码。</span><span class="sxs-lookup"><span data-stu-id="498b3-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="498b3-110">甚至可以通过 Script Lab 试用预览 API。</span><span class="sxs-lookup"><span data-stu-id="498b3-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="498b3-p103">到目前为止听起来不错？观看以下片长一分钟的视频，在操作中了解 Script Lab。</span><span class="sxs-lookup"><span data-stu-id="498b3-p103">Sounds good so far? Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="498b3-113">[![展示 Script Lab 在 Excel、Word 和 PowerPoint 中运行的预览视频。](../images/screenshot-wide-youtube.png 'Script Lab 预览视频。')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="498b3-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video.')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="498b3-114">关键功能</span><span class="sxs-lookup"><span data-stu-id="498b3-114">Key features</span></span>

<span data-ttu-id="498b3-115">Script Lab 提供许多功能，可帮助你探索 Office JavaScript API 和原型加载项功能。</span><span class="sxs-lookup"><span data-stu-id="498b3-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="498b3-116">浏览示例</span><span class="sxs-lookup"><span data-stu-id="498b3-116">Explore samples</span></span>

<span data-ttu-id="498b3-117">通过一系列展示如何使用 API 完成任务的内置示例快速入门。</span><span class="sxs-lookup"><span data-stu-id="498b3-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="498b3-118">可以运行示例来立即查看任务窗格或文档中的结果，检查示例来了解 API 的工作原理，甚至可以使用示例来构建自己的加载项的原型。</span><span class="sxs-lookup"><span data-stu-id="498b3-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![示例。](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="498b3-120">代码和样式</span><span class="sxs-lookup"><span data-stu-id="498b3-120">Code and style</span></span>

<span data-ttu-id="498b3-121">除了用于调用 Office JS API 的 JavaScript 或 TypeScript 代码之外，每个代码段还包含用于定义任务窗格内容的 HTML 标记和用于定义任务窗格外观的 CSS。</span><span class="sxs-lookup"><span data-stu-id="498b3-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="498b3-122">在为自己的加载项确定任务窗格设计原型时，可以自定义该 HTML 标记 和 CSS，对元素放置和样式设计进行试验。</span><span class="sxs-lookup"><span data-stu-id="498b3-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="498b3-123">若要在代码段中调用预览 API，需更新该代码段的库，令其使用 beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) 和预览类型定义 `@types/office-js-preview`。</span><span class="sxs-lookup"><span data-stu-id="498b3-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="498b3-124">此外，仅当注册 [Office 预览体验计划](https://insider.office.com)后、运行 Office 预览体验计划版本时，才能访问某些预览 API。</span><span class="sxs-lookup"><span data-stu-id="498b3-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://insider.office.com) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="498b3-125">保存和共享代码段</span><span class="sxs-lookup"><span data-stu-id="498b3-125">Save and share snippets</span></span>

<span data-ttu-id="498b3-126">默认情况下，在 Script Lab 中打开的代码段将保存到浏览器缓存中。</span><span class="sxs-lookup"><span data-stu-id="498b3-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="498b3-127">若要永久保存代码段，可将其导出到 [GitHub gist](https://gist.github.com)。</span><span class="sxs-lookup"><span data-stu-id="498b3-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="498b3-128">可创建机密 gist 来保存自己专用的代码段，或创建公用 gist 以便与他人共享。</span><span class="sxs-lookup"><span data-stu-id="498b3-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![共享选项。](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="498b3-130">导入代码段</span><span class="sxs-lookup"><span data-stu-id="498b3-130">Import snippets</span></span>

<span data-ttu-id="498b3-131">可通过指定存用于储代码段 YAML 的公共 [GitHub gist](https://gist.github.com) URL，或通过在代码段的完整 YAML 中粘贴，将代码段导入到 Script Lab。</span><span class="sxs-lookup"><span data-stu-id="498b3-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="498b3-132">当其他人通过发布到 GitHub gist 或提供 YAML 来与你共享其代码段时，此功能可能很有用。</span><span class="sxs-lookup"><span data-stu-id="498b3-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![导入代码段选项。](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="498b3-134">支持的客户端</span><span class="sxs-lookup"><span data-stu-id="498b3-134">Supported clients</span></span>

<span data-ttu-id="498b3-135">以下客户端上的 Excel、Word 和 PowerPoint 支持 Script Lab。</span><span class="sxs-lookup"><span data-stu-id="498b3-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="498b3-136">Microsoft 365 Office 订阅</span><span class="sxs-lookup"><span data-stu-id="498b3-136">Microsoft 365 Office subscription</span></span>
- <span data-ttu-id="498b3-137">Mac 上的 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="498b3-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="498b3-138">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="498b3-138">Office on the web</span></span>

<span data-ttu-id="498b3-139">适用于 Outlook 的 Script Lab 在以下客户端上可用。</span><span class="sxs-lookup"><span data-stu-id="498b3-139">Script Lab for Outlook is available on the following clients.</span></span>

- <span data-ttu-id="498b3-140">Microsoft 365 Office 订阅</span><span class="sxs-lookup"><span data-stu-id="498b3-140">Microsoft 365 Office subscription</span></span>
- <span data-ttu-id="498b3-141">Mac 版 Outlook 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="498b3-141">Outlook 2016 or later on Mac</span></span>
- <span data-ttu-id="498b3-142">使用 Chrome、Microsoft Edge 或 Safari 浏览器时的 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="498b3-142">Outlook on the web when using Chrome, Microsoft Edge, or Safari browsers</span></span>

<span data-ttu-id="498b3-143">有关适用于 Outlook 的 Script Lab 的更多详细信息，请参阅相关[博客文章](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/)。</span><span class="sxs-lookup"><span data-stu-id="498b3-143">For more details on Script Lab for Outlook, see the related [blog post](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="498b3-144">在 2021 年某个时间，Script Lab 将停止处理使用 Internet Explorer 托管加载项的平台和 Office 版本组合。这包括通过 Office 2019 一次性购买的 Office 版本，以及一些旧版本的 Microsoft 365（订阅）Office。</span><span class="sxs-lookup"><span data-stu-id="498b3-144">Sometime in 2021, Script Lab will stop working on the combinations of platform and Office version that use Internet Explorer to host add-ins. This includes one-time purchase Office versions through Office 2019, and some older versions of Microsoft 365 (subscription) Office.</span></span> <span data-ttu-id="498b3-145">（有关详细信息，请参阅[ Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。）需要其他平台和版本组合来浏览和测试使用 Script Lab 的 Office JavaScript 库 API。</span><span class="sxs-lookup"><span data-stu-id="498b3-145">(For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) You'll need other platform and version combinations to explore and test the Office JavaScript Library APIs with Script Lab.</span></span> <span data-ttu-id="498b3-146">但这些 API 的行为在 Internet Explorer 中并无不同，因此这不是 Script Lab 的一个弱点。</span><span class="sxs-lookup"><span data-stu-id="498b3-146">But the behavior of these APIs isn't different in Internet Explorer, so this isn't really a weakness of Script Lab.</span></span> <span data-ttu-id="498b3-147">请注意，提交到 [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 的 Office 加载项必须支持使用 Internet Explorer 托管加载项的平台和版本组合。</span><span class="sxs-lookup"><span data-stu-id="498b3-147">Note that Office Add-ins submitted to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) must support the platform and version combinations that use Internet Explorer to host add-ins.</span></span>

## <a name="next-steps"></a><span data-ttu-id="498b3-148">后续步骤</span><span class="sxs-lookup"><span data-stu-id="498b3-148">Next steps</span></span>

<span data-ttu-id="498b3-149">若要在 Excel、Word 或 PowerPoint 中使用 Script Lab，请从 AppSource 安装 [Script Lab 加载项](https://appsource.microsoft.com/product/office/WA104380862)。</span><span class="sxs-lookup"><span data-stu-id="498b3-149">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="498b3-150">若要使用适用于 Outlook 的 Script Lab，请从 AppSource 安装 [适用于 Outlook 的 Script Lab 加载项](https://appsource.microsoft.com/product/office/wa200001603)。</span><span class="sxs-lookup"><span data-stu-id="498b3-150">To use Script Lab for Outlook, install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) from AppSource.</span></span>

<span data-ttu-id="498b3-151">欢迎将新代码段发布到 [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub 存储库，以扩充 Script Lab 中的示例库。</span><span class="sxs-lookup"><span data-stu-id="498b3-151">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="498b3-152">准备好创建你的首个 Office 加载项时，请尝试使用 [Excel](../quickstarts/excel-quickstart-jquery.md)、[Outlook](../quickstarts/outlook-quickstart.md)、[Word](../quickstarts/word-quickstart.md)、[OneNote](../quickstarts/onenote-quickstart.md)、[PowerPoint](../quickstarts/powerpoint-quickstart.md) 或 [Project](../quickstarts/project-quickstart.md) 快速入门。</span><span class="sxs-lookup"><span data-stu-id="498b3-152">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="498b3-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="498b3-153">See also</span></span>

- [<span data-ttu-id="498b3-154">获取适用于 Excel、Word 或 Powerpoint 的 Script Lab</span><span class="sxs-lookup"><span data-stu-id="498b3-154">Get Script Lab for Excel, Word, or Powerpoint</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="498b3-155">获取适用于 Outlook 的 Script Lab</span><span class="sxs-lookup"><span data-stu-id="498b3-155">Get Script Lab for Outlook</span></span>](https://appsource.microsoft.com/product/office/wa200001603)
- [<span data-ttu-id="498b3-156">详细了解 Script Lab</span><span class="sxs-lookup"><span data-stu-id="498b3-156">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="498b3-157">加入 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="498b3-157">Join the Microsoft 365 developer program</span></span>](https://developer.microsoft.com/office/dev-program)
- [<span data-ttu-id="498b3-158">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="498b3-158">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="498b3-159">了解 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="498b3-159">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
