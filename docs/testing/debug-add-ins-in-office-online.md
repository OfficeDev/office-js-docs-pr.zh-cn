---
title: 在 Office 网页版中调试加载项
description: 如何使用 Office 网页版来测试和调试加载项。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: c840d0a16e2a4cdf0bb9f4b213099cb74c2aa815
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719810"
---
# <a name="debug-add-ins-in-office-on-the-web"></a><span data-ttu-id="94272-103">在 Office 网页版中调试加载项</span><span class="sxs-lookup"><span data-stu-id="94272-103">Debug add-ins in Office on the web</span></span>


<span data-ttu-id="94272-104">您可以在并非运行 Windows 或 Office 2013 或 Office 2016 桌面客户端的计算机上构建和调试外接程序，例如，如果您正在使用 Mac 进行开发。本文介绍如何使用 Office Online 测试和调试您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="94272-104">You can build and debug add-ins on a computer that isn't running Windows or the Office desktop client&mdash;for example, if you're developing on a Mac.</span></span> <span data-ttu-id="94272-105">本文介绍了如何使用 Office 网页版来测试和调试加载项。</span><span class="sxs-lookup"><span data-stu-id="94272-105">This article describes how to use Office on the web to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="94272-106">先决条件</span><span class="sxs-lookup"><span data-stu-id="94272-106">Prerequisites</span></span>

<span data-ttu-id="94272-107">首先，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="94272-107">To get started:</span></span>

- <span data-ttu-id="94272-108">获取 Office 365 开发人员帐户（如果还没有的话），或获取对 SharePoint 网站的访问权限。</span><span class="sxs-lookup"><span data-stu-id="94272-108">Get an Office 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>

  > [!NOTE]
  > <span data-ttu-id="94272-p102">若要获取 90 天免费的可续订 Office 365 开发人员订阅，请加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)。请参阅 [Office 365 开发人员计划文档](/office/developer-program/office-365-developer-program)，逐步了解如何加入 Office 365 开发人员计划并配置订阅。</span><span class="sxs-lookup"><span data-stu-id="94272-p102">To get a free, 90-day renewable Office 365 developer subscription, join our [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program). See the [Office 365 Developer Program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and configure your subscription.</span></span>

- <span data-ttu-id="94272-p103">在 Office 365 (SharePoint Online) 上创建应用程序目录。应用程序目录是 SharePoint Online 中的专用网站集，用于托管 Office 加载项的文档库。如果你有自己的 SharePoint 网站，可以创建应用程序目录文档库。有关详细信息，请参阅[向 SharePoint 上的应用程序目录发布任务窗格加载项和内容加载项](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。</span><span class="sxs-lookup"><span data-stu-id="94272-p103">Set up an app catalog on Office 365 (SharePoint Online). An app catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an app catalog document library. For more information, see [Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a><span data-ttu-id="94272-114">在 Excel 网页版或 Word 网页版中调试加载项</span><span class="sxs-lookup"><span data-stu-id="94272-114">Debug your add-in from Excel or Word on the web</span></span>

<span data-ttu-id="94272-115">若要使用 Office 网页版调试加载项，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="94272-115">To debug your add-in by using Office on the web:</span></span>

1. <span data-ttu-id="94272-116">将加载项部署到支持 SSL 的服务器上。</span><span class="sxs-lookup"><span data-stu-id="94272-116">Deploy your add-in to a server that supports SSL.</span></span>

    > [!NOTE]
    > <span data-ttu-id="94272-117">建议使用 [Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建和托管加载项。</span><span class="sxs-lookup"><span data-stu-id="94272-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>

2. <span data-ttu-id="94272-p104">在[加载项清单文件](../develop/add-in-manifests.md)中，将 **SourceLocation** 元素值更新为包括绝对 URI，而不是相对 URI。例如：</span><span class="sxs-lookup"><span data-stu-id="94272-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. <span data-ttu-id="94272-120">将清单上传到 SharePoint 上应用程序目录中的 Office 加载项文档库。</span><span class="sxs-lookup"><span data-stu-id="94272-120">Upload the manifest to the Office Add-ins library in the app catalog on SharePoint.</span></span>

4. <span data-ttu-id="94272-121">使用 Office 365 中的应用程序启动器来启动 Excel 网页版或 Word 网页版，并打开新文档。</span><span class="sxs-lookup"><span data-stu-id="94272-121">Launch Excel or Word on the web from the app launcher in Office 365, and open a new document.</span></span>

5. <span data-ttu-id="94272-122">在“插入”选项卡上选择“我的外接程序”\*\*\*\* 或“Office 外接程序”\*\*\*\* 以插入您的外接程序并在应用程序中进行测试。</span><span class="sxs-lookup"><span data-stu-id="94272-122">On the Insert tab, choose **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>

6. <span data-ttu-id="94272-123">使用常用浏览器工具调试器调试加载项。</span><span class="sxs-lookup"><span data-stu-id="94272-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="94272-124">潜在问题</span><span class="sxs-lookup"><span data-stu-id="94272-124">Potential issues</span></span>

<span data-ttu-id="94272-125">下面介绍了一些在调试过程中可能会遇到的问题：</span><span class="sxs-lookup"><span data-stu-id="94272-125">The following are some issues that you might encounter as you debug:</span></span>

- <span data-ttu-id="94272-126">你看到的一些 JavaScript 错误可能源自 Office 网页版。</span><span class="sxs-lookup"><span data-stu-id="94272-126">Some JavaScript errors that you see might originate from Office on the web.</span></span>

- <span data-ttu-id="94272-127">浏览器可能会显示无效证书错误，你需要忽略此错误。</span><span class="sxs-lookup"><span data-stu-id="94272-127">The browser might show an invalid certificate error that you will need to bypass.</span></span> <span data-ttu-id="94272-128">执行此操作的过程因浏览器而异，而且用于执行此操作的各种浏览器的 UI 会定期进行更改。</span><span class="sxs-lookup"><span data-stu-id="94272-128">The process for doing this varies with the browser and the various browsers' UIs for doing this change periodically.</span></span> <span data-ttu-id="94272-129">有关说明，可搜索浏览器的“帮助”或“联机搜索”。</span><span class="sxs-lookup"><span data-stu-id="94272-129">You should search the browser's help or search online for instructions.</span></span> <span data-ttu-id="94272-130">（例如，搜索“Microsoft Edge 无效证书警告”。）大多数浏览器在“警告”页面上都有一个链接，可以通过此链接单击进入“加载项”页。</span><span class="sxs-lookup"><span data-stu-id="94272-130">(For example, search for "Microsoft Edge invalid certificate warning".) Most browsers will have a link on the warning page that enables you to click through to the add-in page.</span></span> <span data-ttu-id="94272-131">例如，Microsoft Edge 有一个链接“转到网页（不推荐）”。</span><span class="sxs-lookup"><span data-stu-id="94272-131">For example, Microsoft Edge has a link "Go on to the webpage (Not recommended)".</span></span> <span data-ttu-id="94272-132">但是每次加载项重新加载时，通常都必须通过此链接来完成。</span><span class="sxs-lookup"><span data-stu-id="94272-132">But you will usually have to go through this link every time the add-in reloads.</span></span> <span data-ttu-id="94272-133">如需更长久地忽略，请参阅建议的帮助。</span><span class="sxs-lookup"><span data-stu-id="94272-133">For a longer lasting bypass, see the help as suggested.</span></span>

- <span data-ttu-id="94272-134">如果你在代码中设置了断点，Office 网页版可能会抛出错误，指明它无法保存。</span><span class="sxs-lookup"><span data-stu-id="94272-134">If you set breakpoints in your code, Office on the web might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="94272-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="94272-135">See also</span></span>

- [<span data-ttu-id="94272-136">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="94272-136">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="94272-137">AppSource 验证策略</span><span class="sxs-lookup"><span data-stu-id="94272-137">AppSource validation policies</span></span>](/office/dev/store/validation-policies)  
- [<span data-ttu-id="94272-138">创建有效的 AppSource 应用和加载项</span><span class="sxs-lookup"><span data-stu-id="94272-138">Create effective AppSource apps and add-ins</span></span>](/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="94272-139">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="94272-139">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
    
