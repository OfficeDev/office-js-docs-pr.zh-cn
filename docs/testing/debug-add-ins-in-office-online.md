---
title: 在 Office Online 中调试加载项
description: 如何使用 Office Online 测试和调试加载项。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: ff77f3d8b3e332288d4ccb3e2d2305d1b1c4a825
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451526"
---
# <a name="debug-add-ins-in-office-online"></a><span data-ttu-id="84a85-103">在 Office Online 中调试加载项</span><span class="sxs-lookup"><span data-stu-id="84a85-103">Debug add-ins in Office Online</span></span>


<span data-ttu-id="84a85-104">您可以在并非运行 Windows 或 Office 2013 或 Office 2016 桌面客户端的计算机上构建和调试外接程序，例如，如果您正在使用 Mac 进行开发。本文介绍如何使用 Office Online 测试和调试您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="84a85-104">You can build and debug add-ins on a computer that isn't running Windows or the Office desktop client&mdash;for example, if you're developing on a Mac.</span></span> <span data-ttu-id="84a85-105">本文介绍如何使用 Office Online 测试和调试加载项。</span><span class="sxs-lookup"><span data-stu-id="84a85-105">This article describes how to use Office Online to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="84a85-106">先决条件</span><span class="sxs-lookup"><span data-stu-id="84a85-106">Prerequisites</span></span>

<span data-ttu-id="84a85-107">首先，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="84a85-107">To get started:</span></span>

- <span data-ttu-id="84a85-108">获取 Office 365 开发人员帐户（如果还没有的话），或获取对 SharePoint 网站的访问权限。</span><span class="sxs-lookup"><span data-stu-id="84a85-108">Get an Office 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>
    
  > [!NOTE]
  > <span data-ttu-id="84a85-p102">若要注册免费 Office 365 开发人员订阅，请加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)。 请参阅 [Office 365 开发人员计划文档](/office/developer-program/office-365-developer-program)，逐步了解如何加入 Office 365 开发人员计划并注册和配置订阅。</span><span class="sxs-lookup"><span data-stu-id="84a85-p102">To sign up for a free Office 365 developer subscription, join our [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program). See the [Office 365 Developer Program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and sign up and configure your subscription.</span></span>
     
- <span data-ttu-id="84a85-p103">对 Office 365 (SharePoint Online) 设置加载项目录。加载项目录是 SharePoint Online 中的专用网站集，用于托管 Office 加载项的文档库。如果有自己的 SharePoint 网站，可以设置加载项目录文档库。有关详细信息，请参阅[向 SharePoint 上的加载项目录发布任务窗格和内容加载项](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。</span><span class="sxs-lookup"><span data-stu-id="84a85-p103">Set up an add-in catalog on Office 365 (SharePoint Online). An add-in catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an add-in catalog document library. For more information, see [Publish task pane and content add-ins to an add-in catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a><span data-ttu-id="84a85-114">通过 Excel Online 或 Word Online 调试加载项</span><span class="sxs-lookup"><span data-stu-id="84a85-114">Debug your add-in from Excel Online or Word Online</span></span>

<span data-ttu-id="84a85-115">要使用 Office Online 调试您的外接程序，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="84a85-115">To debug your add-in by using Office Online:</span></span>

1. <span data-ttu-id="84a85-116">将加载项部署到支持 SSL 的服务器上。</span><span class="sxs-lookup"><span data-stu-id="84a85-116">Deploy your add-in to a server that supports SSL.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="84a85-117">建议使用 [Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建和托管加载项。</span><span class="sxs-lookup"><span data-stu-id="84a85-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>
     
2. <span data-ttu-id="84a85-p104">在[加载项清单文件](../develop/add-in-manifests.md)中，将 **SourceLocation** 元素值更新为包括绝对 URI，而不是相对 URI。例如：</span><span class="sxs-lookup"><span data-stu-id="84a85-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. <span data-ttu-id="84a85-120">将清单上传到 SharePoint 上加载项目录中的“Office 加载项”库。</span><span class="sxs-lookup"><span data-stu-id="84a85-120">Upload the manifest to the Office Add-ins library in the add-in catalog on SharePoint.</span></span>
    
4. <span data-ttu-id="84a85-121">从 Office 365 中的应用程序启动程序启动 Excel Online 或 Word Online，并打开一个新文档。</span><span class="sxs-lookup"><span data-stu-id="84a85-121">Launch Excel Online or Word Online from the app launcher in Office 365, and open a new document.</span></span>
    
5. <span data-ttu-id="84a85-122">在“插入”选项卡上，选择“**我的外接程序**”或“**Office 外接程序**”以插入你的外接程序并在应用中对其测试。</span><span class="sxs-lookup"><span data-stu-id="84a85-122">On the Insert tab, choose  **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>
    
6. <span data-ttu-id="84a85-123">使用常用浏览器工具调试器调试加载项。</span><span class="sxs-lookup"><span data-stu-id="84a85-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="84a85-124">潜在问题</span><span class="sxs-lookup"><span data-stu-id="84a85-124">Potential issues</span></span>    

<span data-ttu-id="84a85-125">下面介绍了一些在调试过程中可能会遇到的问题：</span><span class="sxs-lookup"><span data-stu-id="84a85-125">The following are some issues that you might encounter as you debug:</span></span>
    
- <span data-ttu-id="84a85-126">您看到的一些 JavaScript 错误可能源自 Office Online。</span><span class="sxs-lookup"><span data-stu-id="84a85-126">Some JavaScript errors that you see might originate from Office Online.</span></span>
      
- <span data-ttu-id="84a85-127">浏览器可能会显示无效证书错误，您需绕过此错误。</span><span class="sxs-lookup"><span data-stu-id="84a85-127">The browser might show an invalid certificate error that you will need to bypass.</span></span>
      
- <span data-ttu-id="84a85-128">如果在代码中设置了断点，Office Online 可能会抛出错误，指示无法保存。</span><span class="sxs-lookup"><span data-stu-id="84a85-128">If you set breakpoints in your code, Office Online might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="84a85-129">另请参阅</span><span class="sxs-lookup"><span data-stu-id="84a85-129">See also</span></span>

- [<span data-ttu-id="84a85-130">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="84a85-130">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="84a85-131">AppSource 验证策略</span><span class="sxs-lookup"><span data-stu-id="84a85-131">AppSource validation policies</span></span>](/office/dev/store/validation-policies)  
- [<span data-ttu-id="84a85-132">创建有效的 AppSource 应用和加载项</span><span class="sxs-lookup"><span data-stu-id="84a85-132">Create effective AppSource apps and add-ins</span></span>](/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="84a85-133">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="84a85-133">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
    
