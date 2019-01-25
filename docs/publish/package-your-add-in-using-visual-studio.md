---
title: 使用 Visual Studio 打包加载项以准备发布 | Microsoft Docs
description: 如何使用 Visual Studio 2017 部署 Web 项目并打包加载项。
ms.date: 01/25/2018
localization_priority: Priority
ms.openlocfilehash: a135e8e72703c3de60290a9eb7b2e03c63449124
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386433"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="19ca7-103">使用 Visual Studio 打包加载项以准备发布</span><span class="sxs-lookup"><span data-stu-id="19ca7-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="19ca7-104">Office 加载项包包含 XML [清单文件](../develop/add-in-manifests.md)，它可用于发布加载项。</span><span class="sxs-lookup"><span data-stu-id="19ca7-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="19ca7-105">你将不得不单独发布项目的 Web 应用程序文件。</span><span class="sxs-lookup"><span data-stu-id="19ca7-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="19ca7-106">本文介绍如何使用 Visual Studio 2017 部署 Web 项目并打包加载项。</span><span class="sxs-lookup"><span data-stu-id="19ca7-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2017.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a><span data-ttu-id="19ca7-107">使用 Visual Studio 2017 部署 Web 项目</span><span class="sxs-lookup"><span data-stu-id="19ca7-107">To deploy your web project using Visual Studio 2017</span></span>

<span data-ttu-id="19ca7-108">完成以下步骤以使用 Visual Studio 2017 部署 Web 项目。</span><span class="sxs-lookup"><span data-stu-id="19ca7-108">Complete the following steps to deploy your web project using Visual Studio 2017.</span></span>

1. <span data-ttu-id="19ca7-109">在“**解决方案资源管理器**”中，打开外接程序项目的快捷菜单，然后选择“**发布**”。</span><span class="sxs-lookup"><span data-stu-id="19ca7-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="19ca7-110">将显示“**发布外接程序**”页。</span><span class="sxs-lookup"><span data-stu-id="19ca7-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="19ca7-111">选择“当前配置文件”\*\*\*\* 下拉列表中的配置文件，或选择“新建…”\*\*\*\* 新建配置文件。</span><span class="sxs-lookup"><span data-stu-id="19ca7-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="19ca7-112">发布配置文件指定你要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。</span><span class="sxs-lookup"><span data-stu-id="19ca7-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="19ca7-113">如果你选择“**新建...**”，则向导将会显示“**创建发布配置文件**”页。</span><span class="sxs-lookup"><span data-stu-id="19ca7-113">If you choose  **New ...**, a wizard appears with the **Create publishing profile** page.</span></span> <span data-ttu-id="19ca7-114">可以使用此向导从托管提供程序（如 Microsoft Azure）的网站导入发布配置文件，或创建新配置文件并添加你的服务器、凭据以及下一过程中的其他设置。</span><span class="sxs-lookup"><span data-stu-id="19ca7-114">You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="19ca7-115">有关导入发布配置文件或创建新发布配置文件的详细信息，请参阅[创建发布配置文件](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)。</span><span class="sxs-lookup"><span data-stu-id="19ca7-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="19ca7-116">在“**发布加载项**”页中，选择“**部署 Web 项目**”链接。</span><span class="sxs-lookup"><span data-stu-id="19ca7-116">On the **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="19ca7-117">将显示“**发布**”对话框。</span><span class="sxs-lookup"><span data-stu-id="19ca7-117">The  **Publish** dialog box appears.</span></span> <span data-ttu-id="19ca7-118">有关如何使用此向导的详细信息，请参阅[操作方法：使用 Visual Studio 中的一键式发布来部署 Web 项目](https://msdn.microsoft.com/library/dd465337.aspx)。</span><span class="sxs-lookup"><span data-stu-id="19ca7-118">For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a><span data-ttu-id="19ca7-119">使用 Visual Studio 2017 打包加载项</span><span class="sxs-lookup"><span data-stu-id="19ca7-119">To package your add-in using Visual Studio 2017</span></span>

<span data-ttu-id="19ca7-120">完成以下步骤以使用 Visual Studio 2017 打包加载项。</span><span class="sxs-lookup"><span data-stu-id="19ca7-120">Complete the following steps to package your add-in using Visual Studio 2017.</span></span>

1. <span data-ttu-id="19ca7-121">在“**发布加载项**”页上，选择“**打包加载项**”按钮。</span><span class="sxs-lookup"><span data-stu-id="19ca7-121">In the **Publish your add-in** page, choose the **Package the add-in** button.</span></span>
    
    <span data-ttu-id="19ca7-122">此时向导将显示“**打包加载项**”页面。</span><span class="sxs-lookup"><span data-stu-id="19ca7-122">A wizard appears with the **Package the add-in** page.</span></span>
    
2. <span data-ttu-id="19ca7-123">在“你的网站托管在哪里?”\*\*\*\* 下拉列表中，选择或输入托管加载项内容文件的网站的 URL，然后选择“完成”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="19ca7-123">In the **Where is your website hosted?** box, enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span>
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="19ca7-124">Azure 网站自动提供 HTTPS 终结点。</span><span class="sxs-lookup"><span data-stu-id="19ca7-124">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="19ca7-125">此时，Visual Studio 生成发布加载项所需的文件，并打开发布输出文件夹。</span><span class="sxs-lookup"><span data-stu-id="19ca7-125">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span>
    
<span data-ttu-id="19ca7-126">如果计划将加载项提交到 AppSource，可以选择“**执行验证检查**”按钮，以发现任何可能会导致加载项遭拒的问题。</span><span class="sxs-lookup"><span data-stu-id="19ca7-126">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="19ca7-127">应先解决所有问题，再将加载项提交到 Microsoft Store。</span><span class="sxs-lookup"><span data-stu-id="19ca7-127">You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="19ca7-p105">现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。XML 清单位于 `app.publish` 文件夹的 `OfficeAppManifests` 中。例如：</span><span class="sxs-lookup"><span data-stu-id="19ca7-p105">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="19ca7-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="19ca7-131">See also</span></span>

- [<span data-ttu-id="19ca7-132">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="19ca7-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="19ca7-133">将解决方案提交到 AppSource 和 Office 应用商店</span><span class="sxs-lookup"><span data-stu-id="19ca7-133">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
