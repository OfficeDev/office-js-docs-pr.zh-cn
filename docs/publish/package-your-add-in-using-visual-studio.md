---
title: 使用 Visual Studio 发布加载项
description: 如何使用 Visual Studio 2019 部署 Web 项目并打包加载项。
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 78b80e0c6a595f83f3a8cdde1db806a7612036ea
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950717"
---
# <a name="publish-your-add-in-using-visual-studio"></a><span data-ttu-id="88e0a-103">使用 Visual Studio 发布加载项</span><span class="sxs-lookup"><span data-stu-id="88e0a-103">Publish your add-in using Visual Studio</span></span>

<span data-ttu-id="88e0a-104">Office 加载项包包含 XML [清单文件](../develop/add-in-manifests.md)，它可用于发布加载项。</span><span class="sxs-lookup"><span data-stu-id="88e0a-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="88e0a-105">你将不得不单独发布项目的 Web 应用程序文件。</span><span class="sxs-lookup"><span data-stu-id="88e0a-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="88e0a-106">本文介绍如何使用 Visual Studio 2019 部署 Web 项目并打包加载项。</span><span class="sxs-lookup"><span data-stu-id="88e0a-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2019.</span></span>

> [!NOTE]
> <span data-ttu-id="88e0a-107">要了解如何发布使用 Yeoman 生成器创建并使用 Visual Studio Code 或任何其他编辑器开发的 Office 加载项，请参阅[发布使用 Visual Studio Code 开发的加载项](publish-add-in-vs-code.md)。</span><span class="sxs-lookup"><span data-stu-id="88e0a-107">For information about publishing an Office Add-in that you created using the Yeoman generator and developed with Visual Studio Code or any other editor, see [Publish an add-in developed with Visual Studio Code](publish-add-in-vs-code.md).</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a><span data-ttu-id="88e0a-108">使用 Visual Studio 2019 部署 Web 项目</span><span class="sxs-lookup"><span data-stu-id="88e0a-108">To deploy your web project using Visual Studio 2019</span></span>

<span data-ttu-id="88e0a-109">完成以下步骤以使用 Visual Studio 2019 部署 Web 项目。</span><span class="sxs-lookup"><span data-stu-id="88e0a-109">Complete the following steps to deploy your web project using Visual Studio 2019.</span></span>

1. <span data-ttu-id="88e0a-110">从“**生成**”选项卡中，选择“**发布 [加载项名称]**”。</span><span class="sxs-lookup"><span data-stu-id="88e0a-110">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>

2. <span data-ttu-id="88e0a-111">在“**选取发布目标**”窗口中，选择其中一个选项以发布到你的首选目标。</span><span class="sxs-lookup"><span data-stu-id="88e0a-111">In the **Pick a publish target** window, choose one of the options to publish to your preferred target.</span></span> <span data-ttu-id="88e0a-112">每个发布目标都要求你提供有关入门的详细信息，例如 Azure 虚拟机或文件夹位置。</span><span class="sxs-lookup"><span data-stu-id="88e0a-112">Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location.</span></span> <span data-ttu-id="88e0a-113">指定发布位置并填写所有必需信息后，选择“**发布**”</span><span class="sxs-lookup"><span data-stu-id="88e0a-113">Once you have specified a publish location and filled in all of the information required, select **Publish**</span></span>

    > [!NOTE]
    > <span data-ttu-id="88e0a-114">选取发布目标将会指定你要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。</span><span class="sxs-lookup"><span data-stu-id="88e0a-114">Picking a publish target specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

3. <span data-ttu-id="88e0a-115">有关每个发布目标选项的部署步骤的详细信息，请参阅[初探 Visual Studio 中的部署](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019)。</span><span class="sxs-lookup"><span data-stu-id="88e0a-115">For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span></span>

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a><span data-ttu-id="88e0a-116">使用 Visual Studio 2019 通过 IIS、FTP 或 Web 部署方法打包并发布加载项</span><span class="sxs-lookup"><span data-stu-id="88e0a-116">To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019</span></span>

<span data-ttu-id="88e0a-117">完成以下步骤以使用 Visual Studio 2019 打包加载项。</span><span class="sxs-lookup"><span data-stu-id="88e0a-117">Complete the following steps to package your add-in using Visual Studio 2019.</span></span>

1. <span data-ttu-id="88e0a-118">从“**生成**”选项卡中，选择“**发布 [加载项名称]**”。</span><span class="sxs-lookup"><span data-stu-id="88e0a-118">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>
2. <span data-ttu-id="88e0a-119">在“**选取发布目标**”窗口中，选择“**IIS、FTP 等**”，然后选择“**配置**”。</span><span class="sxs-lookup"><span data-stu-id="88e0a-119">In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**.</span></span> <span data-ttu-id="88e0a-120">接下来，选择“**发布**”。</span><span class="sxs-lookup"><span data-stu-id="88e0a-120">Next, select **Publish**.</span></span>
3. <span data-ttu-id="88e0a-121">此时将显示一个向导，它将指导你完成该过程。</span><span class="sxs-lookup"><span data-stu-id="88e0a-121">A wizard appears that will help guide you through the process.</span></span> <span data-ttu-id="88e0a-122">确保发布方法是你的首选方法，例如 Web 部署。</span><span class="sxs-lookup"><span data-stu-id="88e0a-122">Ensure the publish method is your preferred method, such as Web Deploy.</span></span>
4. <span data-ttu-id="88e0a-123">在“**目标 URL**”框中，输入托管加载项内容文件的网站的 URL，然后选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="88e0a-123">In the **Destination URL** box, enter the URL of the website that will host the content files of your add-in, and then select **Next**.</span></span> <span data-ttu-id="88e0a-124">如果计划将加载项提交到 AppSource，可以选择“**验证连接**”按钮，以发现任何可能会导致加载项遭拒的问题。</span><span class="sxs-lookup"><span data-stu-id="88e0a-124">If you plan to submit your add-in to AppSource, you can choose the **Validate Connection** button to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="88e0a-125">应先解决所有问题，再将加载项提交到 Microsoft Store。</span><span class="sxs-lookup"><span data-stu-id="88e0a-125">You should address all issues before you submit your add-in to the store.</span></span>
5. <span data-ttu-id="88e0a-126">确认所需的任何设置（包括“**文件发布选项**”），然后选择“**保存**”。</span><span class="sxs-lookup"><span data-stu-id="88e0a-126">Confirm any settings desired including **File Publish Options** and select **Save**.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="88e0a-127">Azure 网站自动提供 HTTPS 终结点。</span><span class="sxs-lookup"><span data-stu-id="88e0a-127">Azure websites automatically provide an HTTPS endpoint.</span></span>

<span data-ttu-id="88e0a-p106">现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。XML 清单位于 `app.publish` 文件夹的 `OfficeAppManifests` 中。例如：</span><span class="sxs-lookup"><span data-stu-id="88e0a-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a><span data-ttu-id="88e0a-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="88e0a-131">See also</span></span>

- [<span data-ttu-id="88e0a-132">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="88e0a-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="88e0a-133">将解决方案提交到 AppSource 和 Office 应用商店</span><span class="sxs-lookup"><span data-stu-id="88e0a-133">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
