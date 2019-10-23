---
title: 使用 Visual Studio 打包加载项以准备发布
description: 如何使用 Visual Studio 2019 部署 Web 项目并打包加载项。
ms.date: 10/14/2019
localization_priority: Priority
ms.openlocfilehash: 784741cffa0e3015caaa9c70fbb56f4b70df9462
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626962"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="43193-103">使用 Visual Studio 打包加载项以准备发布</span><span class="sxs-lookup"><span data-stu-id="43193-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="43193-104">Office 加载项包包含 XML [清单文件](../develop/add-in-manifests.md)，它可用于发布加载项。</span><span class="sxs-lookup"><span data-stu-id="43193-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="43193-105">你将不得不单独发布项目的 Web 应用程序文件。</span><span class="sxs-lookup"><span data-stu-id="43193-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="43193-106">本文介绍如何使用 Visual Studio 2019 部署 Web 项目并打包加载项。</span><span class="sxs-lookup"><span data-stu-id="43193-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2017.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a><span data-ttu-id="43193-107">使用 Visual Studio 2019 部署 Web 项目</span><span class="sxs-lookup"><span data-stu-id="43193-107">To deploy your web project using Visual Studio 2017</span></span>

<span data-ttu-id="43193-108">完成以下步骤以使用 Visual Studio 2019 部署 Web 项目。</span><span class="sxs-lookup"><span data-stu-id="43193-108">Complete the following steps to deploy your web project using Visual Studio 2017.</span></span>

1. <span data-ttu-id="43193-109">从“**生成**”选项卡中，选择“**发布 [加载项名称]**”。</span><span class="sxs-lookup"><span data-stu-id="43193-109">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>

2. <span data-ttu-id="43193-110">在“**选取发布目标**”窗口中，选择其中一个选项以发布到你的首选目标。</span><span class="sxs-lookup"><span data-stu-id="43193-110">In the **Pick a publish target** window, choose one of the options to publish to your preferred target.</span></span> <span data-ttu-id="43193-111">每个发布目标都要求你提供有关入门的详细信息，例如 Azure 虚拟机或文件夹位置。</span><span class="sxs-lookup"><span data-stu-id="43193-111">Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location.</span></span> <span data-ttu-id="43193-112">指定发布位置并填写所有必需信息后，选择“**发布**”</span><span class="sxs-lookup"><span data-stu-id="43193-112">Once you have specified a publish location and filled in all of the information required, select **Publish**</span></span>

    > [!NOTE]
    > <span data-ttu-id="43193-113">选取发布目标将会指定你要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。</span><span class="sxs-lookup"><span data-stu-id="43193-113">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

3. <span data-ttu-id="43193-114">有关每个发布目标选项的部署步骤的详细信息，请参阅[初探 Visual Studio 中的部署](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019)。</span><span class="sxs-lookup"><span data-stu-id="43193-114">For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span></span>

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a><span data-ttu-id="43193-115">使用 Visual Studio 2019 通过 IIS、FTP 或 Web 部署方法打包并发布加载项</span><span class="sxs-lookup"><span data-stu-id="43193-115">To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019</span></span>

<span data-ttu-id="43193-116">完成以下步骤以使用 Visual Studio 2019 打包加载项。</span><span class="sxs-lookup"><span data-stu-id="43193-116">Complete the following steps to package your add-in using Visual Studio 2017.</span></span>

1. <span data-ttu-id="43193-117">从“**生成**”选项卡中，选择“**发布 [加载项名称]**”。</span><span class="sxs-lookup"><span data-stu-id="43193-117">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>
2. <span data-ttu-id="43193-118">在“**选取发布目标**”窗口中，选择“**IIS、FTP 等**”，然后选择“**配置**”。</span><span class="sxs-lookup"><span data-stu-id="43193-118">In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**.</span></span> <span data-ttu-id="43193-119">接下来，选择“**发布**”。</span><span class="sxs-lookup"><span data-stu-id="43193-119">Next, select **Publish**.</span></span>
3. <span data-ttu-id="43193-120">此时将显示一个向导，它将指导你完成该过程。</span><span class="sxs-lookup"><span data-stu-id="43193-120">A wizard appears that will help guide you through the process.</span></span> <span data-ttu-id="43193-121">确保发布方法是你的首选方法，例如 Web 部署。</span><span class="sxs-lookup"><span data-stu-id="43193-121">Ensure the publish method is your preferred method, such as Web Deploy.</span></span>
4. <span data-ttu-id="43193-122">在“**目标 URL**”框中，输入托管加载项内容文件的网站的 URL，然后选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="43193-122">In the **Where is your website hosted?** box, enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> <span data-ttu-id="43193-123">如果计划将加载项提交到 AppSource，可以选择“**验证连接**”按钮，以发现任何可能会导致加载项遭拒的问题。</span><span class="sxs-lookup"><span data-stu-id="43193-123">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** button to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="43193-124">应先解决所有问题，再将加载项提交到 Microsoft Store。</span><span class="sxs-lookup"><span data-stu-id="43193-124">You should address all issues before you submit your add-in to the store.</span></span>
5. <span data-ttu-id="43193-125">确认所需的任何设置（包括“**文件发布选项**”），然后选择“**保存**”。</span><span class="sxs-lookup"><span data-stu-id="43193-125">Confirm any settings desired including **File Publish Options** and select **Save**.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="43193-126">Azure 网站自动提供 HTTPS 终结点。</span><span class="sxs-lookup"><span data-stu-id="43193-126">Azure websites automatically provide an HTTPS endpoint.</span></span>

<span data-ttu-id="43193-p106">现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。XML 清单位于 `app.publish` 文件夹的 `OfficeAppManifests` 中。例如：</span><span class="sxs-lookup"><span data-stu-id="43193-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a><span data-ttu-id="43193-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="43193-130">See also</span></span>

- [<span data-ttu-id="43193-131">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="43193-131">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="43193-132">将解决方案提交到 AppSource 和 Office 应用商店</span><span class="sxs-lookup"><span data-stu-id="43193-132">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
