---
title: 在 Microsoft Azure 上托管 Office 加载项 | Microsoft Docs
description: 了 解如何将外接程序 Web 应用部署到 Azure 并旁加载外接程序以在 Office 客户端应用程序中进行测试。
ms.date: 01/25/2018
ms.openlocfilehash: 32560fdd0655fbba937140152f16cc91c2185411
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004999"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="8b023-103">在 Microsoft Azure 上托管 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="8b023-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="8b023-p101">最简单的 Office 外接程序由 XML 清单文件和 HTML 页构成。XML 清单文件描述了外接程序的特性，例如它的名称、可以运行它的 Office 客户端应用程序以及外接程序 HTML 页的 URL。HTML 页包含在一个 Web 应用中，用户在 Office 客户端应用程序中安装和运行外接程序时将与此 Web 应用进行交互。可以将 Office 外接程序的 Web 应用托管在任意 Web 托管平台（包括 Azure）上。</span><span class="sxs-lookup"><span data-stu-id="8b023-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="8b023-108">本文介绍了如何将外接程序 Web 应用部署到 Azure 并[旁加载外接程序](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)以在 Office 客户端应用程序中进行测试。</span><span class="sxs-lookup"><span data-stu-id="8b023-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="8b023-109">先决条件</span><span class="sxs-lookup"><span data-stu-id="8b023-109">Prerequisites</span></span> 

1. <span data-ttu-id="8b023-110">安装 [Visual Studio 2017](https://www.visualstudio.com/downloads)，并选择添加 **Azure 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="8b023-110">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8b023-111">如果之前已安装 Visual Studio 2017，请[使用 Visual Studio 安装程序](https://docs.microsoft.com/visualstudio/install/modify-visual-studio)，以确保安装 **Azure 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="8b023-111">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="8b023-112">安装 Office。</span><span class="sxs-lookup"><span data-stu-id="8b023-112">Install Office.</span></span> 
    
    > [!NOTE]
    > <span data-ttu-id="8b023-113">如果尚未安装 Office，可以[注册 1 个月免费试用版](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)。</span><span class="sxs-lookup"><span data-stu-id="8b023-113">If you don't already have Office 2016, you can [register for a free 1-month trial](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3.  <span data-ttu-id="8b023-114">获取 Azure 订阅。</span><span class="sxs-lookup"><span data-stu-id="8b023-114">Obtain an Azure subscription.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="8b023-115">如果还没有 Azure 订阅，可以[通过 Visual Studio 订阅获取](https://azure.microsoft.com/en-us/pricing/member-offers/visual-studio-subscriptions/)，也可以[注册免费试用版](https://azure.microsoft.com/pricing/free-trial)。</span><span class="sxs-lookup"><span data-stu-id="8b023-115">If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](https://azure.microsoft.com/en-us/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="8b023-116">第 1 步：创建用于托管加载项 XML 清单文件的共享文件夹</span><span class="sxs-lookup"><span data-stu-id="8b023-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="8b023-117">打开开发计算机的文件资源管理器。</span><span class="sxs-lookup"><span data-stu-id="8b023-117">Open File Explorer on your development computer.</span></span>
    
2. <span data-ttu-id="8b023-118">右键单击 C:\ 驱动器，然后选择“新建”\*\*\*\* > “文件夹”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>
    
3. <span data-ttu-id="8b023-119">将新文件夹命名为 AddinManifests。</span><span class="sxs-lookup"><span data-stu-id="8b023-119">Name the new folder AddinManifests.</span></span>
    
4. <span data-ttu-id="8b023-120">右键单击 AddinManifests 文件夹，然后选择“共享”\*\*\*\* > “特定用户”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>
    
5. <span data-ttu-id="8b023-121">在“文件共享”\*\*\*\* 中，选择下拉箭头，再依次选择“所有人”\*\*\*\* > “添加”\*\*\*\* > “共享”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>
    
> [!NOTE]
> <span data-ttu-id="8b023-p102">本演练要将本地文件共享用作受信任的目录，用来存储加载项 XML 清单文件。在实际方案中，可以改为选择[将 XML 清单文件部署到 SharePoint 目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)，或[将加载项发布到 AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)。</span><span class="sxs-lookup"><span data-stu-id="8b023-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="8b023-124">第 2 步：将文件共享添加到受信任的加载项目录</span><span class="sxs-lookup"><span data-stu-id="8b023-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1.  <span data-ttu-id="8b023-125">启动 Word 并创建文档。</span><span class="sxs-lookup"><span data-stu-id="8b023-125">Start Word 2016 and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8b023-126">尽管本示例使用的是 Word，但也可以使用任何支持 Office 加载项的 Office 应用（如 Excel、Outlook、PowerPoint 或 Project）。</span><span class="sxs-lookup"><span data-stu-id="8b023-126">Although this example uses Word 2016, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project 2016.</span></span>
    
2.  <span data-ttu-id="8b023-127">选择“文件”\*\*\*\* > “选项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-127">Choose **File** > **Options**.</span></span>
    
3.  <span data-ttu-id="8b023-128">在“Word 选项”\*\*\*\* 对话框中，选择“信任中心”\*\*\*\*，然后选择“信任中心设置”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span> 
    
4.  <span data-ttu-id="8b023-p103">在“信任中心”\*\*\*\* 对话框中，选择“受信任的外接程序目录”\*\*\*\*。输入之前创建的文件共享的通用命名约定 (UNC) 路径，作为**目录 URL**（例如，\\\YourMachineName\AddinManifests）。然后选择“添加目录”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 
    
5. <span data-ttu-id="8b023-131">选中“在菜单中显示”\*\*\*\* 复选框。</span><span class="sxs-lookup"><span data-stu-id="8b023-131">Select the check box for **Show in Menu**.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="8b023-132">如果将加载项 XML 清单文件存储到已指定为受信任的 Web 加载项目录的共享中，用户可以转到功能区中的“插入”\*\*\*\* 选项卡，并选择“我的加载项”\*\*\*\*，此时加载项就会显示在“Office 加载项”\*\*\*\* 对话框中的“共享文件夹”\*\*\*\* 下。</span><span class="sxs-lookup"><span data-stu-id="8b023-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="8b023-133">关闭 Word。</span><span class="sxs-lookup"><span data-stu-id="8b023-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="8b023-134">第 3 步：在 Azure 中创建 Web 应用</span><span class="sxs-lookup"><span data-stu-id="8b023-134">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="8b023-135">使用 [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) 或使用 [Azure 门户](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal)在 Azure 中创建空的 Web 应用。</span><span class="sxs-lookup"><span data-stu-id="8b023-135">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="8b023-136">使用 Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="8b023-136">Using Visual Studio 2017</span></span>

<span data-ttu-id="8b023-137">若要使用 Visual Studio 2017 创建 Web 应用，请完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="8b023-137">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="8b023-p104">在 Visual Studio 的“视图”\*\*\*\* 菜单中，选择“服务器资源管理器”\*\*\*\*。右键单击“Azure”\*\*\*\* 并选择“连接到 Microsoft Azure 订阅”\*\*\*\*。请按说明连接到 Azure 订阅。</span><span class="sxs-lookup"><span data-stu-id="8b023-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>
    
2. <span data-ttu-id="8b023-141">在 Visual Studio 的“服务器资源管理器”\*\*\*\* 中，展开“Azure”\*\*\*\*，右键单击“应用服务”\*\*\*\*，然后选择“创建新的应用服务”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-141">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>
    
3. <span data-ttu-id="8b023-142">在“创建应用服务”\*\*\*\* 对话框中，提供此信息：</span><span class="sxs-lookup"><span data-stu-id="8b023-142">In the **Create App Service** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="8b023-p105">为站点输入唯一的“Web 应用名称”\*\*\*\*。Azure 验证站点名称在整个 azurewebsites.net 域中是否是唯一的。</span><span class="sxs-lookup"><span data-stu-id="8b023-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="8b023-145">选择要用于创建此站点的“订阅”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-145">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="8b023-p106">为站点选择“资源组”\*\*\*\*。如果创建新组，还需为新组命名。</span><span class="sxs-lookup"><span data-stu-id="8b023-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
    
      - <span data-ttu-id="8b023-p107">选择要用于创建此站点的“应用服务计划”\*\*\*\*。如果创建新计划，还需为新计划命名。</span><span class="sxs-lookup"><span data-stu-id="8b023-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="8b023-150">选择“创建”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-150">Choose **Create**.</span></span>

    <span data-ttu-id="8b023-151">新的 Web 应用将在“服务器资源管理器”\*\*\*\* 中的“Azure”\*\*\*\* >> “应用服务”\*\*\*\*>>“选择的资源组”下显示。</span><span class="sxs-lookup"><span data-stu-id="8b023-151">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>
    
4. <span data-ttu-id="8b023-p108">右键单击新的 Web 应用，然后选择“在浏览器中查看”\*\*\*\*。随即打开浏览器，并显示包含“应用服务应用已创建”消息的网页。</span><span class="sxs-lookup"><span data-stu-id="8b023-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>
    
5. <span data-ttu-id="8b023-154">在浏览器地址栏中，将 Web 应用 URL 更改为使用 HTTPS，并按 **Enter** 确认已启用 HTTPS 协议。</span><span class="sxs-lookup"><span data-stu-id="8b023-154">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="8b023-155">Azure 网站自动提供 HTTPS 端点。</span><span class="sxs-lookup"><span data-stu-id="8b023-155">Azure websites automatically provide an HTTPS endpoint.</span></span>
    
### <a name="using-the-azure-portal"></a><span data-ttu-id="8b023-156">使用 Azure 门户</span><span class="sxs-lookup"><span data-stu-id="8b023-156">Using the Azure portal</span></span>

<span data-ttu-id="8b023-157">若要使用 Azure 门户创建 Web 应用，请完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="8b023-157">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="8b023-158">使用 Azure 凭据登录到 [Azure 门户](https://portal.azure.com/)。</span><span class="sxs-lookup"><span data-stu-id="8b023-158">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>
    
2. <span data-ttu-id="8b023-159">选择“新建”\*\*\*\* > “Web + 移动”\*\*\*\* > “Web 应用”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-159">Choose **New** > **Web + Mobile** > **Web App**.</span></span> 

3. <span data-ttu-id="8b023-160">在“Web 应用创建”\*\*\*\* 对话框中，提供此信息：</span><span class="sxs-lookup"><span data-stu-id="8b023-160">In the **Web App Create** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="8b023-p109">为站点输入唯一的“应用名称”\*\*\*\*。Azure 验证站点名称在整个 azureweb apps.net 域中是否是唯一的。</span><span class="sxs-lookup"><span data-stu-id="8b023-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="8b023-163">选择要用于创建此站点的“订阅”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-163">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="8b023-p110">为站点选择“资源组”\*\*\*\*。如果创建新组，还需为新组命名。</span><span class="sxs-lookup"><span data-stu-id="8b023-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="8b023-166">为站点选择“操作系统”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-166">Choose the **OS** for your site.</span></span>
    
      - <span data-ttu-id="8b023-p111">选择要用于创建此站点的“应用服务计划”\*\*\*\*。如果创建新计划，还需为新计划命名。</span><span class="sxs-lookup"><span data-stu-id="8b023-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="8b023-169">选择“创建”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-169">Choose **Create**.</span></span>

4. <span data-ttu-id="8b023-170">选择“通知”\*\*\*\*（Azure 门户顶部边缘的钟形图标），然后选择“部署成功”\*\*\*\* 通知，以打开 Azure 门户中的站点“概述”\*\*\*\* 页。</span><span class="sxs-lookup"><span data-stu-id="8b023-170">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8b023-171">网站部署完成后，通知会从“正在部署”\*\*\*\* 更改为“部署成功”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-171">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="8b023-p112">在 Azure 门户的站点“概述”\*\*\*\* 页的“基本信息”\*\*\*\* 部分中，选择“URL”\*\*\*\* 下显示的 URL。随即打开浏览器，并显示包含“应用服务应用已创建”消息的网页。</span><span class="sxs-lookup"><span data-stu-id="8b023-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 
    
6. <span data-ttu-id="8b023-174">在浏览器地址栏中，将 Web 应用 URL 更改为使用 HTTPS，并按 **Enter** 确认已启用 HTTPS 协议。</span><span class="sxs-lookup"><span data-stu-id="8b023-174">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="8b023-175">Azure 网站自动提供 HTTPS 端点。</span><span class="sxs-lookup"><span data-stu-id="8b023-175">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="8b023-176">第 4 步：在 Visual Studio 中创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="8b023-176">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="8b023-177">以管理员身份启动 Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="8b023-177">Start Visual Studio as an administrator.</span></span>
    
2. <span data-ttu-id="8b023-178">选择“文件”\*\*\*\* > “新建”\*\*\*\* > “项目”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-178">Choose **File** > **New** > **Project**.</span></span>
    
3. <span data-ttu-id="8b023-179">在“模板”\*\*\*\* 下，展开“Visual C#”\*\*\*\*（或“Visual Basic”\*\*\*\*），展开“Office/SharePoint”\*\*\*\*，然后选择“外接程序”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-179">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>
    
4. <span data-ttu-id="8b023-180">选择“Word Web 外接程序”\*\*\*\*，然后选择“确定”\*\*\*\* 以接受默认设置。</span><span class="sxs-lookup"><span data-stu-id="8b023-180">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>
       
<span data-ttu-id="8b023-181">Visual Studio 将创建基本的 Word 外接程序，你可以按原样发布，无需对其 Web 项目进行任何更改。</span><span class="sxs-lookup"><span data-stu-id="8b023-181">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="8b023-182">第 5 步：将 Office 外接程序 Web 应用发布到 Azure</span><span class="sxs-lookup"><span data-stu-id="8b023-182">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="8b023-183">在 Visual Studio 中打开外接程序项目后，展开“解决方案资源管理器”\*\*\*\* 中的解决方案节点，以便可以查看解决方案的两个项目。</span><span class="sxs-lookup"><span data-stu-id="8b023-183">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>
    
2. <span data-ttu-id="8b023-p113">右键单击 Web 项目，然后选择“发布”\*\*\*\*。Web 项目包含 Office 外接程序 Web 应用文件，因此，这是你可以发布到 Azure 的项目。</span><span class="sxs-lookup"><span data-stu-id="8b023-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>
    
3. <span data-ttu-id="8b023-186">在“发布”\*\*\*\* 选项卡上：</span><span class="sxs-lookup"><span data-stu-id="8b023-186">On the **Publish** tab:</span></span>

      - <span data-ttu-id="8b023-187">选择“Microsoft Azure 应用服务”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-187">Choose **Microsoft Azure App Service**.</span></span>
      
      - <span data-ttu-id="8b023-188">选择“选择现有”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-188">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="8b023-189">选择“发布”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-189">Choose **Publish**.</span></span> 

6. <span data-ttu-id="8b023-190">在“应用服务”\*\*\*\* 对话框中，找到并选择在[步骤 3：在 Azure 中创建 Web 应用](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure)中创建的 Web 应用，然后选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-190">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="8b023-p114">Visual Studio 会将 Office 外接程序的 Web 项目发布到 Azure Web 应用。Visual Studio 完成发布 Web 项目后，浏览器将打开并显示网页，其中显示“应用服务应用已创建”文本。这是 Web 应用当前的默认页。</span><span class="sxs-lookup"><span data-stu-id="8b023-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

7. <span data-ttu-id="8b023-194">要查看外接程序的网页，请更改 URL 以便它使用 HTTPS 并指定外接程序 HTML 页面的路径（例如：https://YourDomain.azurewebsites.net/Home.html)）。</span><span class="sxs-lookup"><span data-stu-id="8b023-194">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html).</span></span> <span data-ttu-id="8b023-195">这可确认你的外接程序的 Web 应用现在托管于 Azure 上。</span><span class="sxs-lookup"><span data-stu-id="8b023-195">This confirms that your add-in's website is now hosted on Azure.</span></span> <span data-ttu-id="8b023-196">复制根 URL（例如 https://YourDomain.azurewebsites.net)），在本文稍后编辑外接程序清单文件时将需要此 URL。</span><span class="sxs-lookup"><span data-stu-id="8b023-196">Copy this URL because you'll need it when you edit the add-in manifest file later in this topic.</span></span>
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="8b023-197">步骤 6：编辑并部署外接程序 XML 清单文件</span><span class="sxs-lookup"><span data-stu-id="8b023-197">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="8b023-198">在示例 Office 外接程序在“解决方案资源管理器”\*\*\*\* 中打开的 Visual Studio 中，展开该解决方案以显示两个项目。</span><span class="sxs-lookup"><span data-stu-id="8b023-198">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>
    
2. <span data-ttu-id="8b023-p116">展开 Office 外接程序项目（例如 WordWebAddIn），右键单击清单文件夹，然后选择“打开”\*\*\*\*。随即打开外接程序 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="8b023-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>
    
3. <span data-ttu-id="8b023-201">在 XML 清单文件中，找到所有的 "~remoteAppUrl" 实例，并将其全部替换为 Azure 上的外接程序 Web 应用的根 URL。</span><span class="sxs-lookup"><span data-stu-id="8b023-201">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> <span data-ttu-id="8b023-202">这就是之前在将外接程序 Web 应用发布到 Azure 后复制的 URL（例如：https://YourDomain.azurewebsites.net)）。</span><span class="sxs-lookup"><span data-stu-id="8b023-202">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 
    
4. <span data-ttu-id="8b023-p118">选择“文件”\*\*\*\*，然后选择“全部保存”\*\*\*\*。关闭外接程序 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="8b023-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>
    
5. <span data-ttu-id="8b023-205">返回到“解决方案资源管理器”\*\*\*\*，右键单击清单文件夹并选择“在文件资源管理器中打开文件夹”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-205">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>
    
6. <span data-ttu-id="8b023-206">复制外接程序 XML 清单文件（例如 WordWebAddIn.xml）。</span><span class="sxs-lookup"><span data-stu-id="8b023-206">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 
    
7. <span data-ttu-id="8b023-207">浏览到在[步骤 1：创建共享文件夹](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)中创建的网络文件共享，并将清单文件粘贴到此文件夹。</span><span class="sxs-lookup"><span data-stu-id="8b023-207">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="8b023-208">步骤 7：在 Office 客户端应用程序中插入并运行加载项</span><span class="sxs-lookup"><span data-stu-id="8b023-208">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="8b023-209">启动 Word 并创建文档。</span><span class="sxs-lookup"><span data-stu-id="8b023-209">Start Word 2016 and create a document.</span></span>
    
2. <span data-ttu-id="8b023-210">在功能区中选择“插入”\*\*\*\* > “我的外接程序”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8b023-210">On the ribbon, choose **Insert** > **My Add-ins**.</span></span> 
    
3. <span data-ttu-id="8b023-p119">在“Office 外接程序”\*\*\*\* 对话框中，选择“共享文件夹”\*\*\*\*。Word 扫描已列为受信任的外接程序目录（在[步骤 2：将文件共享添加到受信任的外接程序目录](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)）的文件夹，并在对话框中显示外接程序。应该会看到示例外接程序的图标。</span><span class="sxs-lookup"><span data-stu-id="8b023-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>
    
4. <span data-ttu-id="8b023-p120">选择你的外接程序的图标，然后选择“添加”\*\*\*\*。外接程序的“显示任务窗格”\*\*\*\* 按钮将添加到功能区。</span><span class="sxs-lookup"><span data-stu-id="8b023-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span> 

5. <span data-ttu-id="8b023-p121">在“主页”\*\*\*\* 选项卡的功能区上，选择“显示任务窗格”\*\*\*\* 按钮。外接程序在当前文档右侧的任务窗格中打开。</span><span class="sxs-lookup"><span data-stu-id="8b023-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>
    
6. <span data-ttu-id="8b023-p122">选中文档中的某文本，并选择任务窗格中的“突出显示!”\*\*\*\* 按钮，验证加载项是否正常运行。</span><span class="sxs-lookup"><span data-stu-id="8b023-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span> 

## <a name="see-also"></a><span data-ttu-id="8b023-220">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8b023-220">See also</span></span>

- [<span data-ttu-id="8b023-221">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="8b023-221">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="8b023-222">使用 Visual Studio 打包外接程序以准备发布</span><span class="sxs-lookup"><span data-stu-id="8b023-222">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
    
