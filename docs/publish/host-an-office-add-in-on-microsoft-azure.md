---
title: 在 Microsoft Azure 上托管 Office 加载项 | Microsoft Docs
description: 了解如何将加载项 Web 应用部署到 Azure 并旁加载该加载项以便在 Office 客户端应用程序中进行测试。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: a30f1a8219501a68e6f46f013ef46640a59fe4e9
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094230"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="e8737-103">在 Microsoft Azure 上托管 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="e8737-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="e8737-104">The simplest Office Add-in is made up of an XML manifest file and an HTML page.</span><span class="sxs-lookup"><span data-stu-id="e8737-104">The simplest Office Add-in is made up of an XML manifest file and an HTML page.</span></span> <span data-ttu-id="e8737-105">The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page.</span><span class="sxs-lookup"><span data-stu-id="e8737-105">The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page.</span></span> <span data-ttu-id="e8737-106">The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application.</span><span class="sxs-lookup"><span data-stu-id="e8737-106">The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application.</span></span> <span data-ttu-id="e8737-107">You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span><span class="sxs-lookup"><span data-stu-id="e8737-107">You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="e8737-108">本文介绍了如何将外接程序 Web 应用部署到 Azure 并[旁加载外接程序](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)以在 Office 客户端应用程序中进行测试。</span><span class="sxs-lookup"><span data-stu-id="e8737-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e8737-109">先决条件</span><span class="sxs-lookup"><span data-stu-id="e8737-109">Prerequisites</span></span> 

1. <span data-ttu-id="e8737-110">安装 [Visual Studio 2019](https://www.visualstudio.com/downloads)，并选择添加 **Azure 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="e8737-110">Install [Visual Studio 2019](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e8737-111">如果之前已安装 Visual Studio 2019，请[使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Azure 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="e8737-111">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="e8737-112">安装 Office。</span><span class="sxs-lookup"><span data-stu-id="e8737-112">Install Office.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e8737-113">如果尚未安装 Office，可以[注册 1 个月免费试用版](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)。</span><span class="sxs-lookup"><span data-stu-id="e8737-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3. <span data-ttu-id="e8737-114">获取 Azure 订阅。</span><span class="sxs-lookup"><span data-stu-id="e8737-114">Obtain an Azure subscription.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e8737-115">如果还没有 Azure 订阅，可以[通过 Visual Studio 订阅获取 Azure 订阅](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/)，也可以[注册免费试用版](https://azure.microsoft.com/pricing/free-trial)。</span><span class="sxs-lookup"><span data-stu-id="e8737-115">If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="e8737-116">第 1 步：创建用于托管加载项 XML 清单文件的共享文件夹</span><span class="sxs-lookup"><span data-stu-id="e8737-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="e8737-117">打开开发计算机的文件资源管理器。</span><span class="sxs-lookup"><span data-stu-id="e8737-117">Open File Explorer on your development computer.</span></span>

2. <span data-ttu-id="e8737-118">右键单击 C:\ 驱动器，然后选择“新建”\*\*\*\* > “文件夹”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>

3. <span data-ttu-id="e8737-119">将新文件夹命名为 AddinManifests。</span><span class="sxs-lookup"><span data-stu-id="e8737-119">Name the new folder AddinManifests.</span></span>

4. <span data-ttu-id="e8737-120">右键单击 AddinManifests 文件夹，然后选择“共享”\*\*\*\* > “特定用户”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>

5. <span data-ttu-id="e8737-121">在“文件共享”\*\*\*\* 中，选择下拉箭头，再依次选择“所有人”\*\*\*\* > “添加”\*\*\*\* > “共享”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>

> [!NOTE]
> <span data-ttu-id="e8737-122">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file.</span><span class="sxs-lookup"><span data-stu-id="e8737-122">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file.</span></span> <span data-ttu-id="e8737-123">In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span><span class="sxs-lookup"><span data-stu-id="e8737-123">In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="e8737-124">第 2 步：将文件共享添加到受信任的加载项目录</span><span class="sxs-lookup"><span data-stu-id="e8737-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1. <span data-ttu-id="e8737-125">启动 Word 并创建文档。</span><span class="sxs-lookup"><span data-stu-id="e8737-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e8737-126">尽管本示例使用的是 Word，但也可以使用任何支持 Office 加载项的 Office 应用（如 Excel、Outlook、PowerPoint 或 Project）。</span><span class="sxs-lookup"><span data-stu-id="e8737-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="e8737-127">选择“**文件**” > “**选项**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-127">Choose **File** > **Options**.</span></span>

3. <span data-ttu-id="e8737-128">在“Word 选项”\*\*\*\* 对话框中，选择“信任中心”\*\*\*\*，然后选择“信任中心设置”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>

4. <span data-ttu-id="e8737-129">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**.</span><span class="sxs-lookup"><span data-stu-id="e8737-129">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**.</span></span> <span data-ttu-id="e8737-130">Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span><span class="sxs-lookup"><span data-stu-id="e8737-130">Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 

5. <span data-ttu-id="e8737-131">选中“在菜单中显示”\*\*\*\* 复选框。</span><span class="sxs-lookup"><span data-stu-id="e8737-131">Select the check box for **Show in Menu**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e8737-132">如果将加载项 XML 清单文件存储到已指定为受信任的 Web 加载项目录的共享中，用户可以转到功能区中的“插入”\*\*\*\* 选项卡，并选择“我的加载项”\*\*\*\*，此时加载项就会显示在“Office 加载项”\*\*\*\* 对话框中的“共享文件夹”\*\*\*\* 下。</span><span class="sxs-lookup"><span data-stu-id="e8737-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="e8737-133">关闭 Word。</span><span class="sxs-lookup"><span data-stu-id="e8737-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a><span data-ttu-id="e8737-134">第 3 步：使用 Azure 门户在 Azure 中创建 Web 应用</span><span class="sxs-lookup"><span data-stu-id="e8737-134">Step 3: Create a web app in Azure using the Azure portal</span></span>

<span data-ttu-id="e8737-135">若要使用 Azure 门户创建 Web 应用，请完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="e8737-135">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="e8737-136">使用 Azure 凭据登录到 [Azure 门户](https://portal.azure.com/)。</span><span class="sxs-lookup"><span data-stu-id="e8737-136">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>

2. <span data-ttu-id="e8737-137">在“**Azure 服务**”下，选择“**Web 应用**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-137">Under **Azure Services** select **Web Apps**.</span></span>

3. <span data-ttu-id="e8737-138">在“**应用服务**”页面上，选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-138">On the **App Service** page, select **Add**.</span></span> <span data-ttu-id="e8737-139">提供以下信息：</span><span class="sxs-lookup"><span data-stu-id="e8737-139">Provide this information:</span></span>

      - <span data-ttu-id="e8737-140">选择要用于创建此站点的“订阅”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-140">Choose the **Subscription** to use for creating this site.</span></span>
      
      - <span data-ttu-id="e8737-141">Choose the **Resource Group** for your site.</span><span class="sxs-lookup"><span data-stu-id="e8737-141">Choose the **Resource Group** for your site.</span></span> <span data-ttu-id="e8737-142">If you create a new group, you also need to name it.</span><span class="sxs-lookup"><span data-stu-id="e8737-142">If you create a new group, you also need to name it.</span></span>
      
      - <span data-ttu-id="e8737-143">为站点输入唯一的“应用名称”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-143">Enter a unique **App name** for your site.</span></span> <span data-ttu-id="e8737-144">Azure 验证站点名称在整个 azureweb apps.net 域中是否是唯一的。</span><span class="sxs-lookup"><span data-stu-id="e8737-144">Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="e8737-145">选择使用代码还是 Docker 容器进行发布。</span><span class="sxs-lookup"><span data-stu-id="e8737-145">Choose whether to publish using code or a docker container.</span></span>

      - <span data-ttu-id="e8737-146">指定“**运行时堆栈**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-146">Specify a **Runtime stack**.</span></span>

      - <span data-ttu-id="e8737-147">为站点选择“**操作系统**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-147">Choose the **OS** for your site.</span></span>

      - <span data-ttu-id="e8737-148">选择“**区域**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-148">Choose a **Region**.</span></span>

      - <span data-ttu-id="e8737-149">选择要用于创建此站点的“**应用服务计划**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-149">Choose the **App Service plan** to use for creating this site.</span></span>

      - <span data-ttu-id="e8737-150">选择“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-150">Choose **Create**.</span></span>

4. <span data-ttu-id="e8737-151">下一页将显示部署的进行状态和完成时间。</span><span class="sxs-lookup"><span data-stu-id="e8737-151">The next page will let you know that your deployment is underway and when it completes.</span></span> <span data-ttu-id="e8737-152">完成后，选择“**转到资源**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-152">When it is completed, select **Go to resource**.</span></span>  

5. <span data-ttu-id="e8737-153">在“**概述**”节中，选择在“**URL**”下显示的 URL。</span><span class="sxs-lookup"><span data-stu-id="e8737-153">In the **Overview** section, choose the URL that is displayed under **URL**.</span></span> <span data-ttu-id="e8737-154">随即将打开浏览器，并显示包含“应用服务应用已启动且正在运行”消息的网页。</span><span class="sxs-lookup"><span data-stu-id="e8737-154">Your browser opens and displays a webpage with the message "Your App Service app is up and running."</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="e8737-155">Azure 网站自动提供 HTTPS 终结点。</span><span class="sxs-lookup"><span data-stu-id="e8737-155">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="e8737-156">第 4 步：在 Visual Studio 中创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="e8737-156">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="e8737-157">以管理员身份启动 Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="e8737-157">Start Visual Studio as an administrator.</span></span>

2. <span data-ttu-id="e8737-158">选择“**创建新项目**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-158">Choose **Create a new project**.</span></span>

3. <span data-ttu-id="e8737-159">使用搜索框，输入“**加载程序**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-159">Using the search box, enter **add-in**.</span></span>

4. <span data-ttu-id="e8737-160">选择“**Word Web 外接程序**”作为项目类型，然后选择“**下一步**”以接受默认设置。</span><span class="sxs-lookup"><span data-stu-id="e8737-160">Choose **Word Web Add-in** as the project type, and then choose **Next** to accept the default settings.</span></span>

<span data-ttu-id="e8737-161">Visual Studio 将创建基本的 Word 外接程序，你可以按原样发布，无需对其 Web 项目进行任何更改。</span><span class="sxs-lookup"><span data-stu-id="e8737-161">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span> <span data-ttu-id="e8737-162">若要为其他 Office 主机类型（如 Excel）生成外接程序，请重复这些步骤，并选择具有所需 Office 主机的项目类型。</span><span class="sxs-lookup"><span data-stu-id="e8737-162">To make an add-in for a different Office host type, such as Excel, repeat the steps and choose a project type with your desired Office host.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="e8737-163">第 5 步：将 Office 外接程序 Web 应用发布到 Azure</span><span class="sxs-lookup"><span data-stu-id="e8737-163">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="e8737-164">在 Visual Studio 中打开外接程序项目后，展开“**解决方案资源管理器**”中的解决方案节点，然后选择“**应用服务**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-164">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer**, then select **App Service**.</span></span>

2. <span data-ttu-id="e8737-165">Right-click the web project and then choose **Publish**.</span><span class="sxs-lookup"><span data-stu-id="e8737-165">Right-click the web project and then choose **Publish**.</span></span> <span data-ttu-id="e8737-166">The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span><span class="sxs-lookup"><span data-stu-id="e8737-166">The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>

3. <span data-ttu-id="e8737-167">在“发布”\*\*\*\* 选项卡上：</span><span class="sxs-lookup"><span data-stu-id="e8737-167">On the **Publish** tab:</span></span>

      - <span data-ttu-id="e8737-168">选择“Microsoft Azure 应用服务”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-168">Choose **Microsoft Azure App Service**.</span></span>

      - <span data-ttu-id="e8737-169">选择“选择现有”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-169">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="e8737-170">选择“发布”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e8737-170">Choose **Publish**.</span></span>

4. <span data-ttu-id="e8737-171">Visual Studio publishes the web project for your Office Add-in to your Azure web app.</span><span class="sxs-lookup"><span data-stu-id="e8737-171">Visual Studio publishes the web project for your Office Add-in to your Azure web app.</span></span> <span data-ttu-id="e8737-172">When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created."</span><span class="sxs-lookup"><span data-stu-id="e8737-172">When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created."</span></span> <span data-ttu-id="e8737-173">This is the current default page for the web app.</span><span class="sxs-lookup"><span data-stu-id="e8737-173">This is the current default page for the web app.</span></span>

5. <span data-ttu-id="e8737-174">复制根 URL（例如：https://YourDomain.azurewebsites.net)；本文后续部分中编辑加载项清单文件时将需要此 URL。</span><span class="sxs-lookup"><span data-stu-id="e8737-174">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="e8737-175">第 6 步：编辑并部署加载项 XML 清单文件</span><span class="sxs-lookup"><span data-stu-id="e8737-175">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="e8737-176">在示例 Office 外接程序在“解决方案资源管理器”\*\*\*\* 中打开的 Visual Studio 中，展开该解决方案以显示两个项目。</span><span class="sxs-lookup"><span data-stu-id="e8737-176">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>

2. <span data-ttu-id="e8737-177">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**.</span><span class="sxs-lookup"><span data-stu-id="e8737-177">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**.</span></span> <span data-ttu-id="e8737-178">The add-in XML manifest file opens.</span><span class="sxs-lookup"><span data-stu-id="e8737-178">The add-in XML manifest file opens.</span></span>

3. <span data-ttu-id="e8737-179">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure.</span><span class="sxs-lookup"><span data-stu-id="e8737-179">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure.</span></span> <span data-ttu-id="e8737-180">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="e8737-180">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 

4. <span data-ttu-id="e8737-181">选择“**文件**”，然后选择“**全部保存**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-181">Choose **File** and then choose **Save All**.</span></span> <span data-ttu-id="e8737-182">然后复制外接程序 XML 清单文件（例如 WordWebAddIn.xml）。</span><span class="sxs-lookup"><span data-stu-id="e8737-182">Next, Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span>

5. <span data-ttu-id="e8737-183">使用“**文件资源管理器**”程序浏览到在[第 1 步：创建共享文件夹](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)中创建的网络文件共享，并将清单文件粘贴到此文件夹。</span><span class="sxs-lookup"><span data-stu-id="e8737-183">Using the **File Explorer** program, browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="e8737-184">第 7 步：在 Office 客户端应用程序中插入并运行加载项</span><span class="sxs-lookup"><span data-stu-id="e8737-184">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="e8737-185">启动 Word 并创建文档。</span><span class="sxs-lookup"><span data-stu-id="e8737-185">Start Word and create a document.</span></span>

2. <span data-ttu-id="e8737-186">在功能区中选择“**插入**” > “**我的加载项**”。</span><span class="sxs-lookup"><span data-stu-id="e8737-186">On the ribbon, choose **Insert** > **My Add-ins**.</span></span>

3. <span data-ttu-id="e8737-187">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**.</span><span class="sxs-lookup"><span data-stu-id="e8737-187">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**.</span></span> <span data-ttu-id="e8737-188">Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box.</span><span class="sxs-lookup"><span data-stu-id="e8737-188">Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box.</span></span> <span data-ttu-id="e8737-189">You should see an icon for your sample add-in.</span><span class="sxs-lookup"><span data-stu-id="e8737-189">You should see an icon for your sample add-in.</span></span>

4. <span data-ttu-id="e8737-190">Choose the icon for your add-in and then choose **Add**.</span><span class="sxs-lookup"><span data-stu-id="e8737-190">Choose the icon for your add-in and then choose **Add**.</span></span> <span data-ttu-id="e8737-191">A **Show Taskpane** button for your add-in is added to the ribbon.</span><span class="sxs-lookup"><span data-stu-id="e8737-191">A **Show Taskpane** button for your add-in is added to the ribbon.</span></span>

5. <span data-ttu-id="e8737-192">On the ribbon of the **Home** tab, choose the **Show Taskpane** button.</span><span class="sxs-lookup"><span data-stu-id="e8737-192">On the ribbon of the **Home** tab, choose the **Show Taskpane** button.</span></span> <span data-ttu-id="e8737-193">The add-in opens in a task pane to the right of the current document.</span><span class="sxs-lookup"><span data-stu-id="e8737-193">The add-in opens in a task pane to the right of the current document.</span></span>

6. <span data-ttu-id="e8737-194">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!**</span><span class="sxs-lookup"><span data-stu-id="e8737-194">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!**</span></span> <span data-ttu-id="e8737-195">button in the task pane.</span><span class="sxs-lookup"><span data-stu-id="e8737-195">button in the task pane.</span></span>

## <a name="see-also"></a><span data-ttu-id="e8737-196">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e8737-196">See also</span></span>

- [<span data-ttu-id="e8737-197">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="e8737-197">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="e8737-198">使用 Visual Studio 发布加载项</span><span class="sxs-lookup"><span data-stu-id="e8737-198">Publish your add-in using Visual Studio</span></span>](../publish/package-your-add-in-using-visual-studio.md)
