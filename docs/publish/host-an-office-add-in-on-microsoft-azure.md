---
title: 在 Microsoft Azure 上托管 Office 加载项 | Microsoft Docs
description: 了解如何将加载项 Web 应用部署到 Azure 并旁加载该加载项以便在 Office 客户端应用程序中进行测试。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: c9d33823850925d5c05d72422262bf62f78b051e
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159421"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="745a8-103">在 Microsoft Azure 上托管 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="745a8-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="745a8-p101">最简单的 Office 加载项由一个 XML 清单文件和一个 HTML 页面组成。XML 清单文件描述了加载项的特征，如名称、它可以在哪些 Office 桌面客户端中运行，以及外接程序的 HTML 页面的 URL。HTML 页面包含在用户在 Office 客户端应用程序中安装和运行外接程序时与之交互的 web 应用程序。您可以在任何 web 承载平台（包括 Azure）上托管 Office 外接程序的 web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="745a8-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop clients it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="745a8-108">本文介绍了如何将外接程序 Web 应用部署到 Azure 并[旁加载外接程序](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)以在 Office 客户端应用程序中进行测试。</span><span class="sxs-lookup"><span data-stu-id="745a8-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="745a8-109">先决条件</span><span class="sxs-lookup"><span data-stu-id="745a8-109">Prerequisites</span></span> 

1. <span data-ttu-id="745a8-110">安装 [Visual Studio 2019](https://www.visualstudio.com/downloads)，并选择添加 **Azure 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="745a8-110">Install [Visual Studio 2019](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="745a8-111">如果之前已安装 Visual Studio 2019，请[使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Azure 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="745a8-111">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="745a8-112">安装 Office。</span><span class="sxs-lookup"><span data-stu-id="745a8-112">Install Office.</span></span>

    > [!NOTE]
    > <span data-ttu-id="745a8-113">如果尚未安装 Office，可以[注册 1 个月免费试用版](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)。</span><span class="sxs-lookup"><span data-stu-id="745a8-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3. <span data-ttu-id="745a8-114">获取 Azure 订阅。</span><span class="sxs-lookup"><span data-stu-id="745a8-114">Obtain an Azure subscription.</span></span>

    > [!NOTE]
    > <span data-ttu-id="745a8-115">如果还没有 Azure 订阅，可以[通过 Visual Studio 订阅获取 Azure 订阅](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/)，也可以[注册免费试用版](https://azure.microsoft.com/pricing/free-trial)。</span><span class="sxs-lookup"><span data-stu-id="745a8-115">If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="745a8-116">第 1 步：创建用于托管加载项 XML 清单文件的共享文件夹</span><span class="sxs-lookup"><span data-stu-id="745a8-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="745a8-117">打开开发计算机的文件资源管理器。</span><span class="sxs-lookup"><span data-stu-id="745a8-117">Open File Explorer on your development computer.</span></span>

2. <span data-ttu-id="745a8-118">右键单击 C:\ 驱动器，然后选择“新建”\*\*\*\* > “文件夹”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>

3. <span data-ttu-id="745a8-119">将新文件夹命名为 AddinManifests。</span><span class="sxs-lookup"><span data-stu-id="745a8-119">Name the new folder AddinManifests.</span></span>

4. <span data-ttu-id="745a8-120">右键单击 AddinManifests 文件夹，然后选择“共享”\*\*\*\* > “特定用户”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>

5. <span data-ttu-id="745a8-121">在“文件共享”\*\*\*\* 中，选择下拉箭头，再依次选择“所有人”\*\*\*\* > “添加”\*\*\*\* > “共享”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>

> [!NOTE]
> <span data-ttu-id="745a8-p102">本演练要将本地文件共享用作受信任的目录，用来存储加载项 XML 清单文件。在实际方案中，可以改为选择[将 XML 清单文件部署到 SharePoint 目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)，或[将加载项发布到 AppSource](/office/dev/store/submit-to-appsource-via-partner-center)。</span><span class="sxs-lookup"><span data-stu-id="745a8-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="745a8-124">第 2 步：将文件共享添加到受信任的加载项目录</span><span class="sxs-lookup"><span data-stu-id="745a8-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1. <span data-ttu-id="745a8-125">启动 Word 并创建文档。</span><span class="sxs-lookup"><span data-stu-id="745a8-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="745a8-126">尽管本示例使用的是 Word，但也可以使用任何支持 Office 加载项的 Office 应用（如 Excel、Outlook、PowerPoint 或 Project）。</span><span class="sxs-lookup"><span data-stu-id="745a8-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="745a8-127">选择“**文件**” > “**选项**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-127">Choose **File** > **Options**.</span></span>

3. <span data-ttu-id="745a8-128">在“Word 选项”\*\*\*\* 对话框中，选择“信任中心”\*\*\*\*，然后选择“信任中心设置”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>

4. <span data-ttu-id="745a8-p103">在“信任中心”\*\*\*\* 对话框中，选择“受信任的外接程序目录”\*\*\*\*。输入之前创建的文件共享的通用命名约定 (UNC) 路径，作为**目录 URL**（例如，\\\YourMachineName\AddinManifests）。然后选择“添加目录”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 

5. <span data-ttu-id="745a8-131">选中“在菜单中显示”\*\*\*\* 复选框。</span><span class="sxs-lookup"><span data-stu-id="745a8-131">Select the check box for **Show in Menu**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="745a8-132">如果将加载项 XML 清单文件存储到已指定为受信任的 Web 加载项目录的共享中，用户可以转到功能区中的“插入”\*\*\*\* 选项卡，并选择“我的加载项”\*\*\*\*，此时加载项就会显示在“Office 加载项”\*\*\*\* 对话框中的“共享文件夹”\*\*\*\* 下。</span><span class="sxs-lookup"><span data-stu-id="745a8-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="745a8-133">关闭 Word。</span><span class="sxs-lookup"><span data-stu-id="745a8-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a><span data-ttu-id="745a8-134">第 3 步：使用 Azure 门户在 Azure 中创建 Web 应用</span><span class="sxs-lookup"><span data-stu-id="745a8-134">Step 3: Create a web app in Azure using the Azure portal</span></span>

<span data-ttu-id="745a8-135">若要使用 Azure 门户创建 Web 应用，请完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="745a8-135">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="745a8-136">使用 Azure 凭据登录到 [Azure 门户](https://portal.azure.com/)。</span><span class="sxs-lookup"><span data-stu-id="745a8-136">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>

2. <span data-ttu-id="745a8-137">在“**Azure 服务**”下，选择“**Web 应用**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-137">Under **Azure Services** select **Web Apps**.</span></span>

3. <span data-ttu-id="745a8-138">在“**应用服务**”页面上，选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-138">On the **App Service** page, select **Add**.</span></span> <span data-ttu-id="745a8-139">提供以下信息：</span><span class="sxs-lookup"><span data-stu-id="745a8-139">Provide this information:</span></span>

      - <span data-ttu-id="745a8-140">选择要用于创建此站点的“订阅”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-140">Choose the **Subscription** to use for creating this site.</span></span>
      
      - <span data-ttu-id="745a8-p105">为站点选择“资源组”\*\*\*\*。如果创建新组，还需为新组命名。</span><span class="sxs-lookup"><span data-stu-id="745a8-p105">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
      
      - <span data-ttu-id="745a8-143">为站点输入唯一的“应用名称”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-143">Enter a unique **App name** for your site.</span></span> <span data-ttu-id="745a8-144">Azure 验证站点名称在整个 azureweb apps.net 域中是否是唯一的。</span><span class="sxs-lookup"><span data-stu-id="745a8-144">Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="745a8-145">选择使用代码还是 Docker 容器进行发布。</span><span class="sxs-lookup"><span data-stu-id="745a8-145">Choose whether to publish using code or a docker container.</span></span>

      - <span data-ttu-id="745a8-146">指定“**运行时堆栈**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-146">Specify a **Runtime stack**.</span></span>

      - <span data-ttu-id="745a8-147">为站点选择“**操作系统**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-147">Choose the **OS** for your site.</span></span>

      - <span data-ttu-id="745a8-148">选择“**区域**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-148">Choose a **Region**.</span></span>

      - <span data-ttu-id="745a8-149">选择要用于创建此站点的“**应用服务计划**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-149">Choose the **App Service plan** to use for creating this site.</span></span>

      - <span data-ttu-id="745a8-150">选择“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-150">Choose **Create**.</span></span>

4. <span data-ttu-id="745a8-151">下一页将显示部署的进行状态和完成时间。</span><span class="sxs-lookup"><span data-stu-id="745a8-151">The next page will let you know that your deployment is underway and when it completes.</span></span> <span data-ttu-id="745a8-152">完成后，选择“**转到资源**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-152">When it is completed, select **Go to resource**.</span></span>  

5. <span data-ttu-id="745a8-153">在“**概述**”节中，选择在“**URL**”下显示的 URL。</span><span class="sxs-lookup"><span data-stu-id="745a8-153">In the **Overview** section, choose the URL that is displayed under **URL**.</span></span> <span data-ttu-id="745a8-154">随即将打开浏览器，并显示包含“应用服务应用已启动且正在运行”消息的网页。</span><span class="sxs-lookup"><span data-stu-id="745a8-154">Your browser opens and displays a webpage with the message "Your App Service app is up and running."</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="745a8-155">Azure 网站自动提供 HTTPS 终结点。</span><span class="sxs-lookup"><span data-stu-id="745a8-155">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="745a8-156">第 4 步：在 Visual Studio 中创建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="745a8-156">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="745a8-157">以管理员身份启动 Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="745a8-157">Start Visual Studio as an administrator.</span></span>

2. <span data-ttu-id="745a8-158">选择“**创建新项目**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-158">Choose **Create a new project**.</span></span>

3. <span data-ttu-id="745a8-159">使用搜索框，输入“**加载程序**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-159">Using the search box, enter **add-in**.</span></span>

4. <span data-ttu-id="745a8-160">选择“**Word Web 外接程序**”作为项目类型，然后选择“**下一步**”以接受默认设置。</span><span class="sxs-lookup"><span data-stu-id="745a8-160">Choose **Word Web Add-in** as the project type, and then choose **Next** to accept the default settings.</span></span>

<span data-ttu-id="745a8-161">Visual Studio 将创建基本的 Word 外接程序，你可以按原样发布，无需对其 Web 项目进行任何更改。</span><span class="sxs-lookup"><span data-stu-id="745a8-161">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span> <span data-ttu-id="745a8-162">若要为其他 Office 主机类型（如 Excel）生成外接程序，请重复这些步骤，并选择具有所需 Office 主机的项目类型。</span><span class="sxs-lookup"><span data-stu-id="745a8-162">To make an add-in for a different Office host type, such as Excel, repeat the steps and choose a project type with your desired Office host.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="745a8-163">第 5 步：将 Office 外接程序 Web 应用发布到 Azure</span><span class="sxs-lookup"><span data-stu-id="745a8-163">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="745a8-164">在 Visual Studio 中打开外接程序项目后，展开“**解决方案资源管理器**”中的解决方案节点，然后选择“**应用服务**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-164">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer**, then select **App Service**.</span></span>

2. <span data-ttu-id="745a8-p110">右键单击 Web 项目，然后选择“发布”\*\*\*\*。Web 项目包含 Office 外接程序 Web 应用文件，因此，这是你可以发布到 Azure 的项目。</span><span class="sxs-lookup"><span data-stu-id="745a8-p110">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>

3. <span data-ttu-id="745a8-167">在“发布”\*\*\*\* 选项卡上：</span><span class="sxs-lookup"><span data-stu-id="745a8-167">On the **Publish** tab:</span></span>

      - <span data-ttu-id="745a8-168">选择“Microsoft Azure 应用服务”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-168">Choose **Microsoft Azure App Service**.</span></span>

      - <span data-ttu-id="745a8-169">选择“选择现有”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-169">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="745a8-170">选择“发布”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="745a8-170">Choose **Publish**.</span></span>

4. <span data-ttu-id="745a8-p111">Visual Studio 会将 Office 外接程序的 Web 项目发布到 Azure Web 应用。Visual Studio 完成发布 Web 项目后，浏览器将打开并显示网页，其中显示“应用服务应用已创建”文本。这是 Web 应用当前的默认页。</span><span class="sxs-lookup"><span data-stu-id="745a8-p111">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

5. <span data-ttu-id="745a8-174">复制根 URL（例如：https://YourDomain.azurewebsites.net)；本文后续部分中编辑加载项清单文件时将需要此 URL。</span><span class="sxs-lookup"><span data-stu-id="745a8-174">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="745a8-175">第 6 步：编辑并部署加载项 XML 清单文件</span><span class="sxs-lookup"><span data-stu-id="745a8-175">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="745a8-176">在示例 Office 外接程序在“解决方案资源管理器”\*\*\*\* 中打开的 Visual Studio 中，展开该解决方案以显示两个项目。</span><span class="sxs-lookup"><span data-stu-id="745a8-176">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>

2. <span data-ttu-id="745a8-p112">展开 Office 加载项项目（例如 WordWebAddIn），右键单击清单文件夹，然后选择“**打开**”。随即打开加载项 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="745a8-p112">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>

3. <span data-ttu-id="745a8-p113">在 XML 清单文件中，找到所有的“~remoteAppUr”实例，并将其全部替换为 Azure 上的外接程序 Web 应用的根 URL。这就是之前在将外接程序 Web 应用发布到 Azure 后复制的 URL（例如：https://YourDomain.azurewebsites.net)）。</span><span class="sxs-lookup"><span data-stu-id="745a8-p113">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 

4. <span data-ttu-id="745a8-181">选择“**文件**”，然后选择“**全部保存**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-181">Choose **File** and then choose **Save All**.</span></span> <span data-ttu-id="745a8-182">然后复制外接程序 XML 清单文件（例如 WordWebAddIn.xml）。</span><span class="sxs-lookup"><span data-stu-id="745a8-182">Next, Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span>

5. <span data-ttu-id="745a8-183">使用“**文件资源管理器**”程序浏览到在[第 1 步：创建共享文件夹](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)中创建的网络文件共享，并将清单文件粘贴到此文件夹。</span><span class="sxs-lookup"><span data-stu-id="745a8-183">Using the **File Explorer** program, browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="745a8-184">第 7 步：在 Office 客户端应用程序中插入并运行加载项</span><span class="sxs-lookup"><span data-stu-id="745a8-184">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="745a8-185">启动 Word 并创建文档。</span><span class="sxs-lookup"><span data-stu-id="745a8-185">Start Word and create a document.</span></span>

2. <span data-ttu-id="745a8-186">在功能区中选择“**插入**” > “**我的加载项**”。</span><span class="sxs-lookup"><span data-stu-id="745a8-186">On the ribbon, choose **Insert** > **My Add-ins**.</span></span>

3. <span data-ttu-id="745a8-p115">在“Office 外接程序”\*\*\*\* 对话框中，选择“共享文件夹”\*\*\*\*。Word 扫描已列为受信任的外接程序目录（在[步骤 2：将文件共享添加到受信任的外接程序目录](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)）的文件夹，并在对话框中显示外接程序。应该会看到示例外接程序的图标。</span><span class="sxs-lookup"><span data-stu-id="745a8-p115">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>

4. <span data-ttu-id="745a8-p116">选择你的外接程序的图标，然后选择“添加”\*\*\*\*。外接程序的“显示任务窗格”\*\*\*\* 按钮将添加到功能区。</span><span class="sxs-lookup"><span data-stu-id="745a8-p116">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span>

5. <span data-ttu-id="745a8-p117">在“主页”\*\*\*\* 选项卡的功能区上，选择“显示任务窗格”\*\*\*\* 按钮。外接程序在当前文档右侧的任务窗格中打开。</span><span class="sxs-lookup"><span data-stu-id="745a8-p117">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>

6. <span data-ttu-id="745a8-p118">选中文档中的某文本，并选择任务窗格中的“突出显示!”\*\*\*\* 按钮，验证加载项是否正常运行。</span><span class="sxs-lookup"><span data-stu-id="745a8-p118">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span>

## <a name="see-also"></a><span data-ttu-id="745a8-196">另请参阅</span><span class="sxs-lookup"><span data-stu-id="745a8-196">See also</span></span>

- [<span data-ttu-id="745a8-197">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="745a8-197">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="745a8-198">使用 Visual Studio 发布加载项</span><span class="sxs-lookup"><span data-stu-id="745a8-198">Publish your add-in using Visual Studio</span></span>](../publish/package-your-add-in-using-visual-studio.md)
