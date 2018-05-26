---
title: '? Microsoft Azure ??? Office ???'
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: f0d6a5a10d2ce0620b42566be03e2d36f8a922f2
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="dc9e1-102">? Microsoft Azure ??? Office ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-102">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="dc9e1-p101">???? Office ????? XML ????? HTML ????XML ???????????????????????????? Office ????????????? HTML ?? URL?HTML ?????? Web ??????? Office ????????????????????? Web ?????????? Office ????? Web ??????? Web ??????? Azure???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="dc9e1-107">???????????? Web ????? Azure ?[???????](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)?? Office ?????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-107">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="dc9e1-108">????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-108">Prerequisites</span></span> 

1. <span data-ttu-id="dc9e1-109">?? [Visual Studio 2017](https://www.visualstudio.com/downloads)?????? **Azure ??**?????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-109">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="dc9e1-110">??????? Visual Studio 2017??[?? Visual Studio ????](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio)?????? **Azure ??**?????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="dc9e1-111">?? Office 2016?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-111">Install Office 2016.</span></span> 
    
    > [!NOTE]
    > <span data-ttu-id="dc9e1-112">?????? Office 2016???[?? 1 ???????](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786)?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-112">If you don't already have Office 2016, you can [register for a free 1-month trial](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span></span>

3.  <span data-ttu-id="dc9e1-113">?? Azure ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-113">Obtain an Azure subscription.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="dc9e1-114">????? Azure ?????[?? MSDN ???? Azure ??](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/)????[???????](https://azure.microsoft.com/en-us/pricing/free-trial)?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-114">If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/) or [register for a free trial](https://azure.microsoft.com/en-us/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="dc9e1-115">? 1 ??????????? XML ??????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-115">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="dc9e1-116">????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-116">Open File Explorer on your development computer.</span></span>
    
2. <span data-ttu-id="dc9e1-117">???? C:\ ????????????**** > ?????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-117">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>
    
3. <span data-ttu-id="dc9e1-118">???????? AddinManifests?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-118">Name the new folder AddinManifests.</span></span>
    
4. <span data-ttu-id="dc9e1-119">???? AddinManifests ????????????**** > ??????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-119">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>
    
5. <span data-ttu-id="dc9e1-120">???????****???????????????????**** > ????**** > ????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-120">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>
    
> [!NOTE]
> <span data-ttu-id="dc9e1-p102">??????????????????????????? XML ??????????????????[? XML ??????? SharePoint ??](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)??[??????? AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="dc9e1-123">? 2 ???????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-123">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1.  <span data-ttu-id="dc9e1-124">?? Word 2016 ??????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-124">Start Word 2016 and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="dc9e1-125">????????? Word 2016??????????? Office ???? Office ???? Excel?Outlook?PowerPoint ? Project 2016??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-125">Although this example uses Word 2016, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project 2016.</span></span>
    
2.  <span data-ttu-id="dc9e1-126">??????**** > ????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-126">Choose **File** > **Options**.</span></span>
    
3.  <span data-ttu-id="dc9e1-127">??Word ???****?????????????****?????????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-127">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span> 
    
4.  <span data-ttu-id="dc9e1-p103">???????****???????????????????****??????????????????? (UNC) ?????**?? URL**????\\\YourMachineName\AddinManifests????????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 
    
5. <span data-ttu-id="dc9e1-130">??????????****????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-130">Select the check box for **Show in Menu**.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="dc9e1-131">?????? XML ??????????????? Web ?????????????????????????****??????????????****????????????Office ????****????????????****??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-131">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="dc9e1-132">?? Word 2016?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-132">Close Word 2016.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="dc9e1-133">? 3 ??? Azure ??? Web ??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-133">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="dc9e1-134">?? [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) ??? [Azure ??](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal)? Azure ????? Web ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-134">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="dc9e1-135">?? Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="dc9e1-135">Using Visual Studio 2017</span></span>

<span data-ttu-id="dc9e1-136">???? Visual Studio 2017 ?? Web ???????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-136">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="dc9e1-p104">? Visual Studio ?????****????????????????****??????Azure?**** ??????? Microsoft Azure ???****???????? Azure ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>
    
2. <span data-ttu-id="dc9e1-140">? Visual Studio ???????????****?????Azure?****???????????****???????????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-140">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>
    
3. <span data-ttu-id="dc9e1-141">?????????****???????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-141">In the **Create App Service** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="dc9e1-p105">?????????Web ?????****?Azure ????????? azurewebsites.net ?????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="dc9e1-144">???????????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-144">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="dc9e1-p106">??????????****????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
    
      - <span data-ttu-id="dc9e1-p107">???????????????????****??????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="dc9e1-149">??????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-149">Choose **Create**.</span></span>

    <span data-ttu-id="dc9e1-150">?? Web ??????????????****???Azure?**** >> ??????****>>????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-150">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>
    
4. <span data-ttu-id="dc9e1-p108">?????? Web ????????????????****???????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>
    
5. <span data-ttu-id="dc9e1-153">?????????? Web ?? URL ????? HTTPS??? **Enter** ????? HTTPS ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-153">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="dc9e1-154"> Azure ?????? HTTPS ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-154">Azure websites automatically provide an HTTPS endpoint.</span></span>
    
### <a name="using-the-azure-portal"></a><span data-ttu-id="dc9e1-155">?? Azure ??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-155">Using the Azure portal</span></span>

<span data-ttu-id="dc9e1-156">???? Azure ???? Web ???????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-156">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="dc9e1-157">?? Azure ????? [Azure ??](https://portal.azure.com/)?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-157">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>
    
2. <span data-ttu-id="dc9e1-158">??????**** > ?Web + ???**** > ?Web ???****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-158">Choose **New** > **Web + Mobile** > **Web App**.</span></span> 

3. <span data-ttu-id="dc9e1-159">??Web ?????****???????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-159">In the **Web App Create** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="dc9e1-p109">??????????????****?Azure ????????? azureweb apps.net ?????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="dc9e1-162">???????????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-162">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="dc9e1-p110">??????????****????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="dc9e1-165">???????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-165">Choose the **OS** for your site.</span></span>
    
      - <span data-ttu-id="dc9e1-p111">???????????????????****??????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="dc9e1-168">??????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-168">Choose **Create**.</span></span>

4. <span data-ttu-id="dc9e1-169">??????****?Azure ???????????????????????****?????? Azure ??????????****??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-169">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="dc9e1-170">??????????????????****?????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-170">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="dc9e1-p112">? Azure ?????????****????????****???????URL?****???? URL???????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 
    
6. <span data-ttu-id="dc9e1-173">?????????? Web ?? URL ????? HTTPS??? **Enter** ????? HTTPS ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-173">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="dc9e1-174"> Azure ?????? HTTPS ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-174">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="dc9e1-175">? 4 ??? Visual Studio ??? Office ????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-175">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="dc9e1-176">???????? Visual Studio?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-176">Start Visual Studio as an administrator.</span></span>
    
2. <span data-ttu-id="dc9e1-177">??????**** > ????**** > ????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-177">Choose **File** > **New** > **Project**.</span></span>
    
3. <span data-ttu-id="dc9e1-178">?????****?????Visual C#?****???Visual Basic?****?????Office/SharePoint?****???????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-178">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>
    
4. <span data-ttu-id="dc9e1-179">???Word Web ?????****?????????****????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-179">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>
       
<span data-ttu-id="dc9e1-180">Visual Studio ?????? Word ?????????????????? Web ?????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-180">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="dc9e1-181">? 5 ??? Office ???? Web ????? Azure</span><span class="sxs-lookup"><span data-stu-id="dc9e1-181">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="dc9e1-182">? Visual Studio ????????????????????????****?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-182">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>
    
2. <span data-ttu-id="dc9e1-p113">???? Web ???????????****?Web ???? Office ???? Web ???????????????? Azure ????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>
    
3. <span data-ttu-id="dc9e1-185">?????****?????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-185">On the **Publish** tab:</span></span>

      - <span data-ttu-id="dc9e1-186">???Microsoft Azure ?????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-186">Choose **Microsoft Azure App Service**.</span></span>
      
      - <span data-ttu-id="dc9e1-187">????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-187">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="dc9e1-188">??????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-188">Choose **Publish**.</span></span> 

6. <span data-ttu-id="dc9e1-189">???????****???????????[?? 3?? Azure ??? Web ??](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure)???? Web ???????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-189">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="dc9e1-p114">Visual Studio ?? Office ????? Web ????? Azure Web ???Visual Studio ???? Web ???????????????????????????????????? Web ?????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

7. <span data-ttu-id="dc9e1-193">?????????????? URL ????? HTTPS ??????? HTML ?????????https://YourDomain.azurewebsites.net/Home.html)??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-193">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html).</span></span> <span data-ttu-id="dc9e1-194">?????????? Web ??????? Azure ??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-194">This confirms that your add-in's website is now hosted on Azure.</span></span> <span data-ttu-id="dc9e1-195">??? URL??? https://YourDomain.azurewebsites.net)?????????????????????? URL?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-195">Copy this URL because you'll need it when you edit the add-in manifest file later in this topic.</span></span>
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="dc9e1-196">? 6 ??????????? XML ????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-196">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="dc9e1-197">??? Office ????????????????****???? Visual Studio ?????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-197">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>
    
2. <span data-ttu-id="dc9e1-p116">?? Office ????????? WordWebAddIn????????????????**??**????????? XML ?????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>
    
3. <span data-ttu-id="dc9e1-200">? XML ??????????? "~remoteAppUrl" ??????????? Azure ?????? Web ???? URL?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-200">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> <span data-ttu-id="dc9e1-201">??????????? Web ????? Azure ???? URL????https://YourDomain.azurewebsites.net)??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-201">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 
    
4. <span data-ttu-id="dc9e1-p118">??** ??**?????**????**??????? XML ?????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>
    
5. <span data-ttu-id="dc9e1-204">??????????????****?????????????????????????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-204">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>
    
6. <span data-ttu-id="dc9e1-205">?????? XML ??????? WordWebAddIn.xml??</span><span class="sxs-lookup"><span data-stu-id="dc9e1-205">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 
    
7. <span data-ttu-id="dc9e1-206">????[?? 1????????](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file)?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-206">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="dc9e1-207">?? 7?? Office ????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-207">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="dc9e1-208">?? Word 2016 ??????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-208">Start Word 2016 and create a document.</span></span>
    
2. <span data-ttu-id="dc9e1-209">???????????**** > ????????****?</span><span class="sxs-lookup"><span data-stu-id="dc9e1-209">On the ribbon, choose **Insert** > **My Add-ins**.</span></span> 
    
3. <span data-ttu-id="dc9e1-p119">??Office ?????****??????????????****?Word ?????????????????[?? 2???????????????????](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>
    
4. <span data-ttu-id="dc9e1-p120">????????????????????****??????????????****??????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span> 

5. <span data-ttu-id="dc9e1-p121">?????****???????????????????****???????????????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>
    
6. <span data-ttu-id="dc9e1-p122">????????????????????????!?****???????????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span> 

## <a name="see-also"></a><span data-ttu-id="dc9e1-219">????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-219">See also</span></span>

- [<span data-ttu-id="dc9e1-220">?? Office ???</span><span class="sxs-lookup"><span data-stu-id="dc9e1-220">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="dc9e1-221">?? Visual Studio ???????????</span><span class="sxs-lookup"><span data-stu-id="dc9e1-221">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
    
