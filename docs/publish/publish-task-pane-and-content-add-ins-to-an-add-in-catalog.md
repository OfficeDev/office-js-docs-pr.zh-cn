---
title: 将任务窗格和内容加载项发布到 SharePoint 目录
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 6bf63c36d952b901faaa16b0d93748023ac0fef9
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925295"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="4fa20-102">将任务窗格和内容加载项发布到 SharePoint 目录</span><span class="sxs-lookup"><span data-stu-id="4fa20-102">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="4fa20-p101">加载项目录是 SharePoint Web 应用或 SharePoint Online 租赁中的专用网站集，用于托管 Office 和 SharePoint 加载项的文档库。若要向组织用户分发 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的加载项目录。如果管理员将加载项目录注册为受信任的目录，用户就可以通过 Office 客户端应用中的插入 UI 插入加载项。</span><span class="sxs-lookup"><span data-stu-id="4fa20-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="4fa20-105">SharePoint 上的加载项目录不支持在[加载项清单](../develop/add-in-manifests.md)的 `VersionOverrides` 节点中实现的加载项功能（如加载项命令）。</span><span class="sxs-lookup"><span data-stu-id="4fa20-105">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="4fa20-106">如果面向的是云或混合环境，建议通过 [Office 365 管理中心使用集中部署](../publish/centralized-deployment.md)来发布加载项。</span><span class="sxs-lookup"><span data-stu-id="4fa20-106">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="4fa20-p102">Office 2016 for Mac 不支持 SharePoint 目录。若要向 Mac 客户端部署 Office 加载项，必须将它们提交到 [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)。</span><span class="sxs-lookup"><span data-stu-id="4fa20-p102">SharePoint catalogs are not supported for Office 2016 for Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="set-up-an-add-in-catalog"></a><span data-ttu-id="4fa20-109">设置加载项目录</span><span class="sxs-lookup"><span data-stu-id="4fa20-109">Set up an add-in catalog</span></span>

<span data-ttu-id="4fa20-110">完成以下部分之一中的步骤，以在 SharePoint 或 Office 365 上设置加载项目录。</span><span class="sxs-lookup"><span data-stu-id="4fa20-110">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a><span data-ttu-id="4fa20-111">在 SharePoint 上设置加载项目录</span><span class="sxs-lookup"><span data-stu-id="4fa20-111">To set up an add-in catalog on SharePoint</span></span>

1. <span data-ttu-id="4fa20-112">浏览到“**管理中心网站**”（“**开始**” > “**所有程序**” > “**Microsoft SharePoint 2013 产品**” > “**SharePoint 2013 管理中心**”）。</span><span class="sxs-lookup"><span data-stu-id="4fa20-112">Browse to the  **Central Administration Site** ( **Start** > **All Programs** > **Microsoft SharePoint 2013 Products** > **SharePoint 2013 Central Administration**).</span></span>
    
2. <span data-ttu-id="4fa20-113">在左侧的任务窗格中，选择“**外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-113">In the left task pane, choose  **Add-ins**.</span></span>
    
3. <span data-ttu-id="4fa20-114">在“**外接程序**”页面的“**外接程序管理**”下，选择“**管理外接程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-114">On the  **Add-ins** page, under **Add-in Management**, choose  **Manage Add-in Catalog**.</span></span>
    
4. <span data-ttu-id="4fa20-115">在“**管理外接程序目录**”页上，确保在“**Web 应用程序选择器**”中选择了正确的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="4fa20-115">On the  **Manage Add-in Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>
    
5. <span data-ttu-id="4fa20-116">选择“**查看网站设置**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-116">Choose  **View site settings**.</span></span>
    
6. <span data-ttu-id="4fa20-117">在“**网站设置**”页上选择“**网站集管理员**”以指定网站集管理员，然后选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-117">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>
    
7. <span data-ttu-id="4fa20-118">若要向用户授予网站权限，请选择“**网站权限**”，然后选择“**授予权限**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-118">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>
    
8. <span data-ttu-id="4fa20-119">在“**共享‘应用程序目录网站’**”对话框中，指定一个或多个网站用户，为他们设置相应的权限，选择性地设置其他选项，然后选择“**共享**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-119">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>
    
9. <span data-ttu-id="4fa20-120">若要向 Office 加载项加载项目录添加加载项，请选择“Office 加载项”****。</span><span class="sxs-lookup"><span data-stu-id="4fa20-120">To add an add-in to the Office Add-ins add-in catalog, choose **Office Add-ins**.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a><span data-ttu-id="4fa20-121">在 Office 365 上设置加载项目录</span><span class="sxs-lookup"><span data-stu-id="4fa20-121">To set up an add-in catalog on Office 365</span></span>

1. <span data-ttu-id="4fa20-122">在“Office 365 管理中心”页上，选择“**管理**”，然后选择“**SharePoint**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-122">On the Office 365 admin center page, choose  **Admin**, and then choose  **SharePoint**.</span></span>
    
2. <span data-ttu-id="4fa20-123">在左侧的任务窗格中，选择“**外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-123">In the left task pane, choose  **add-ins**.</span></span>
    
3. <span data-ttu-id="4fa20-124">在“**外接程序**”页上，选择“**外接程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-124">On the  **add-ins** page, choose **Add-in Catalog**.</span></span>
    
4. <span data-ttu-id="4fa20-125">在“**外接程序目录网站**”页上，选择“**确定**”以接受默认选项，并新建外接程序目录网站。</span><span class="sxs-lookup"><span data-stu-id="4fa20-125">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>
    
5. <span data-ttu-id="4fa20-126">在“**创建外接程序目录网站集**”页上，指定外接程序目录网站的标题。</span><span class="sxs-lookup"><span data-stu-id="4fa20-126">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>
    
6. <span data-ttu-id="4fa20-127">指定网站地址。</span><span class="sxs-lookup"><span data-stu-id="4fa20-127">Specify the web site address.</span></span>
    
7. <span data-ttu-id="4fa20-p103">将“**存储配额**”设置为可能的最低值（当前为 110）。你将仅在该网站集上安装外接程序包，它们非常小。</span><span class="sxs-lookup"><span data-stu-id="4fa20-p103">Set the  **Storage Quota** to the lowest possible value (currently 110). You will only be installing add-in packages on this site collection and they are very small.</span></span>
    
8. <span data-ttu-id="4fa20-p104">将“**服务器资源配额**”设置为 0（零）。（服务器资源配额与限制性能不佳的沙盒解决方案有关，但你不会在外接程序目录网站上安装任何沙盒解决方案。）</span><span class="sxs-lookup"><span data-stu-id="4fa20-p104">Set the  **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>
    
9. <span data-ttu-id="4fa20-132">选择“确定”****。</span><span class="sxs-lookup"><span data-stu-id="4fa20-132">Choose  **OK**.</span></span>
    
10. <span data-ttu-id="4fa20-p105">若要将加载项添加到加载项目录网站，请转到刚刚创建的网站。在左侧导航窗格中，依次选择“Office 加载项”**** 和“新加载项”****，以上传 Office 加载项清单文件。</span><span class="sxs-lookup"><span data-stu-id="4fa20-p105">To add an add-in to the Add-in Catalog Site, browse to the site you have just created. In the left navigation pane, choose  **Office Add-ins**, and then, to upload an Office Add-in manifest file, choose  **new add-in**.</span></span>

## <a name="publish-an-add-in-to-an-add-in-catalog"></a><span data-ttu-id="4fa20-135">将加载项发布到加载项目录</span><span class="sxs-lookup"><span data-stu-id="4fa20-135">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="4fa20-136">若要将加载项发布到加载项目录，请完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="4fa20-136">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="4fa20-137">转到加载项目录：</span><span class="sxs-lookup"><span data-stu-id="4fa20-137">Browse to the add-in catalog:</span></span>

    - <span data-ttu-id="4fa20-138">打开 SharePoint 管理中心主页。</span><span class="sxs-lookup"><span data-stu-id="4fa20-138">Open the SharePoint Central Administration main page.</span></span>
    
    - <span data-ttu-id="4fa20-139">选择“加载项”****。</span><span class="sxs-lookup"><span data-stu-id="4fa20-139">Select  **Add-ins**.</span></span>
    
    - <span data-ttu-id="4fa20-140">选择“管理加载项目录”****。</span><span class="sxs-lookup"><span data-stu-id="4fa20-140">Select  **Manage Add-in Catalog**.</span></span>
    
    - <span data-ttu-id="4fa20-141">依次选择所提供的链接和左侧导航栏上的“Office 加载项”****。</span><span class="sxs-lookup"><span data-stu-id="4fa20-141">Choose the link provided, and then choose  **Office Add-ins** on the left navigation bar.</span></span>
    
2. <span data-ttu-id="4fa20-142">选择“单击添加新项”**** 链接。</span><span class="sxs-lookup"><span data-stu-id="4fa20-142">Choose the  **Click to add new item** link.</span></span>
    
3. <span data-ttu-id="4fa20-143">选择“浏览”****，再指定要上传的[清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="4fa20-143">Choose  **Browse**, and then specify the [manifest](../develop/add-in-manifests.md) to upload.</span></span>
    
    <span data-ttu-id="4fa20-p106">此目录中的内容和任务窗格外接程序现在可从“**Office 外接程序**”对话框提供。若要访问这些外接程序，请在“**插入**”选项卡上选择“**我的外接程序**”，然后选择“**我的组织**”。</span><span class="sxs-lookup"><span data-stu-id="4fa20-p106">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="4fa20-146">加载项目录的最终用户体验</span><span class="sxs-lookup"><span data-stu-id="4fa20-146">End user experience with the add-in catalog</span></span>

<span data-ttu-id="4fa20-147">最终用户可以通过完成以下步骤来访问 Office 应用程序中的加载项目录：</span><span class="sxs-lookup"><span data-stu-id="4fa20-147">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="4fa20-148">在 Office 应用程序中，转到“文件”**** > “选项”****“信任中心” > **** > 信任中心设置**** > “受信任的加载项目录”****。</span><span class="sxs-lookup"><span data-stu-id="4fa20-148">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
    
2. <span data-ttu-id="4fa20-149">指定加载项目录的_父级 SharePoint 网站集_的 URL。</span><span class="sxs-lookup"><span data-stu-id="4fa20-149">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 
    
    <span data-ttu-id="4fa20-150">例如，如果 Office 加载项目录的 URL 是：</span><span class="sxs-lookup"><span data-stu-id="4fa20-150">For example, if the URL of the Office Add-ins catalog is:</span></span>
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    <span data-ttu-id="4fa20-151">仅指定父网站集的 URL：</span><span class="sxs-lookup"><span data-stu-id="4fa20-151">Specify just the URL of the parent site collection:</span></span>
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. <span data-ttu-id="4fa20-p107">|||UNTRANSLATED_CONTENT_START|||Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="4fa20-p107">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="4fa20-154">或者，管理员可以使用组策略在 SharePoint 上指定 Office 加载项目录。</span><span class="sxs-lookup"><span data-stu-id="4fa20-154">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="4fa20-155">有关详细信息，请参阅[使用组策略管理用户安装和使用 Office 加载项的方式](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office)一节。</span><span class="sxs-lookup"><span data-stu-id="4fa20-155">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office) on TechNet.</span></span>
