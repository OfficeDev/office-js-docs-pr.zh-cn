---
title: 将任务窗格和内容加载项发布到 SharePoint 目录
description: 为使组织内的用户可访问 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的加载项目录中。
ms.date: 05/22/2019
localization_priority: Priority
ms.openlocfilehash: bffbf3e83a2e6d8d0c63252c27ba54826611f78b
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432241"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="8f308-103">将任务窗格和内容加载项发布到 SharePoint 目录</span><span class="sxs-lookup"><span data-stu-id="8f308-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="8f308-p101">加载项目录是 SharePoint Web 应用或 SharePoint Online 租赁中的专用网站集，用于托管 Office 和 SharePoint 加载项的文档库。若要向组织用户分发 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的加载项目录。如果管理员将加载项目录注册为受信任的目录，用户就可以通过 Office 客户端应用中的插入 UI 插入加载项。</span><span class="sxs-lookup"><span data-stu-id="8f308-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="8f308-106">SharePoint 上的加载项目录不支持在[加载项清单](../develop/add-in-manifests.md)的 `VersionOverrides` 节点中实现的加载项功能（如加载项命令）。</span><span class="sxs-lookup"><span data-stu-id="8f308-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="8f308-107">如果面向的是云或混合环境，建议通过 [Office 365 管理中心使用集中部署](../publish/centralized-deployment.md)来发布加载项。</span><span class="sxs-lookup"><span data-stu-id="8f308-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="8f308-108">SharePoint 目录不支持 Office for Mac。</span><span class="sxs-lookup"><span data-stu-id="8f308-108">SharePoint catalogs are not supported for Office for Mac.</span></span> <span data-ttu-id="8f308-109">若要向 Mac 客户端部署 Office 加载项，必须将其提交到 [AppSource](/office/dev/store/submit-to-the-office-store)。</span><span class="sxs-lookup"><span data-stu-id="8f308-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="create-an-add-in-catalog"></a><span data-ttu-id="8f308-110">创建加载项目录</span><span class="sxs-lookup"><span data-stu-id="8f308-110">Create an add-in catalog</span></span>

<span data-ttu-id="8f308-111">完成以下部分之一中的步骤，以在 SharePoint 或 Office 365 上设置加载项目录。</span><span class="sxs-lookup"><span data-stu-id="8f308-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-create-an-add-in-catalog-for-on-premises-sharepoint"></a><span data-ttu-id="8f308-112">为本地 SharePoint 创建加载项目录</span><span class="sxs-lookup"><span data-stu-id="8f308-112">To set up an add-in catalog for on-premises SharePoint</span></span>

> [!NOTE]
> <span data-ttu-id="8f308-113">本地 SharePoint 中的 UI 仍将加载项称为**应用程序**。</span><span class="sxs-lookup"><span data-stu-id="8f308-113">The UI in on-premises SharePoint still refers to add-ins as **apps**.</span></span>

1. <span data-ttu-id="8f308-114">浏览到**管理中心网站**。</span><span class="sxs-lookup"><span data-stu-id="8f308-114">Browse to the  **Central Administration Site**.</span></span>

2. <span data-ttu-id="8f308-115">在左侧的任务窗格中，选择“**应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-115">In the left task pane, choose  **Apps**.</span></span>

3. <span data-ttu-id="8f308-116">在“**应用程序**”页的“**应用程序管理**”下方，选择“**管理应用程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-116">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>

4. <span data-ttu-id="8f308-117">在“**管理应用程序目录**”页上，确保在“**Web 应用程序选择器**”中选择了正确的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="8f308-117">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>

5. <span data-ttu-id="8f308-118">选择“**查看网站设置**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-118">Choose  **View site settings**.</span></span>

6. <span data-ttu-id="8f308-119">在“**网站设置**”页上选择“**网站集管理员**”以指定网站集管理员，然后选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-119">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>

7. <span data-ttu-id="8f308-120">若要向用户授予网站权限，请选择“**网站权限**”，然后选择“**授予权限**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-120">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>

8. <span data-ttu-id="8f308-121">在“**共享‘应用程序目录网站’**”对话框中，指定一个或多个网站用户，为他们设置相应的权限，选择性地设置其他选项，然后选择“**共享**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-121">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>

9. <span data-ttu-id="8f308-122">若要向 Office 加载项加载项目录添加加载项，请选择“**针对 Office 的应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-122">To add an add-in to the Office Add-ins add-in catalog, choose **Apps for Office**.</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="8f308-123">在 Office 365 上创建应用目录</span><span class="sxs-lookup"><span data-stu-id="8f308-123">To create an app catalog on Office 365</span></span>

<span data-ttu-id="8f308-124">尽管 SharePoint 会将此目录命名为“应用”目录，但你仍可以在应用目录中注册 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8f308-124">Even though SharePoint names the catalog an "app" catalog, you can register Office Add-ins in the app catalog.</span></span>

1. <span data-ttu-id="8f308-125">转到 Microsoft 365 管理中心。</span><span class="sxs-lookup"><span data-stu-id="8f308-125">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="8f308-126">有关如何查找管理中心的信息，请参阅[关于 Microsoft 365 管理中心](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="8f308-126">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="8f308-127">在 Microsoft 365 管理中心页面上，展开“**管理中心**”列表，然后选择“**SharePoint**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-127">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8f308-128">需要使用经典 SharePoint 管理中心才能创建目录。</span><span class="sxs-lookup"><span data-stu-id="8f308-128">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="8f308-129">如果位于新的 SharePoint 管理中心，请在左侧窗格中选择“**经典 SharePoint 管理中心**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-129">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="8f308-130">在左侧的任务窗格中，选择“**应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-130">In the left task pane, choose  **Apps**.</span></span>

4. <span data-ttu-id="8f308-131">在“**应用程序**”页面上，选择“**应用程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-131">On the **apps** page, select **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="8f308-132">如果已创建应用程序目录并且它显示在此页面上，则你可以跳过其余步骤并转至本文下一章节，将你的加载项分步至目录。</span><span class="sxs-lookup"><span data-stu-id="8f308-132">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="8f308-133">在“**应用程序目录网站**”页上，选择“**确定**”以接受默认选项并创建新的加载项目录网站。</span><span class="sxs-lookup"><span data-stu-id="8f308-133">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>

6. <span data-ttu-id="8f308-134">在“**创建应用程序目录网站集**”页上，指定应用程序目录网站的标题。</span><span class="sxs-lookup"><span data-stu-id="8f308-134">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

7. <span data-ttu-id="8f308-135">指定**网站地址**。</span><span class="sxs-lookup"><span data-stu-id="8f308-135">Specify the web site address.</span></span>

8. <span data-ttu-id="8f308-136">指定**管理员**。</span><span class="sxs-lookup"><span data-stu-id="8f308-136">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="8f308-137">将**服务器资源配额**设为 0（零）。</span><span class="sxs-lookup"><span data-stu-id="8f308-137">Set the  **Server Resource Quota** to 0 (zero).</span></span> <span data-ttu-id="8f308-138">（服务器资源配额与限制性能不佳的沙盒解决方案有关，但你不会在应用程序目录网站上安装任何沙盒解决方案。）</span><span class="sxs-lookup"><span data-stu-id="8f308-138">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

10. <span data-ttu-id="8f308-139">选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-139">Choose **OK**.</span></span>

<span data-ttu-id="8f308-140">现在已创建应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="8f308-140">The app catalog is now created.</span></span>

## <a name="publish-an-add-in-to-an-app-catalog"></a><span data-ttu-id="8f308-141">将加载项发布到应用程序目录</span><span class="sxs-lookup"><span data-stu-id="8f308-141">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="8f308-142">若要将加载项发布到现有应用程序目录中，请完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="8f308-142">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="8f308-143">转到 Microsoft 365 管理中心。</span><span class="sxs-lookup"><span data-stu-id="8f308-143">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="8f308-144">有关如何查找管理中心的信息，请参阅[关于 Microsoft 365 管理中心](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="8f308-144">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="8f308-145">在 Microsoft 365 管理中心页面上，展开“**管理中心**”列表，然后选择“**SharePoint**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-145">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="8f308-146">需要使用经典 SharePoint 管理中心才能创建目录。</span><span class="sxs-lookup"><span data-stu-id="8f308-146">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="8f308-147">如果位于新的 SharePoint 管理中心，请在左侧窗格中选择“**经典 SharePoint 管理中心**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-147">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="8f308-148">在左侧的任务窗格中，选择“**应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-148">In the left task pane, choose  **Apps**.</span></span>
4. <span data-ttu-id="8f308-149">在“**应用程序**”页面上，选择“**应用程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-149">On the **apps** page, select **App Catalog**.</span></span>
5. <span data-ttu-id="8f308-150">选择“**分发 Office 应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-150">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="8f308-151">在“**Office 应用程序**”页中，选择“**新建**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-151">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="8f308-152">在“**添加文档**”对话框中，选择“**选择文件**”按钮。</span><span class="sxs-lookup"><span data-stu-id="8f308-152">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="8f308-153">找到并指定要上传的“[清单文件](../develop/add-in-manifests.md)”，并选择“**打开**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-153">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="8f308-154">在“**添加文档**”对话框中，选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-154">In the **Add a document** dialog box, choose **OK**.</span></span>

    <span data-ttu-id="8f308-p108">此目录中的内容和任务窗格外接程序现在可从“**Office 外接程序**”对话框提供。若要访问这些外接程序，请在“**插入**”选项卡上选择“**我的外接程序**”，然后选择“**我的组织**”。</span><span class="sxs-lookup"><span data-stu-id="8f308-p108">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="8f308-157">加载项目录的最终用户体验</span><span class="sxs-lookup"><span data-stu-id="8f308-157">End user experience with the add-in catalog</span></span>

<span data-ttu-id="8f308-158">最终用户可以通过完成以下步骤来访问 Office 应用程序中的加载项目录：</span><span class="sxs-lookup"><span data-stu-id="8f308-158">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="8f308-159">在 Office 应用程序中，转到“文件”\*\*\*\* > “选项”\*\*\*\*“信任中心” > \*\*\*\* > 信任中心设置\*\*\*\* > “受信任的加载项目录”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8f308-159">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>

2. <span data-ttu-id="8f308-160">指定加载项目录的_父级 SharePoint 网站集_的 URL。</span><span class="sxs-lookup"><span data-stu-id="8f308-160">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 

    <span data-ttu-id="8f308-161">例如，如果 Office 加载项目录的 URL 是：</span><span class="sxs-lookup"><span data-stu-id="8f308-161">For example, if the URL of the Office Add-ins catalog is:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    <span data-ttu-id="8f308-162">仅指定父网站集的 URL：</span><span class="sxs-lookup"><span data-stu-id="8f308-162">Specify just the URL of the parent site collection:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. <span data-ttu-id="8f308-p109">关闭并重新打开 Office 应用。此时，加载项目录会出现在“**Office 加载项**”对话框中。</span><span class="sxs-lookup"><span data-stu-id="8f308-p109">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="8f308-165">或者，管理员可以使用组策略在 SharePoint 上指定 Office 加载项目录。</span><span class="sxs-lookup"><span data-stu-id="8f308-165">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="8f308-166">有关详细信息，请参阅[使用组策略管理用户如何安装和使用 Office 加载项](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office)一节。</span><span class="sxs-lookup"><span data-stu-id="8f308-166">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>
