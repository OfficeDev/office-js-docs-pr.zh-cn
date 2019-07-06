---
title: 将任务窗格和内容加载项发布到 SharePoint 应用程序目录
description: 为使组织内的用户可访问 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的应用程序目录中。
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: af1f96615c74065d9a194f4372e69853caa2c6e3
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575301"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a><span data-ttu-id="4491f-103">将任务窗格和内容加载项发布到 SharePoint 应用程序目录</span><span class="sxs-lookup"><span data-stu-id="4491f-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="4491f-p101">应用程序目录是 SharePoint Web 应用程序或 SharePoint Online 租户中的专用网站集，用于托管 Office 和 SharePoint 加载项的文档库。若要向组织用户分发 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的应用程序目录。如果管理员将应用程序目录注册为受信任的目录，用户就可以通过 Office 客户端应用程序中的插入 UI 插入加载项。</span><span class="sxs-lookup"><span data-stu-id="4491f-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="4491f-106">SharePoint 上的应用程序目录不支持在[加载项清单](../develop/add-in-manifests.md)的 `VersionOverrides` 节点中实现的加载项功能（如加载项命令）。</span><span class="sxs-lookup"><span data-stu-id="4491f-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="4491f-107">如果面向的是云或混合环境，建议通过 [Office 365 管理中心使用集中部署](../publish/centralized-deployment.md)来发布加载项。</span><span class="sxs-lookup"><span data-stu-id="4491f-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="4491f-108">Mac 版 Office 不支持 SharePoint 上的应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="4491f-108">App catalogs on SharePoint are not supported in Office on Mac.</span></span> <span data-ttu-id="4491f-109">若要向 Mac 客户端部署 Office 加载项，必须将其提交到 [AppSource](/office/dev/store/submit-to-the-office-store)。</span><span class="sxs-lookup"><span data-stu-id="4491f-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="create-an-app-catalog"></a><span data-ttu-id="4491f-110">创建应用程序目录</span><span class="sxs-lookup"><span data-stu-id="4491f-110">Create app catalog site</span></span>

<span data-ttu-id="4491f-111">完成以下某个部分中的步骤，以使用本地 SharePoint Server 或 Office 365 创建应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="4491f-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a><span data-ttu-id="4491f-112">为本地 SharePoint Server 创建应用程序目录</span><span class="sxs-lookup"><span data-stu-id="4491f-112">To create an app catalog for on-premises SharePoint Server</span></span>

<span data-ttu-id="4491f-113">若要创建 SharePoint 应用程序目录，请按照[配置 Web 应用程序的应用程序目录网站](/sharepoint/administration/manage-the-app-catalog)中的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="4491f-113">To create the SharePoint app catalog, follow the instructions at [Configure the App Catalog site for a web application](/sharepoint/administration/manage-the-app-catalog).</span></span>

<span data-ttu-id="4491f-114">创建应用程序目录后，请按照相关步骤[发布 Office 加载项](#publish-an-office-add-in)。</span><span class="sxs-lookup"><span data-stu-id="4491f-114">Once you have created the app catalog follow the steps to [publish an Office Add-in](#publish-an-office-add-in).</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="4491f-115">在 Office 365 上创建应用程序目录</span><span class="sxs-lookup"><span data-stu-id="4491f-115">To create an app catalog on Office 365</span></span>

1. <span data-ttu-id="4491f-116">转到 Microsoft 365 管理中心。</span><span class="sxs-lookup"><span data-stu-id="4491f-116">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="4491f-117">有关如何查找管理中心的信息，请参阅[关于 Microsoft 365 管理中心](/office365/admin/admin-overview/about-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="4491f-117">For information on how to find the admin center, see [About the Microsoft 365 admin center](/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="4491f-118">在 Microsoft 365 管理中心页面上，展开“**管理中心**”列表，然后选择“**SharePoint**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-118">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="4491f-119">需要使用经典 SharePoint 管理中心才能创建目录。</span><span class="sxs-lookup"><span data-stu-id="4491f-119">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="4491f-120">如果位于新的 SharePoint 管理中心，请在左侧窗格中选择“**经典 SharePoint 管理中心**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-120">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="4491f-121">在左侧的任务窗格中，选择“**应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-121">In the left task pane, choose  **Apps**.</span></span>

4. <span data-ttu-id="4491f-122">在“**应用程序**”页面上，选择“**应用程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-122">On the **apps** page, select **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="4491f-123">如果已创建应用程序目录并且它显示在此页面上，则你可以跳过其余步骤并转至本文下一章节，将你的加载项发布到目录。</span><span class="sxs-lookup"><span data-stu-id="4491f-123">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="4491f-124">在“**应用程序目录网站**”页上，选择“**确定**”以接受默认选项并创建新的应用程序目录网站。</span><span class="sxs-lookup"><span data-stu-id="4491f-124">On the **App Catalog Site** page, select **OK** to accept the default option and create a new app catalog site.</span></span>

6. <span data-ttu-id="4491f-125">在“**创建应用程序目录网站集**”页上，指定应用程序目录网站的标题。</span><span class="sxs-lookup"><span data-stu-id="4491f-125">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

7. <span data-ttu-id="4491f-126">指定**网站地址**。</span><span class="sxs-lookup"><span data-stu-id="4491f-126">Specify the web site address.</span></span>

8. <span data-ttu-id="4491f-127">指定**管理员**。</span><span class="sxs-lookup"><span data-stu-id="4491f-127">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="4491f-128">将**服务器资源配额**设为 0（零）。</span><span class="sxs-lookup"><span data-stu-id="4491f-128">Set the  **Server Resource Quota** to 0 (zero).</span></span> <span data-ttu-id="4491f-129">（服务器资源配额与限制性能不佳的沙盒解决方案有关，但你不会在应用程序目录网站上安装任何沙盒解决方案。）</span><span class="sxs-lookup"><span data-stu-id="4491f-129">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

10. <span data-ttu-id="4491f-130">选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-130">Choose **OK**.</span></span>

## <a name="publish-an-office-add-in"></a><span data-ttu-id="4491f-131">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4491f-131">Publish an Office Add-in</span></span>

<span data-ttu-id="4491f-132">完成以下某个部分中的步骤，以将 Office 加载项发布到 Office 365 或本地 SharePoint Server 上的应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="4491f-132">Complete the steps in one of the following sections to publish an Office Add-in to an app catalog on Office 365 or on-premises SharePoint Server.</span></span>

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a><span data-ttu-id="4491f-133">将 Office 加载项发布到 Office 365 上的 SharePoint 应用程序目录</span><span class="sxs-lookup"><span data-stu-id="4491f-133">To publish an Office add-in to a SharePoint app catalog on Office 365</span></span>

1. <span data-ttu-id="4491f-134">转到 Microsoft 365 管理中心。</span><span class="sxs-lookup"><span data-stu-id="4491f-134">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="4491f-135">有关如何查找管理中心的信息，请参阅[关于 Microsoft 365 管理中心](/office365/admin/admin-overview/about-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="4491f-135">For information on how to find the admin center, see [About the Microsoft 365 admin center](/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="4491f-136">在 Microsoft 365 管理中心页面上，展开“**管理中心**”列表，然后选择“**SharePoint**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-136">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="4491f-137">需要使用经典 SharePoint 管理中心才能创建目录。</span><span class="sxs-lookup"><span data-stu-id="4491f-137">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="4491f-138">如果位于新的 SharePoint 管理中心，请在左侧窗格中选择“**经典 SharePoint 管理中心**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-138">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="4491f-139">在左侧的任务窗格中，选择“**应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-139">In the left task pane, choose  **Apps**.</span></span>
4. <span data-ttu-id="4491f-140">在“**应用程序**”页面上，选择“**应用程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-140">On the **apps** page, select **App Catalog**.</span></span>
5. <span data-ttu-id="4491f-141">选择“**分发 Office 应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-141">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="4491f-142">在“**Office 应用程序**”页中，选择“**新建**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-142">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="4491f-143">在“**添加文档**”对话框中，选择“**选择文件**”按钮。</span><span class="sxs-lookup"><span data-stu-id="4491f-143">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="4491f-144">找到并指定要上传的“[清单文件](../develop/add-in-manifests.md)”，并选择“**打开**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-144">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="4491f-145">在“**添加文档**”对话框中，选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-145">In the **Add a document** dialog box, choose **OK**.</span></span>

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a><span data-ttu-id="4491f-146">使用本地 SharePoint Server 将加载项发布到应用程序目录</span><span class="sxs-lookup"><span data-stu-id="4491f-146">To publish an add-in to an app catalog with on-premises SharePoint Server</span></span>

1. <span data-ttu-id="4491f-147">打开“**管理中心**”页。</span><span class="sxs-lookup"><span data-stu-id="4491f-147">Open the SharePoint Central Administration main page.</span></span>
2. <span data-ttu-id="4491f-148">在左侧的任务窗格中，选择“**应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-148">In the left task pane, choose  **Apps**.</span></span>
3. <span data-ttu-id="4491f-149">在“**应用程序**”页的“**应用程序管理**”下方，选择“**管理应用程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-149">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>
4. <span data-ttu-id="4491f-150">在“**管理应用程序目录**”页上，确保在“**Web 应用程序**”选择器中选择了正确的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="4491f-150">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>
5. <span data-ttu-id="4491f-151">选择“**网站 URL**”下的 URL 以打开应用程序目录网站。</span><span class="sxs-lookup"><span data-stu-id="4491f-151">Choose the URL under the **Site URL** to open the app catalog site.</span></span>
6. <span data-ttu-id="4491f-152">选择“**分发 Office 应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-152">Choose **Distribute apps for Office**.</span></span>
7. <span data-ttu-id="4491f-153">在“**Office 应用程序**”页中，选择“**新建**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-153">In the **Apps for Office** page, choose **New**.</span></span>
8. <span data-ttu-id="4491f-154">在“**添加文档**”对话框中，选择“**选择文件**”按钮。</span><span class="sxs-lookup"><span data-stu-id="4491f-154">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
9. <span data-ttu-id="4491f-155">找到并指定要上传的“[清单文件](../develop/add-in-manifests.md)”，并选择“**打开**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-155">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
10. <span data-ttu-id="4491f-156">在“**添加文档**”对话框中，选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-156">In the **Add a document** dialog box, choose **OK**.</span></span>

## <a name="insert-office-add-ins-from-the-app-catalog"></a><span data-ttu-id="4491f-157">从应用程序目录插入 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="4491f-157">Insert Office Add-ins from the app catalog</span></span>

<span data-ttu-id="4491f-158">对于联机 Office 应用程序，你可以通过完成以下步骤从应用程序目录中找到 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4491f-158">For online Office applications, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="4491f-159">打开联机 Office 应用程序（Excel、PowerPoint 或 Word）。</span><span class="sxs-lookup"><span data-stu-id="4491f-159">Open the online Office application (Excel, PowerPoint, or Word).</span></span>
2. <span data-ttu-id="4491f-160">创建或打开文档。</span><span class="sxs-lookup"><span data-stu-id="4491f-160">Create or open a document.</span></span>
3. <span data-ttu-id="4491f-161">选择“**插入**” > “**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-161">Choose **Insert** > **Add-ins**.</span></span>
4. <span data-ttu-id="4491f-162">在“Office 加载项”对话框中，选择“**我的组织**”选项卡。此时将列出 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4491f-162">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
5. <span data-ttu-id="4491f-163">选择 Office 加载项，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-163">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="4491f-164">对于桌面上的 Office 应用程序，你可以通过完成以下步骤从应用程序目录中找到 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4491f-164">For Office applications on the desktop, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="4491f-165">打开桌面版 Office 应用程序（Excel、Word 或 PowerPoint）</span><span class="sxs-lookup"><span data-stu-id="4491f-165">Open the desktop Office application (Excel, Word, or PowerPoint)</span></span>
2. <span data-ttu-id="4491f-166">选择“**文件**” > “**选项**” > “**信任中心**” > “**信任中心设置**” > “**受信任的加载项目录**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-166">Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
3. <span data-ttu-id="4491f-167">在“**目录 URL**”框中输入 SharePoint 应用程序目录的 URL，然后选择“**添加目录**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-167">Enter the URL of the SharePoint app catalog in the **Catalog Url** box and choose **Add catalog**.</span></span>
    <span data-ttu-id="4491f-168">使用较短形式的 URL。</span><span class="sxs-lookup"><span data-stu-id="4491f-168">Use the shorter form of the URL.</span></span> <span data-ttu-id="4491f-169">例如，如果 SharePoint 应用程序目录的 URL 为：</span><span class="sxs-lookup"><span data-stu-id="4491f-169">For example, if the URL of the Office Add-ins catalog is:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    
    <span data-ttu-id="4491f-170">仅指定父网站集的 URL：</span><span class="sxs-lookup"><span data-stu-id="4491f-170">Specify just the URL of the parent site collection:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. <span data-ttu-id="4491f-171">关闭并重新打开 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="4491f-171">Close and reopen the Office application.</span></span> 
5. <span data-ttu-id="4491f-172">选择“**插入**” > “**获取加载项**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-172">Choose **Insert** > **Get Add-ins**.</span></span>
4. <span data-ttu-id="4491f-173">在“Office 加载项”对话框中，选择“**我的组织**”选项卡。此时将列出 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4491f-173">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
5. <span data-ttu-id="4491f-174">选择 Office 加载项，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="4491f-174">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="4491f-175">或者，管理员可以使用组策略在 SharePoint 上指定应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="4491f-175">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="4491f-176">有关详细信息，请参阅[使用组策略管理用户如何安装和使用 Office 加载项](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office)一节。</span><span class="sxs-lookup"><span data-stu-id="4491f-176">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>
