---
title: 将任务窗格和内容加载项发布到 SharePoint 应用程序目录
description: 为使组织内的用户可访问 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的应用程序目录中。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 827c11c5c8666bc1478e36bb9568c536a61c1f63
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408794"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a><span data-ttu-id="8cecd-103">将任务窗格和内容加载项发布到 SharePoint 应用程序目录</span><span class="sxs-lookup"><span data-stu-id="8cecd-103">Publish task pane and content add-ins to a SharePoint app catalog</span></span>

<span data-ttu-id="8cecd-p101">应用程序目录是 SharePoint Web 应用程序或 SharePoint Online 租户中的专用网站集，用于托管 Office 和 SharePoint 加载项的文档库。若要向组织用户分发 Office 加载项，管理员可以将 Office 加载项清单文件上传到组织的应用程序目录。如果管理员将应用程序目录注册为受信任的目录，用户就可以通过 Office 客户端应用程序中的插入 UI 插入加载项。</span><span class="sxs-lookup"><span data-stu-id="8cecd-p101">An app catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the app catalog for their organization. When an administrator registers an app catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="8cecd-106">SharePoint 上的应用程序目录不支持在[加载项清单](../develop/add-in-manifests.md)的 `VersionOverrides` 节点中实现的加载项功能（如加载项命令）。</span><span class="sxs-lookup"><span data-stu-id="8cecd-106">App catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="8cecd-107">如果您针对的是云或混合环境，我们建议 [通过 Microsoft 365 管理中心使用集中部署](../publish/centralized-deployment.md) 来发布你的外接程序。</span><span class="sxs-lookup"><span data-stu-id="8cecd-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Microsoft 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="8cecd-108">Mac 版 Office 不支持 SharePoint 上的应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="8cecd-108">App catalogs on SharePoint are not supported in Office on Mac.</span></span> <span data-ttu-id="8cecd-109">若要向 Mac 客户端部署 Office 加载项，必须将其提交到 [AppSource](/office/dev/store/submit-to-the-office-store)。</span><span class="sxs-lookup"><span data-stu-id="8cecd-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="create-an-app-catalog"></a><span data-ttu-id="8cecd-110">创建应用程序目录</span><span class="sxs-lookup"><span data-stu-id="8cecd-110">Create an app catalog</span></span>

<span data-ttu-id="8cecd-111">完成以下某个部分中的步骤，以使用本地 SharePoint Server 或 Office 365 创建应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="8cecd-111">Complete the steps in one of the following sections to create an app catalog with on-premises SharePoint Server or on Office 365.</span></span>

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a><span data-ttu-id="8cecd-112">为本地 SharePoint Server 创建应用程序目录</span><span class="sxs-lookup"><span data-stu-id="8cecd-112">To create an app catalog for on-premises SharePoint Server</span></span>

<span data-ttu-id="8cecd-113">若要创建 SharePoint 应用程序目录，请按照[配置 Web 应用程序的应用程序目录网站](/sharepoint/administration/manage-the-app-catalog)中的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="8cecd-113">To create the SharePoint app catalog, follow the instructions at [Configure the App Catalog site for a web application](/sharepoint/administration/manage-the-app-catalog).</span></span>

<span data-ttu-id="8cecd-114">创建应用程序目录后，请按照相关步骤[发布 Office 加载项](#publish-an-office-add-in)。</span><span class="sxs-lookup"><span data-stu-id="8cecd-114">Once you have created the app catalog follow the steps to [publish an Office Add-in](#publish-an-office-add-in).</span></span>

### <a name="to-create-an-app-catalog-on-microsoft-365"></a><span data-ttu-id="8cecd-115">在 Microsoft 365 上创建应用程序目录</span><span class="sxs-lookup"><span data-stu-id="8cecd-115">To create an app catalog on Microsoft 365</span></span>

<span data-ttu-id="8cecd-116">若要创建 SharePoint 应用程序目录，请按照 [create The App catalog site collection](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection)中的说明操作。</span><span class="sxs-lookup"><span data-stu-id="8cecd-116">To create the SharePoint app catalog, follow the instructions at [Create the App Catalog site collection](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection).</span></span> <span data-ttu-id="8cecd-117">创建应用程序目录后，请按照下一节中的步骤发布 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="8cecd-117">Once you have created the app catalog, follow the steps in the next section to publish an Office Add-in.</span></span>

## <a name="publish-an-office-add-in"></a><span data-ttu-id="8cecd-118">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="8cecd-118">Publish an Office Add-in</span></span>

<span data-ttu-id="8cecd-119">完成以下各节之一中的步骤，将 Office 外接程序发布到 Microsoft 365 或本地 SharePoint Server 上的应用程序目录。</span><span class="sxs-lookup"><span data-stu-id="8cecd-119">Complete the steps in one of the following sections to publish an Office Add-in to an app catalog on Microsoft 365 or on-premises SharePoint Server.</span></span>

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-microsoft-365"></a><span data-ttu-id="8cecd-120">将 Office 外接程序发布到 Microsoft 365 上的 SharePoint 应用程序目录</span><span class="sxs-lookup"><span data-stu-id="8cecd-120">To publish an Office add-in to a SharePoint app catalog on Microsoft 365</span></span>

1. <span data-ttu-id="8cecd-121">转到[新的 SharePoint 管理中心的“活动站点”页面](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true)，然后使用在组织中具有[管理员权限](/sharepoint/sharepoint-admin-role)的帐户进行登录。</span><span class="sxs-lookup"><span data-stu-id="8cecd-121">Go to the [Active sites page of the new SharePoint admin center](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true) and sign in with an account that has [admin permissions](/sharepoint/sharepoint-admin-role) for your organization.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8cecd-122">如果你有 Microsoft 365 德国，请 [登录 microsoft 365 管理中心](https://go.microsoft.com/fwlink/p/?linkid=848041)，然后浏览到 SharePoint 管理中心并打开 "更多功能" 页面。</span><span class="sxs-lookup"><span data-stu-id="8cecd-122">If you have Microsoft 365 Germany, [sign in to the Microsoft 365 admin center](https://go.microsoft.com/fwlink/p/?linkid=848041), then browse to the SharePoint admin center and open the More features page.</span></span> <br><span data-ttu-id="8cecd-123">如果你有由世纪互联运营的 Microsoft 365 (中国) ，请 [登录到 Microsoft 365 管理中心](https://go.microsoft.com/fwlink/p/?linkid=850627)，然后浏览到 SharePoint 管理中心并打开 "更多功能" 页面。</span><span class="sxs-lookup"><span data-stu-id="8cecd-123">If you have Microsoft 365 operated by 21Vianet (China), [sign in to the Microsoft 365 admin center](https://go.microsoft.com/fwlink/p/?linkid=850627), then browse to the SharePoint admin center and open the More features page.</span></span>

1. <span data-ttu-id="8cecd-124">通过在 "URL" 列中选择应用程序目录网站的 URL 来打开该网站。</span><span class="sxs-lookup"><span data-stu-id="8cecd-124">Open the app catalog site by selecting its URL in the URL column.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8cecd-125">如果你刚刚在上一节中创建了应用程序目录网站，可能需要几分钟时间才能完成网站设置。</span><span class="sxs-lookup"><span data-stu-id="8cecd-125">If you just created the app catalog site in the previous section, it can take a few minutes for the site to finish setting up.</span></span>

1. <span data-ttu-id="8cecd-126">选择“**分发 Office 应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-126">Choose **Distribute apps for Office**.</span></span>
1. <span data-ttu-id="8cecd-127">在“**Office 应用程序**”页中，选择“**新建**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-127">In the **Apps for Office** page, choose **New**.</span></span>
1. <span data-ttu-id="8cecd-128">在“**添加文档**”对话框中，选择“**选择文件**”按钮。</span><span class="sxs-lookup"><span data-stu-id="8cecd-128">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
1. <span data-ttu-id="8cecd-129">找到并指定要上传的“[清单文件](../develop/add-in-manifests.md)”，并选择“**打开**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-129">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
1. <span data-ttu-id="8cecd-130">在“**添加文档**”对话框中，选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-130">In the **Add a document** dialog, choose **OK**.</span></span>

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a><span data-ttu-id="8cecd-131">使用本地 SharePoint Server 将加载项发布到应用程序目录</span><span class="sxs-lookup"><span data-stu-id="8cecd-131">To publish an add-in to an app catalog with on-premises SharePoint Server</span></span>

1. <span data-ttu-id="8cecd-132">打开“**管理中心**”页。</span><span class="sxs-lookup"><span data-stu-id="8cecd-132">Open the **Central Administration** page.</span></span>
1. <span data-ttu-id="8cecd-133">在左侧的任务窗格中，选择“**应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-133">In the left task pane, choose **Apps**.</span></span>
1. <span data-ttu-id="8cecd-134">在“**应用程序**”页的“**应用程序管理**”下方，选择“**管理应用程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-134">On the **Apps** page, under **App Management**, choose **Manage App Catalog**.</span></span>
1. <span data-ttu-id="8cecd-135">在“**管理应用程序目录**”页上，确保在“**Web 应用程序**”选择器中选择了正确的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="8cecd-135">On the **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application** Selector.</span></span>
1. <span data-ttu-id="8cecd-136">选择“**网站 URL**”下的 URL 以打开应用程序目录网站。</span><span class="sxs-lookup"><span data-stu-id="8cecd-136">Choose the URL under the **Site URL** to open the app catalog site.</span></span>
1. <span data-ttu-id="8cecd-137">选择“**分发 Office 应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-137">Choose **Distribute apps for Office**.</span></span>
1. <span data-ttu-id="8cecd-138">在“**Office 应用程序**”页中，选择“**新建**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-138">In the **Apps for Office** page, choose **New**.</span></span>
1. <span data-ttu-id="8cecd-139">在“**添加文档**”对话框中，选择“**选择文件**”按钮。</span><span class="sxs-lookup"><span data-stu-id="8cecd-139">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
1. <span data-ttu-id="8cecd-140">找到并指定要上传的“[清单文件](../develop/add-in-manifests.md)”，并选择“**打开**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-140">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
1. <span data-ttu-id="8cecd-141">在“**添加文档**”对话框中，选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-141">In the **Add a document** dialog, choose **OK**.</span></span>

## <a name="insert-office-add-ins-from-the-app-catalog"></a><span data-ttu-id="8cecd-142">从应用程序目录插入 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="8cecd-142">Insert Office Add-ins from the app catalog</span></span>

<span data-ttu-id="8cecd-143">对于联机 Office 应用程序，你可以通过完成以下步骤从应用程序目录中找到 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8cecd-143">For online Office applications, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="8cecd-144">打开联机 Office 应用程序（Excel、PowerPoint 或 Word）。</span><span class="sxs-lookup"><span data-stu-id="8cecd-144">Open the online Office application (Excel, PowerPoint, or Word).</span></span>
1. <span data-ttu-id="8cecd-145">创建或打开文档。</span><span class="sxs-lookup"><span data-stu-id="8cecd-145">Create or open a document.</span></span>
1. <span data-ttu-id="8cecd-146">选择“**插入**” > “**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-146">Choose **Insert** > **Add-ins**.</span></span>
1. <span data-ttu-id="8cecd-147">在“Office 加载项”对话框中，选择“**我的组织**”选项卡。此时将列出 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8cecd-147">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
1. <span data-ttu-id="8cecd-148">选择 Office 加载项，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-148">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="8cecd-149">对于桌面上的 Office 应用程序，你可以通过完成以下步骤从应用程序目录中找到 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8cecd-149">For Office applications on the desktop, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="8cecd-150">打开桌面版 Office 应用程序（Excel、Word 或 PowerPoint）</span><span class="sxs-lookup"><span data-stu-id="8cecd-150">Open the desktop Office application (Excel, Word, or PowerPoint)</span></span>
1. <span data-ttu-id="8cecd-151">选择“**文件**” > “**选项**” > “**信任中心**” > “**信任中心设置**” > “**受信任的加载项目录**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-151">Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
1. <span data-ttu-id="8cecd-152">在“**目录 URL**”框中输入 SharePoint 应用程序目录的 URL，然后选择“**添加目录**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-152">Enter the URL of the SharePoint app catalog in the **Catalog Url** box and choose **Add catalog**.</span></span>
    <span data-ttu-id="8cecd-153">使用较短形式的 URL。</span><span class="sxs-lookup"><span data-stu-id="8cecd-153">Use the shorter form of the URL.</span></span> <span data-ttu-id="8cecd-154">例如，如果 SharePoint 应用程序目录的 URL 为：</span><span class="sxs-lookup"><span data-stu-id="8cecd-154">For example, if the URL of the SharePoint app catalog is:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`

    <span data-ttu-id="8cecd-155">仅指定父网站集的 URL：</span><span class="sxs-lookup"><span data-stu-id="8cecd-155">Specify just the URL of the parent site collection:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
1. <span data-ttu-id="8cecd-156">关闭并重新打开 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="8cecd-156">Close and reopen the Office application.</span></span>
1. <span data-ttu-id="8cecd-157">选择“**插入**” > “**获取加载项**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-157">Choose **Insert** > **Get Add-ins**.</span></span>
1. <span data-ttu-id="8cecd-158">在“Office 加载项”对话框中，选择“**我的组织**”选项卡。此时将列出 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8cecd-158">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
1. <span data-ttu-id="8cecd-159">选择 Office 加载项，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="8cecd-159">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="8cecd-160">或者，管理员可以使用组策略在 SharePoint 上指定应用目录。</span><span class="sxs-lookup"><span data-stu-id="8cecd-160">Alternatively, an administrator can specify an app catalog on SharePoint by using Group Policy.</span></span> <span data-ttu-id="8cecd-161">可在 [管理模板文件 (ADMX/ADML) For Microsoft 365 Apps、office 2019 和 office 2016](https://www.microsoft.com/download/details.aspx?id=49030) 中找到相关的策略设置，并可在 **User Configuration\Policies\Administrative \microsoft Office 2016 \ Security 设置 \ 信任中心目录**中找到这些设置。</span><span class="sxs-lookup"><span data-stu-id="8cecd-161">The relevant policy settings are available in the [Administrative Template files (ADMX/ADML) for Microsoft 365 Apps, Office 2019, and Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) and be found under **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs**.</span></span>
