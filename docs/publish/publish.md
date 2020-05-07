---
title: 部署和发布 Office 加载项
description: 部署 Office 加载项以进行测试或分发给用户的方法和选项。
ms.date: 05/06/2020
localization_priority: Priority
ms.openlocfilehash: b4ee2dab2fec0237ad06776861678c4c23da6ca8
ms.sourcegitcommit: 1e3d7051ba7c8c0ccfd6607099d90469fc8141e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2020
ms.locfileid: "44066086"
---
# <a name="deploy-and-publish-office-add-ins"></a><span data-ttu-id="e5a45-103">部署和发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="e5a45-103">Deploy and publish Office Add-ins</span></span>

<span data-ttu-id="e5a45-104">可以使用几种方法之一来部署 Office 外接程序，以用于对用户进行测试或分发：</span><span class="sxs-lookup"><span data-stu-id="e5a45-104">You can use one of several methods to deploy your Office Add-in for testing or distribution to users.</span></span>

|<span data-ttu-id="e5a45-105">**方法**</span><span class="sxs-lookup"><span data-stu-id="e5a45-105">**Method**</span></span>|<span data-ttu-id="e5a45-106">**Use...**</span><span class="sxs-lookup"><span data-stu-id="e5a45-106">**Use...**</span></span>|
|:---------|:------------|
|[<span data-ttu-id="e5a45-107">旁加载</span><span class="sxs-lookup"><span data-stu-id="e5a45-107">Sideloading</span></span>](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|<span data-ttu-id="e5a45-108">在开发过程中测试在 Windows、iPad、Mac 或浏览器中运行的加载项。</span><span class="sxs-lookup"><span data-stu-id="e5a45-108">As part of your development process, to test your add-in running on Windows, iPad, Mac, or in a browser.</span></span>|
|[<span data-ttu-id="e5a45-109">集中部署</span><span class="sxs-lookup"><span data-stu-id="e5a45-109">Centralized Deployment</span></span>](centralized-deployment.md)|<span data-ttu-id="e5a45-110">在云部署中，使用 Office 365 管理中心将加载项分发给组织中的用户。</span><span class="sxs-lookup"><span data-stu-id="e5a45-110">In a cloud deployment, to distribute your add-in to users in your organization by using the Office 365 admin center.</span></span>|
|[<span data-ttu-id="e5a45-111">SharePoint 目录</span><span class="sxs-lookup"><span data-stu-id="e5a45-111">SharePoint catalog</span></span>](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|<span data-ttu-id="e5a45-112">在本地环境中，用于向组织用户分发加载项。</span><span class="sxs-lookup"><span data-stu-id="e5a45-112">In an on-premises environment, to distribute your add-in to users in your organization.</span></span>|
|[<span data-ttu-id="e5a45-113">AppSource</span><span class="sxs-lookup"><span data-stu-id="e5a45-113">AppSource</span></span>](/office/dev/store/submit-to-appsource-via-partner-center)|<span data-ttu-id="e5a45-114">用于向用户公开分发加载项。</span><span class="sxs-lookup"><span data-stu-id="e5a45-114">To distribute your add-in publicly to users.</span></span>|
|[<span data-ttu-id="e5a45-115">Exchange 服务器</span><span class="sxs-lookup"><span data-stu-id="e5a45-115">Exchange server</span></span>](#outlook-add-in-deployment)|<span data-ttu-id="e5a45-116">在本地或在线环境中，用于向用户分发 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="e5a45-116">In an on-premises or online environment, to distribute Outlook add-ins to users.</span></span>|
|[<span data-ttu-id="e5a45-117">网络共享</span><span class="sxs-lookup"><span data-stu-id="e5a45-117">Network share</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|<span data-ttu-id="e5a45-118">在网络上的 Windows 计算机中（要在其中托管加载项），转到要用作共享文件夹目录的文件夹的父文件夹或驱动器号。</span><span class="sxs-lookup"><span data-stu-id="e5a45-118">On a Windows computer on a network where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>|

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="deployment-options-by-office-host"></a><span data-ttu-id="e5a45-119">Office 主机提供的部署选项</span><span class="sxs-lookup"><span data-stu-id="e5a45-119">Deployment options by Office host</span></span>

<span data-ttu-id="e5a45-120">可用的部署选项具体取决于你定位的 Office 主机以及所创建的加载项的类型。</span><span class="sxs-lookup"><span data-stu-id="e5a45-120">The deployment options that are available depend on the Office host that you're targeting and the type of add-in you create.</span></span>

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a><span data-ttu-id="e5a45-121">Word、Excel 和 PowerPoint 加载项的部署选项</span><span class="sxs-lookup"><span data-stu-id="e5a45-121">Deployment options for Word, Excel, and PowerPoint add-ins</span></span>

| <span data-ttu-id="e5a45-122">扩展点</span><span class="sxs-lookup"><span data-stu-id="e5a45-122">Extension point</span></span> | <span data-ttu-id="e5a45-123">旁加载</span><span class="sxs-lookup"><span data-stu-id="e5a45-123">Sideloading</span></span> | <span data-ttu-id="e5a45-124">Office 365 管理中心</span><span class="sxs-lookup"><span data-stu-id="e5a45-124">Office 365 admin center</span></span> |<span data-ttu-id="e5a45-125">AppSource</span><span class="sxs-lookup"><span data-stu-id="e5a45-125">AppSource</span></span>   | <span data-ttu-id="e5a45-126">SharePoint 目录\*</span><span class="sxs-lookup"><span data-stu-id="e5a45-126">SharePoint catalog\*</span></span> |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| <span data-ttu-id="e5a45-127">内容</span><span class="sxs-lookup"><span data-stu-id="e5a45-127">Content</span></span>         | <span data-ttu-id="e5a45-128">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-128">X</span></span>           | <span data-ttu-id="e5a45-129">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-129">X</span></span>                       | <span data-ttu-id="e5a45-130">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-130">X</span></span>          | <span data-ttu-id="e5a45-131">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-131">X</span></span>                    |
| <span data-ttu-id="e5a45-132">任务窗格</span><span class="sxs-lookup"><span data-stu-id="e5a45-132">Task pane</span></span>       | <span data-ttu-id="e5a45-133">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-133">X</span></span>           | <span data-ttu-id="e5a45-134">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-134">X</span></span>                       | <span data-ttu-id="e5a45-135">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-135">X</span></span>          | <span data-ttu-id="e5a45-136">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-136">X</span></span>                    |
| <span data-ttu-id="e5a45-137">命令</span><span class="sxs-lookup"><span data-stu-id="e5a45-137">Command</span></span>         | <span data-ttu-id="e5a45-138">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-138">X</span></span>           | <span data-ttu-id="e5a45-139">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-139">X</span></span>                       | <span data-ttu-id="e5a45-140">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-140">X</span></span>          |                      |

<span data-ttu-id="e5a45-141">&#42; SharePoint 目录不支持 Mac 版 Office。</span><span class="sxs-lookup"><span data-stu-id="e5a45-141">&#42; SharePoint catalogs do not support Office on Mac.</span></span>

### <a name="deployment-options-for-outlook-add-ins"></a><span data-ttu-id="e5a45-142">Outlook 加载项的部署选项</span><span class="sxs-lookup"><span data-stu-id="e5a45-142">Deployment options for Outlook add-ins</span></span>

| <span data-ttu-id="e5a45-143">扩展点</span><span class="sxs-lookup"><span data-stu-id="e5a45-143">Extension point</span></span> | <span data-ttu-id="e5a45-144">旁加载</span><span class="sxs-lookup"><span data-stu-id="e5a45-144">Sideloading</span></span> | <span data-ttu-id="e5a45-145">Exchange 服务器</span><span class="sxs-lookup"><span data-stu-id="e5a45-145">Exchange server</span></span> | <span data-ttu-id="e5a45-146">AppSource</span><span class="sxs-lookup"><span data-stu-id="e5a45-146">AppSource</span></span>    |
|:----------------|:-----------:|:---------------:|:------------:|
| <span data-ttu-id="e5a45-147">邮件应用</span><span class="sxs-lookup"><span data-stu-id="e5a45-147">Mail app</span></span>        | <span data-ttu-id="e5a45-148">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-148">X</span></span>           | <span data-ttu-id="e5a45-149">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-149">X</span></span>               | <span data-ttu-id="e5a45-150">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-150">X</span></span>            |
| <span data-ttu-id="e5a45-151">命令</span><span class="sxs-lookup"><span data-stu-id="e5a45-151">Command</span></span>         | <span data-ttu-id="e5a45-152">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-152">X</span></span>           | <span data-ttu-id="e5a45-153">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-153">X</span></span>               | <span data-ttu-id="e5a45-154">X</span><span class="sxs-lookup"><span data-stu-id="e5a45-154">X</span></span>            |

## <a name="deployment-methods"></a><span data-ttu-id="e5a45-155">部署方法</span><span class="sxs-lookup"><span data-stu-id="e5a45-155">Deployment methods</span></span>

<span data-ttu-id="e5a45-156">以下各节提供了有关向组织中的用户分发 Office 加载项的最常用部署方法的其他信息。</span><span class="sxs-lookup"><span data-stu-id="e5a45-156">The following sections provide additional information about the deployment methods that are most commonly used to distribute Office Add-ins to users within an organization.</span></span>

<span data-ttu-id="e5a45-157">有关最终用户如何获取、插入和运行加载项的信息，请参阅[开始使用 Office 加载项](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)。</span><span class="sxs-lookup"><span data-stu-id="e5a45-157">For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).</span></span>

### <a name="centralized-deployment-via-the-office-365-admin-center"></a><span data-ttu-id="e5a45-158">通过 Office 365 管理中心进行集中部署</span><span class="sxs-lookup"><span data-stu-id="e5a45-158">Centralized Deployment via the Office 365 admin center</span></span> 

<span data-ttu-id="e5a45-p101">通过 Office 365 管理中心，管理员可以为组织中的用户和组轻松部署 Office 加载项。在管理员通过管理中心部署加载项后，用户可以立即在 Office 应用中使用加载项，而无需进行任何客户端配置。通过集中部署，可以部署内部加载项和 ISV 提供的加载项。</span><span class="sxs-lookup"><span data-stu-id="e5a45-p101">The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="e5a45-162">有关详细信息，请参阅[通过 Office 365 管理中心进行集中部署来发布 Office 加载项](centralized-deployment.md)。</span><span class="sxs-lookup"><span data-stu-id="e5a45-162">For more information, see [Publish Office Add-ins using Centralized Deployment via the Office 365 admin center](centralized-deployment.md).</span></span>

### <a name="sharepoint-app-catalog-deployment"></a><span data-ttu-id="e5a45-163">SharePoint 应用目录部署</span><span class="sxs-lookup"><span data-stu-id="e5a45-163">SharePoint app catalog deployment</span></span>

<span data-ttu-id="e5a45-p102">SharePoint 应用目录是特殊网站集，创建后可用于托管 Word、Excel 和 PowerPoint 加载项。由于 SharePoint 目录不支持在清单的 `VersionOverrides` 节点中实现的新加载项功能（包括加载项命令），因此建议尽可能通过管理中心进行集中部署。通过 SharePoint 目录部署的加载项命令默认在任务窗格中打开。</span><span class="sxs-lookup"><span data-stu-id="e5a45-p102">A SharePoint app catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, we recommend that you use Centralized Deployment via the admin center if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.</span></span>

<span data-ttu-id="e5a45-p103">如果要在本地环境中部署外接程序，请使用 SharePoint 目录。有关详细信息，请参阅[将任务窗格和内容外接程序发布到 SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。</span><span class="sxs-lookup"><span data-stu-id="e5a45-p103">If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>

> [!NOTE]
> <span data-ttu-id="e5a45-168">SharePoint 目录不支持 Mac 版 Office。</span><span class="sxs-lookup"><span data-stu-id="e5a45-168">SharePoint catalogs do not support Office on Mac.</span></span> <span data-ttu-id="e5a45-169">若要向 Mac 客户端部署 Office 加载项，必须将其提交到 [AppSource](/office/dev/store/submit-to-the-office-store)。</span><span class="sxs-lookup"><span data-stu-id="e5a45-169">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

### <a name="outlook-add-in-deployment"></a><span data-ttu-id="e5a45-170">Outlook 加载项部署</span><span class="sxs-lookup"><span data-stu-id="e5a45-170">Outlook add-in deployment</span></span>

<span data-ttu-id="e5a45-171">对于不使用 Azure AD 标识服务的本地和联机环境，可以通过 Exchange 服务器部署 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="e5a45-171">For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server.</span></span>

<span data-ttu-id="e5a45-172">Outlook 外接程序部署需要以下内容：</span><span class="sxs-lookup"><span data-stu-id="e5a45-172">Outlook add-in deployment requires:</span></span>

- <span data-ttu-id="e5a45-173">Office 365、Exchange Online 或 Exchange Server 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="e5a45-173">Office 365, Exchange Online, or Exchange Server 2013 or later</span></span>
- <span data-ttu-id="e5a45-174">Outlook 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="e5a45-174">Outlook 2013 or later</span></span>

<span data-ttu-id="e5a45-p105">若要将加载项分配给租户，请使用 Exchange 管理中心通过文件或 URL 直接上传清单，或从 AppSource 添加加载项。若要将加载项分配给单个用户，必须使用 Exchange PowerShell。有关详细信息，请参阅 TechNet 上的[为组织安装或删除 Outlook 加载项](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx)。</span><span class="sxs-lookup"><span data-stu-id="e5a45-p105">To assign add-ins to tenants, you use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from AppSource. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) on TechNet.</span></span>

## <a name="see-also"></a><span data-ttu-id="e5a45-178">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e5a45-178">See also</span></span>

- [<span data-ttu-id="e5a45-179">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="e5a45-179">Sideload Outlook add-ins for testing</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- <span data-ttu-id="e5a45-180">[提交到 AppSource][AppSource]</span><span class="sxs-lookup"><span data-stu-id="e5a45-180">[Submit to AppSource][AppSource]</span></span>
- [<span data-ttu-id="e5a45-181">Office 加载项的设计准则</span><span class="sxs-lookup"><span data-stu-id="e5a45-181">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="e5a45-182">创建有效的 AppSource 一览</span><span class="sxs-lookup"><span data-stu-id="e5a45-182">Create effective AppSource listings</span></span>](/office/dev/store/create-effective-office-store-listings)
- [<span data-ttu-id="e5a45-183">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="e5a45-183">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
