---
title: 部署和发布 Office 外接程序
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 83581b729f5004c36d267bda14795275a5153a9c
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925344"
---
# <a name="deploy-and-publish-your-office-add-in"></a><span data-ttu-id="1f408-102">部署和发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="1f408-102">Deploy and publish your Office Add-in</span></span>

<span data-ttu-id="1f408-103">可以使用几种方法之一来部署 Office 外接程序，以用于对用户进行测试或分发：</span><span class="sxs-lookup"><span data-stu-id="1f408-103">You can use one of several methods to deploy your Office Add-in for testing or distribution to users.</span></span>

|<span data-ttu-id="1f408-104">**方法**</span><span class="sxs-lookup"><span data-stu-id="1f408-104">**Method**</span></span>|<span data-ttu-id="1f408-105">**使用...**</span><span class="sxs-lookup"><span data-stu-id="1f408-105">**Use...**</span></span>|
|:---------|:------------|
|[<span data-ttu-id="1f408-106">旁加载</span><span class="sxs-lookup"><span data-stu-id="1f408-106">Sideloading</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|<span data-ttu-id="1f408-107">在开发过程中测试在 Windows、Office Online、iPad 或 Mac 上运行的加载项。</span><span class="sxs-lookup"><span data-stu-id="1f408-107">As part of your development process, to test your add-in running on Windows, Office Online, iPad, or Mac.</span></span>|
|[<span data-ttu-id="1f408-108">集中部署</span><span class="sxs-lookup"><span data-stu-id="1f408-108">Centralized Deployment</span></span>](centralized-deployment.md)|<span data-ttu-id="1f408-109">在云或混合部署中，使用 Office 365 管理中心将加载项分发给组织中的用户。</span><span class="sxs-lookup"><span data-stu-id="1f408-109">In a cloud or hybrid deployment, to distribute your add-in to users in your organization by using the Office 365 admin center.</span></span>|
|[<span data-ttu-id="1f408-110">SharePoint 目录</span><span class="sxs-lookup"><span data-stu-id="1f408-110">SharePoint catalog</span></span>](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|<span data-ttu-id="1f408-111">在本地环境中，用于向组织用户分发加载项。</span><span class="sxs-lookup"><span data-stu-id="1f408-111">In an on-premises environment, to distribute your add-in to users in your organization.</span></span>|
|[<span data-ttu-id="1f408-112">AppSource</span><span class="sxs-lookup"><span data-stu-id="1f408-112">AppSource</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)|<span data-ttu-id="1f408-113">用于向用户公开分发加载项。</span><span class="sxs-lookup"><span data-stu-id="1f408-113">To distribute your add-in publicly to users.</span></span>|
|[<span data-ttu-id="1f408-114">Exchange 服务器</span><span class="sxs-lookup"><span data-stu-id="1f408-114">Exchange server</span></span>](#outlook-add-in-deployment)|<span data-ttu-id="1f408-115">在本地或在线环境中，用于向用户分发 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="1f408-115">In an on-premises or online environment, to distribute Outlook add-ins to users.</span></span>|
|[<span data-ttu-id="1f408-116">网络共享</span><span class="sxs-lookup"><span data-stu-id="1f408-116">Network share</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|<span data-ttu-id="1f408-117">在网络上的 Windows 计算机中（要在其中托管加载项），转到要用作共享文件夹目录的文件夹的父文件夹或驱动器号。</span><span class="sxs-lookup"><span data-stu-id="1f408-117">On a Windows computer on a network where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>|

> [!NOTE]
> <span data-ttu-id="1f408-p101">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="1f408-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="deployment-options-by-office-host"></a><span data-ttu-id="1f408-120">Office 主机提供的部署选项</span><span class="sxs-lookup"><span data-stu-id="1f408-120">Deployment options by Office host</span></span>

<span data-ttu-id="1f408-121">可用的部署选项具体取决于你定位的 Office 主机以及所创建的加载项的类型。</span><span class="sxs-lookup"><span data-stu-id="1f408-121">The deployment options that are available depend on the Office host that you're targeting and the type of add-in you create.</span></span>

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a><span data-ttu-id="1f408-122">Word、Excel 和 PowerPoint 加载项的部署选项</span><span class="sxs-lookup"><span data-stu-id="1f408-122">Deployment options for Word, Excel, and PowerPoint add-ins</span></span>

| <span data-ttu-id="1f408-123">扩展点</span><span class="sxs-lookup"><span data-stu-id="1f408-123">Extension point</span></span> | <span data-ttu-id="1f408-124">旁加载</span><span class="sxs-lookup"><span data-stu-id="1f408-124">Sideloading</span></span> | <span data-ttu-id="1f408-125">Office 365 管理中心</span><span class="sxs-lookup"><span data-stu-id="1f408-125">Office 365 admin center</span></span> |<span data-ttu-id="1f408-126">AppSource</span><span class="sxs-lookup"><span data-stu-id="1f408-126">AppSource</span></span>| <span data-ttu-id="1f408-127">SharePoint 目录\*</span><span class="sxs-lookup"><span data-stu-id="1f408-127">SharePoint catalog\*</span></span>  |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| <span data-ttu-id="1f408-128">内容</span><span class="sxs-lookup"><span data-stu-id="1f408-128">Content</span></span>         | <span data-ttu-id="1f408-129">X</span><span class="sxs-lookup"><span data-stu-id="1f408-129">X</span></span>           | <span data-ttu-id="1f408-130">X</span><span class="sxs-lookup"><span data-stu-id="1f408-130">X</span></span>                       | <span data-ttu-id="1f408-131">X</span><span class="sxs-lookup"><span data-stu-id="1f408-131">X</span></span>          | <span data-ttu-id="1f408-132">X</span><span class="sxs-lookup"><span data-stu-id="1f408-132">X</span></span>                    |
| <span data-ttu-id="1f408-133">任务窗格</span><span class="sxs-lookup"><span data-stu-id="1f408-133">Task pane</span></span>       | <span data-ttu-id="1f408-134">X</span><span class="sxs-lookup"><span data-stu-id="1f408-134">X</span></span>           | <span data-ttu-id="1f408-135">X</span><span class="sxs-lookup"><span data-stu-id="1f408-135">X</span></span>                       | <span data-ttu-id="1f408-136">X</span><span class="sxs-lookup"><span data-stu-id="1f408-136">X</span></span>          | <span data-ttu-id="1f408-137">X</span><span class="sxs-lookup"><span data-stu-id="1f408-137">X</span></span>                    |
| <span data-ttu-id="1f408-138">命令</span><span class="sxs-lookup"><span data-stu-id="1f408-138">Command</span></span>           | <span data-ttu-id="1f408-139">X</span><span class="sxs-lookup"><span data-stu-id="1f408-139">X</span></span>           | <span data-ttu-id="1f408-140">X</span><span class="sxs-lookup"><span data-stu-id="1f408-140">X</span></span>                       | <span data-ttu-id="1f408-141">X</span><span class="sxs-lookup"><span data-stu-id="1f408-141">X</span></span>          |                      |

<span data-ttu-id="1f408-142">\* SharePoint 目录不支持 Office 2016 for Mac。</span><span class="sxs-lookup"><span data-stu-id="1f408-142">&#42; SharePoint catalogs do not support Office 2016 for Mac.</span></span>

### <a name="deployment-options-for-outlook-add-ins"></a><span data-ttu-id="1f408-143">Outlook 外接程序的部署选项</span><span class="sxs-lookup"><span data-stu-id="1f408-143">Deployment options for Outlook add-ins</span></span>

| <span data-ttu-id="1f408-144">扩展点</span><span class="sxs-lookup"><span data-stu-id="1f408-144">Extension point</span></span> | <span data-ttu-id="1f408-145">旁加载</span><span class="sxs-lookup"><span data-stu-id="1f408-145">Sideloading</span></span> | <span data-ttu-id="1f408-146">Exchange 服务器</span><span class="sxs-lookup"><span data-stu-id="1f408-146">Exchange server</span></span> | <span data-ttu-id="1f408-147">AppSource</span><span class="sxs-lookup"><span data-stu-id="1f408-147">AppSource</span></span> |
|:----------------|:-----------:|:---------------:|:------------:|
| <span data-ttu-id="1f408-148">邮件应用</span><span class="sxs-lookup"><span data-stu-id="1f408-148">Mail app</span></span>        | <span data-ttu-id="1f408-149">X</span><span class="sxs-lookup"><span data-stu-id="1f408-149">X</span></span>           | <span data-ttu-id="1f408-150">X</span><span class="sxs-lookup"><span data-stu-id="1f408-150">X</span></span>               | <span data-ttu-id="1f408-151">X</span><span class="sxs-lookup"><span data-stu-id="1f408-151">X</span></span>            |
| <span data-ttu-id="1f408-152">命令</span><span class="sxs-lookup"><span data-stu-id="1f408-152">Command</span></span>         | <span data-ttu-id="1f408-153">X</span><span class="sxs-lookup"><span data-stu-id="1f408-153">X</span></span>           | <span data-ttu-id="1f408-154">X</span><span class="sxs-lookup"><span data-stu-id="1f408-154">X</span></span>               | <span data-ttu-id="1f408-155">X</span><span class="sxs-lookup"><span data-stu-id="1f408-155">X</span></span>            |

## <a name="deployment-methods"></a><span data-ttu-id="1f408-156">部署方法</span><span class="sxs-lookup"><span data-stu-id="1f408-156">Deployment methods</span></span>

<span data-ttu-id="1f408-157">以下各节提供了有关向组织中的用户分发 Office 加载项的最常用的部署方法的其他信息。</span><span class="sxs-lookup"><span data-stu-id="1f408-157">The following sections provide additional information about the deployment methods that are most commonly used to distribute Office add-ins to users within an organization.</span></span>

<span data-ttu-id="1f408-158">有关最终用户如何获取、插入和运行加载项的信息，请参阅[开始使用 Office 加载项](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)。</span><span class="sxs-lookup"><span data-stu-id="1f408-158">For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).</span></span>

### <a name="centralized-deployment-via-the-office-365-admin-center"></a><span data-ttu-id="1f408-159">通过 Office 365 管理中心进行集中部署</span><span class="sxs-lookup"><span data-stu-id="1f408-159">Centralized Deployment via the Office 365 admin center</span></span> 

<span data-ttu-id="1f408-p102">通过 Office 365 管理中心，管理员可以为组织中的用户和组轻松部署 Office 加载项。在管理员通过管理中心部署加载项后，用户可以立即在 Office 应用中使用加载项，而无需进行任何客户端配置。通过集中部署，可以部署内部加载项和 ISV 提供的加载项。</span><span class="sxs-lookup"><span data-stu-id="1f408-p102">The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="1f408-163">有关详细信息，请参阅[通过 Office 365 管理中心进行集中部署来发布 Office 加载项](centralized-deployment.md)。</span><span class="sxs-lookup"><span data-stu-id="1f408-163">For more information, see [Publish Office Add-ins using Centralized Deployment via the Office 365 admin center](centralized-deployment.md).</span></span>

### <a name="sharepoint-catalog-deployment"></a><span data-ttu-id="1f408-164">SharePoint 目录部署</span><span class="sxs-lookup"><span data-stu-id="1f408-164">SharePoint catalog deployment</span></span>

<span data-ttu-id="1f408-p103">SharePoint 加载项目录是特殊网站集，创建后可用于托管 Word、Excel 和 PowerPoint 加载项。由于 SharePoint 目录不支持在清单的 `VersionOverrides` 节点中实现的新加载项功能（包括加载项命令），因此建议尽可能通过管理中心进行集中部署。通过 SharePoint 目录部署的加载项命令默认在任务窗格中打开。</span><span class="sxs-lookup"><span data-stu-id="1f408-p103">A SharePoint add-in catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, we recommend that you use Centralized Deployment via the admin center if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.</span></span>

<span data-ttu-id="1f408-p104">如果要在本地环境中部署外接程序，请使用 SharePoint 目录。有关详细信息，请参阅[将任务窗格和内容外接程序发布到 SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。</span><span class="sxs-lookup"><span data-stu-id="1f408-p104">If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>

> [!NOTE]
> <span data-ttu-id="1f408-p105">Office 2016 for Mac 不支持 SharePoint 目录。若要向 Mac 客户端部署 Office 加载项，必须将它们提交到 [AppSource]。</span><span class="sxs-lookup"><span data-stu-id="1f408-p105">SharePoint catalogs do not support Office 2016 for Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource].</span></span> 

### <a name="outlook-add-in-deployment"></a><span data-ttu-id="1f408-171">Outlook 加载项部署</span><span class="sxs-lookup"><span data-stu-id="1f408-171">Outlook add-in deployment</span></span>

<span data-ttu-id="1f408-172">对于不使用 Azure AD 标识服务的本地和联机环境，可以通过 Exchange 服务器部署 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="1f408-172">For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server.</span></span> 

<span data-ttu-id="1f408-173">Outlook 外接程序部署需要以下内容：</span><span class="sxs-lookup"><span data-stu-id="1f408-173">Outlook add-in deployment requires:</span></span>

- <span data-ttu-id="1f408-174">Office 365、Exchange Online 或 Exchange Server 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="1f408-174">Office 365, Exchange Online, or Exchange Server 2013 or later</span></span>
- <span data-ttu-id="1f408-175">Outlook 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="1f408-175">Outlook 2013 or later</span></span>

<span data-ttu-id="1f408-p106">若要将加载项分配给租户，请使用 Exchange 管理中心通过文件或 URL 直接上传清单，或从 AppSource 添加加载项。若要将加载项分配给单个用户，必须使用 Exchange PowerShell。有关详细信息，请参阅 TechNet 上的[为组织安装或删除 Outlook 加载项](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx)。</span><span class="sxs-lookup"><span data-stu-id="1f408-p106">To assign add-ins to tenants, you use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from AppSource. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) on TechNet.</span></span>

## <a name="see-also"></a><span data-ttu-id="1f408-179">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1f408-179">See also</span></span>

- [<span data-ttu-id="1f408-180">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="1f408-180">Sideload Outlook add-ins for testing</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- <span data-ttu-id="1f408-181">[提交到 AppSource][AppSource]</span><span class="sxs-lookup"><span data-stu-id="1f408-181">[Submit to AppSource][AppSource]</span></span>
- [<span data-ttu-id="1f408-182">Office 加载项设计指南</span><span class="sxs-lookup"><span data-stu-id="1f408-182">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="1f408-183">创建有效的 AppSource 一览</span><span class="sxs-lookup"><span data-stu-id="1f408-183">Create effective AppSource listings</span></span>](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)
- [<span data-ttu-id="1f408-184">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="1f408-184">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)

[AppSource]: https://docs.microsoft.com/office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
