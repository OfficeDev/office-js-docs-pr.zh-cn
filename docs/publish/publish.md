---
title: ????? Office ????
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: d8264667306dcdac2e9d5e5d6e6607a2a2100546
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="deploy-and-publish-your-office-add-in"></a><span data-ttu-id="abecc-102">????? Office ???</span><span class="sxs-lookup"><span data-stu-id="abecc-102">Deploy and publish your Office Add-in</span></span>

<span data-ttu-id="abecc-103">????????????? Office ???????????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-103">You can use one of several methods to deploy your Office Add-in for testing or distribution to users.</span></span>

|<span data-ttu-id="abecc-104">**??**</span><span class="sxs-lookup"><span data-stu-id="abecc-104">**Method**</span></span>|<span data-ttu-id="abecc-105">**??...**</span><span class="sxs-lookup"><span data-stu-id="abecc-105">**Use...**</span></span>|
|:---------|:------------|
|[<span data-ttu-id="abecc-106">???</span><span class="sxs-lookup"><span data-stu-id="abecc-106">Sideloading</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|<span data-ttu-id="abecc-107">????????? Windows?Office Online?iPad ? Mac ????????</span><span class="sxs-lookup"><span data-stu-id="abecc-107">As part of your development process, to test your add-in running on Windows, Office Online, iPad, or Mac.</span></span>|
|[<span data-ttu-id="abecc-108">????</span><span class="sxs-lookup"><span data-stu-id="abecc-108">Centralized Deployment</span></span>](centralized-deployment.md)|<span data-ttu-id="abecc-109">??????????? Office 365 ??????????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-109">In a cloud or hybrid deployment, to distribute your add-in to users in your organization by using the Office 365 admin center.</span></span>|
|[<span data-ttu-id="abecc-110">SharePoint ??</span><span class="sxs-lookup"><span data-stu-id="abecc-110">SharePoint catalog</span></span>](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|<span data-ttu-id="abecc-111">????????????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-111">In an on-premises environment, to distribute your add-in to users in your organization.</span></span>|
|[<span data-ttu-id="abecc-112">AppSource</span><span class="sxs-lookup"><span data-stu-id="abecc-112">AppSource</span></span>](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)|<span data-ttu-id="abecc-113">?????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-113">To distribute your add-in publicly to users.</span></span>|
|[<span data-ttu-id="abecc-114">Exchange ???</span><span class="sxs-lookup"><span data-stu-id="abecc-114">Exchange server</span></span>](#outlook-add-in-deployment)|<span data-ttu-id="abecc-115">????????????????? Outlook ????</span><span class="sxs-lookup"><span data-stu-id="abecc-115">In an on-premises or online environment, to distribute Outlook add-ins to users.</span></span>|
|[<span data-ttu-id="abecc-116">????</span><span class="sxs-lookup"><span data-stu-id="abecc-116">Network share</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|<span data-ttu-id="abecc-117">????? Windows ???????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-117">On a Windows computer on a network where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>|

> [!NOTE]
> <span data-ttu-id="abecc-p101">????????[??](../publish/publish.md)? AppSource ???? Office ???????? [AppSource ????](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)??????????????????????????????????????????[? 4.12 ??](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)?? [Office ?????????](../overview/office-add-in-availability.md)????</span><span class="sxs-lookup"><span data-stu-id="abecc-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="deployment-options-by-office-host"></a><span data-ttu-id="abecc-120">Office ?????????</span><span class="sxs-lookup"><span data-stu-id="abecc-120">Deployment options by Office host</span></span>

<span data-ttu-id="abecc-121">???????????????? Office ???????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-121">The deployment options that are available depend on the Office host that you're targeting and the type of add-in you create.</span></span>

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a><span data-ttu-id="abecc-122">Word?Excel ? PowerPoint ????????</span><span class="sxs-lookup"><span data-stu-id="abecc-122">Deployment options for Word, Excel, and PowerPoint add-ins</span></span>

| <span data-ttu-id="abecc-123">???</span><span class="sxs-lookup"><span data-stu-id="abecc-123">Extension point</span></span> | <span data-ttu-id="abecc-124">???</span><span class="sxs-lookup"><span data-stu-id="abecc-124">Sideloading</span></span> | <span data-ttu-id="abecc-125">Office 365 ????</span><span class="sxs-lookup"><span data-stu-id="abecc-125">Office 365 admin center</span></span> |<span data-ttu-id="abecc-126">AppSource</span><span class="sxs-lookup"><span data-stu-id="abecc-126">AppSource</span></span>| <span data-ttu-id="abecc-127">SharePoint ??\*</span><span class="sxs-lookup"><span data-stu-id="abecc-127">SharePoint catalog\*</span></span>  |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| <span data-ttu-id="abecc-128">??</span><span class="sxs-lookup"><span data-stu-id="abecc-128">Content</span></span>         | <span data-ttu-id="abecc-129">X</span><span class="sxs-lookup"><span data-stu-id="abecc-129">X</span></span>           | <span data-ttu-id="abecc-130">X</span><span class="sxs-lookup"><span data-stu-id="abecc-130">X</span></span>                       | <span data-ttu-id="abecc-131">X</span><span class="sxs-lookup"><span data-stu-id="abecc-131">X</span></span>          | <span data-ttu-id="abecc-132">X</span><span class="sxs-lookup"><span data-stu-id="abecc-132">X</span></span>                    |
| <span data-ttu-id="abecc-133">????</span><span class="sxs-lookup"><span data-stu-id="abecc-133">Task pane</span></span>       | <span data-ttu-id="abecc-134">X</span><span class="sxs-lookup"><span data-stu-id="abecc-134">X</span></span>           | <span data-ttu-id="abecc-135">X</span><span class="sxs-lookup"><span data-stu-id="abecc-135">X</span></span>                       | <span data-ttu-id="abecc-136">X</span><span class="sxs-lookup"><span data-stu-id="abecc-136">X</span></span>          | <span data-ttu-id="abecc-137">X</span><span class="sxs-lookup"><span data-stu-id="abecc-137">X</span></span>                    |
| <span data-ttu-id="abecc-138">??</span><span class="sxs-lookup"><span data-stu-id="abecc-138">Command</span></span>           | <span data-ttu-id="abecc-139">X</span><span class="sxs-lookup"><span data-stu-id="abecc-139">X</span></span>           | <span data-ttu-id="abecc-140">X</span><span class="sxs-lookup"><span data-stu-id="abecc-140">X</span></span>                       | <span data-ttu-id="abecc-141">X</span><span class="sxs-lookup"><span data-stu-id="abecc-141">X</span></span>          |                      |

<span data-ttu-id="abecc-142">\* SharePoint ????? Office 2016 for Mac?</span><span class="sxs-lookup"><span data-stu-id="abecc-142">&#42; SharePoint catalogs do not support Office 2016 for Mac.</span></span>

### <a name="deployment-options-for-outlook-add-ins"></a><span data-ttu-id="abecc-143">Outlook ?????????</span><span class="sxs-lookup"><span data-stu-id="abecc-143">Deployment options for Outlook add-ins</span></span>

| <span data-ttu-id="abecc-144">???</span><span class="sxs-lookup"><span data-stu-id="abecc-144">Extension point</span></span> | <span data-ttu-id="abecc-145">???</span><span class="sxs-lookup"><span data-stu-id="abecc-145">Sideloading</span></span> | <span data-ttu-id="abecc-146">Exchange ???</span><span class="sxs-lookup"><span data-stu-id="abecc-146">Exchange server</span></span> | <span data-ttu-id="abecc-147">AppSource</span><span class="sxs-lookup"><span data-stu-id="abecc-147">AppSource</span></span> |
|:----------------|:-----------:|:---------------:|:------------:|
| <span data-ttu-id="abecc-148">????</span><span class="sxs-lookup"><span data-stu-id="abecc-148">Mail app</span></span>        | <span data-ttu-id="abecc-149">X</span><span class="sxs-lookup"><span data-stu-id="abecc-149">X</span></span>           | <span data-ttu-id="abecc-150">X</span><span class="sxs-lookup"><span data-stu-id="abecc-150">X</span></span>               | <span data-ttu-id="abecc-151">X</span><span class="sxs-lookup"><span data-stu-id="abecc-151">X</span></span>            |
| <span data-ttu-id="abecc-152">??</span><span class="sxs-lookup"><span data-stu-id="abecc-152">Command</span></span>         | <span data-ttu-id="abecc-153">X</span><span class="sxs-lookup"><span data-stu-id="abecc-153">X</span></span>           | <span data-ttu-id="abecc-154">X</span><span class="sxs-lookup"><span data-stu-id="abecc-154">X</span></span>               | <span data-ttu-id="abecc-155">X</span><span class="sxs-lookup"><span data-stu-id="abecc-155">X</span></span>            |

## <a name="deployment-methods"></a><span data-ttu-id="abecc-156">????</span><span class="sxs-lookup"><span data-stu-id="abecc-156">Deployment methods</span></span>

<span data-ttu-id="abecc-157">?????????????????? Office ??????????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-157">The following sections provide additional information about the deployment methods that are most commonly used to distribute Office add-ins to users within an organization.</span></span>

<span data-ttu-id="abecc-158">??????????????????????????[???? Office ???](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)?</span><span class="sxs-lookup"><span data-stu-id="abecc-158">For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).</span></span>

### <a name="centralized-deployment-via-the-office-365-admin-center"></a><span data-ttu-id="abecc-159">?? Office 365 ??????????</span><span class="sxs-lookup"><span data-stu-id="abecc-159">Centralized Deployment via the Office 365 admin center</span></span> 

<span data-ttu-id="abecc-p102">?? Office 365 ??????????????????????? Office ???????????????????????????? Office ??????????????????????????????????????? ISV ???????</span><span class="sxs-lookup"><span data-stu-id="abecc-p102">The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="abecc-163">??????????[?? Office 365 ????????????? Office ???](centralized-deployment.md)?</span><span class="sxs-lookup"><span data-stu-id="abecc-163">For more information, see [Publish Office Add-ins using Centralized Deployment via the Office 365 admin center](centralized-deployment.md).</span></span>

### <a name="sharepoint-catalog-deployment"></a><span data-ttu-id="abecc-164">SharePoint ????</span><span class="sxs-lookup"><span data-stu-id="abecc-164">SharePoint catalog deployment</span></span>

<span data-ttu-id="abecc-p103">SharePoint ???????????????????? Word?Excel ? PowerPoint ?????? SharePoint ????????? `VersionOverrides` ???????????????????????????????????????????? SharePoint ?????????????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-p103">A SharePoint add-in catalog is a special site collection that you can create to host Word, Excel, and PowerPoint add-ins. Because SharePoint catalogs don't support new add-in features implemented in the `VersionOverrides` node of the manifest, including add-in commands, we recommend that you use Centralized Deployment via the admin center if possible. Add-in commands deployed via a SharePoint catalog open in a task pane by default.</span></span>

<span data-ttu-id="abecc-p104">??????????????????? SharePoint ?????????????[??????????????? SharePoint ??](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)?</span><span class="sxs-lookup"><span data-stu-id="abecc-p104">If you are deploying add-ins in an on-premises environment, use a SharePoint catalog. For details, see [Publish task pane and content add-ins to a SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>

> [!NOTE]
> <span data-ttu-id="abecc-p105">Office 2016 for Mac ??? SharePoint ?????? Mac ????? Office ???????????? [AppSource]?</span><span class="sxs-lookup"><span data-stu-id="abecc-p105">SharePoint catalogs do not support Office 2016 for Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource].</span></span> 

### <a name="outlook-add-in-deployment"></a><span data-ttu-id="abecc-171">Outlook ?????</span><span class="sxs-lookup"><span data-stu-id="abecc-171">Outlook add-in deployment</span></span>

<span data-ttu-id="abecc-172">????? Azure AD ????????????????? Exchange ????? Outlook ?????</span><span class="sxs-lookup"><span data-stu-id="abecc-172">For on-premises and online environments that do not use the Azure AD identity service, you can deploy Outlook add-ins via the Exchange server.</span></span> 

<span data-ttu-id="abecc-173">Outlook ?????????????</span><span class="sxs-lookup"><span data-stu-id="abecc-173">Outlook add-in deployment requires:</span></span>

- <span data-ttu-id="abecc-174">Office 365?Exchange Online ? Exchange Server 2013 ?????</span><span class="sxs-lookup"><span data-stu-id="abecc-174">Office 365, Exchange Online, or Exchange Server 2013 or later</span></span>
- <span data-ttu-id="abecc-175">Outlook 2013 ?????</span><span class="sxs-lookup"><span data-stu-id="abecc-175">Outlook 2013 or later</span></span>

<span data-ttu-id="abecc-p106">??????????????? Exchange ????????? URL ????????? AppSource ???????????????????????? Exchange PowerShell??????????? TechNet ??[???????? Outlook ???](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx)?</span><span class="sxs-lookup"><span data-stu-id="abecc-p106">To assign add-ins to tenants, you use the Exchange admin center to upload a manifest directly, either from a file or a URL, or add an add-in from AppSource. To assign add-ins to individual users, you must use Exchange PowerShell. For details, see [Install or remove Outlook add-ins for your organization](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx) on TechNet.</span></span>

## <a name="see-also"></a><span data-ttu-id="abecc-179">????</span><span class="sxs-lookup"><span data-stu-id="abecc-179">See also</span></span>

- [<span data-ttu-id="abecc-180">??? Outlook ???????</span><span class="sxs-lookup"><span data-stu-id="abecc-180">Sideload Outlook add-ins for testing</span></span>](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- <span data-ttu-id="abecc-181">[??? AppSource][AppSource]</span><span class="sxs-lookup"><span data-stu-id="abecc-181">[Submit to AppSource][AppSource]</span></span>
- [<span data-ttu-id="abecc-182">Office ???????</span><span class="sxs-lookup"><span data-stu-id="abecc-182">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="abecc-183">????? AppSource ??</span><span class="sxs-lookup"><span data-stu-id="abecc-183">Create effective AppSource listings</span></span>](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings)
- [<span data-ttu-id="abecc-184">?? Office ?????????</span><span class="sxs-lookup"><span data-stu-id="abecc-184">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)

[AppSource]: https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
