---
title: Office 客户端应用程序和平台的 Office 加载项可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 07/13/2021
localization_priority: Priority
ms.openlocfilehash: 7b3bd770d74f29d1a0b27da5080284aa62146101
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455493"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a><span data-ttu-id="46575-103">Office 客户端应用程序和平台的 Office 加载项可用性</span><span class="sxs-lookup"><span data-stu-id="46575-103">Office client application and platform availability for Office Add-ins</span></span>

<span data-ttu-id="46575-p101">为了能够按预期运行，Office 加载项可能会依赖特定的 Office 应用程序、要求集、API 成员或 API 版本。下表列出了每个 Office 应用程序目前所支持的平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="46575-p101">To work as expected, your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span data-ttu-id="46575-106"><span>Excel</span></a></span><span class="sxs-lookup"><span data-stu-id="46575-106"><span>Excel</span></a></span></span>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span data-ttu-id="46575-107"><span>OneNote</span></a></span><span class="sxs-lookup"><span data-stu-id="46575-107"><span>OneNote</span></a></span></span>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span data-ttu-id="46575-108"><span>Outlook</span></a></span><span class="sxs-lookup"><span data-stu-id="46575-108"><span>Outlook</span></a></span></span>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span data-ttu-id="46575-109"><span>PowerPoint</span></a></span><span class="sxs-lookup"><span data-stu-id="46575-109"><span>PowerPoint</span></a></span></span>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span data-ttu-id="46575-110"><span>Project</span></a></span><span class="sxs-lookup"><span data-stu-id="46575-110"><span>Project</span></a></span></span>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span data-ttu-id="46575-111"><span>Word</span></a></span><span class="sxs-lookup"><span data-stu-id="46575-111"><span>Word</span></a></span></span>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> <span data-ttu-id="46575-112">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="46575-112">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="46575-113">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="46575-113">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span> <span data-ttu-id="46575-114">作为[Office 云存储合作伙伴计划](https://developer.microsoft.com/office/cloud-storage-partner-program)成员的所有服务可能不支持 Office 加载项，这使集成 Office 网页版能够在其服务产品中使用 Office 文档。</span><span class="sxs-lookup"><span data-stu-id="46575-114">Office Add-ins may not be supported on all services that are members of the [Office Cloud Storage Partner Program](https://developer.microsoft.com/office/cloud-storage-partner-program), which enables integrating Office on the web to work with Office documents as part of their service offering.</span></span> <span data-ttu-id="46575-115">有关详细信息，请联系成员服务。</span><span class="sxs-lookup"><span data-stu-id="46575-115">For more information, please contact the member service.</span></span>

## <a name="excel"></a><span data-ttu-id="46575-116">Excel</span><span class="sxs-lookup"><span data-stu-id="46575-116">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="46575-117">平台</span><span class="sxs-lookup"><span data-stu-id="46575-117">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="46575-118">扩展点</span><span class="sxs-lookup"><span data-stu-id="46575-118">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="46575-119">API 要求集</span><span class="sxs-lookup"><span data-stu-id="46575-119">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="46575-120"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="46575-120"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-121">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="46575-121">Office on the web</span></span></td>
    <td><span data-ttu-id="46575-122">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-122">
      - TaskPane</span></span><br><span data-ttu-id="46575-123">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-123">
      - Content</span></span><br><span data-ttu-id="46575-124">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="46575-124">
      - CustomFunctions</span></span><br><span data-ttu-id="46575-125">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-125">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-126">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-126">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="46575-127">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-127">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="46575-128">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-128">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="46575-129">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-129">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="46575-130">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-130">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="46575-131">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-131">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="46575-132">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-132">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="46575-133">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-133">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="46575-134">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="46575-134">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="46575-135">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="46575-135">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="46575-136">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="46575-136">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="46575-137">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="46575-137">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="46575-138">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="46575-138">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="46575-139">
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="46575-139">
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a></span></span><br><span data-ttu-id="46575-140">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-140">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-141">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-141">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-142">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-142">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="46575-143">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-143">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-144">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-144">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-145">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-145">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-146">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-146">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-147">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-147">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-148">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-148">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-149">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-149">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-150">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-150">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-151">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-151">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-152">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-152">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-153">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-153">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-154">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-154">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-155">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-155">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-156">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-156">Office on Windows</span></span><br><span data-ttu-id="46575-157">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-157">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-158">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-158">
      - TaskPane</span></span><br><span data-ttu-id="46575-159">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-159">
      - Content</span></span><br><span data-ttu-id="46575-160">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="46575-160">
      - CustomFunctions</span></span><br><span data-ttu-id="46575-161">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-161">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-162">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-162">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="46575-163">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-163">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="46575-164">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-164">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="46575-165">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-165">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="46575-166">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-166">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="46575-167">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-167">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="46575-168">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-168">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="46575-169">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-169">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="46575-170">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="46575-170">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="46575-171">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="46575-171">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="46575-172">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="46575-172">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="46575-173">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="46575-173">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="46575-174">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="46575-174">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="46575-175">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-175">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-176">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-176">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-177">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-177">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-178">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-178">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="46575-179">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-179">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span></span><br><span data-ttu-id="46575-180">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-180">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="46575-181">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-181">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-182">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-182">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-183">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-183">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-184">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-184">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-185">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-185">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-186">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-186">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-187">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-187">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-188">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-188">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-189">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-189">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-190">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-190">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-191">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-191">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-192">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-192">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-193">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-193">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-194">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-194">Office 2019 on Windows</span></span><br><span data-ttu-id="46575-195">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-195">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-196">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-196">
      - TaskPane</span></span><br><span data-ttu-id="46575-197">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-197">
      - Content</span></span><br><span data-ttu-id="46575-198">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-198">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-199">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-199">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="46575-200">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-200">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="46575-201">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-201">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="46575-202">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-202">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="46575-203">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-203">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="46575-204">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-204">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="46575-205">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-205">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="46575-206">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-206">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="46575-207">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-207">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-208">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-208">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-209">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-209">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-210">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-210">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-211">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-211">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-212">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-212">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-213">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-213">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-214">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-214">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-215">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-215">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-216">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-216">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-217">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-217">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-218">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-218">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-219">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-219">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-220">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-220">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-221">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-221">Office 2016 on Windows</span></span><br><span data-ttu-id="46575-222">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-223">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-223">
      - TaskPane</span></span><br><span data-ttu-id="46575-224">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-224">
      - Content</span></span> </td>
    <td><span data-ttu-id="46575-225">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-225">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="46575-226">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-226">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-227">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-227">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-228">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-228">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-229">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-229">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-230">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-230">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-231">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-231">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-232">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-232">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-233">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-233">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-234">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-234">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-235">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-235">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-236">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-236">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-237">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-237">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-238">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-238">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-239">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-239">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-240">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="46575-240">Office 2013 on Windows</span></span><br><span data-ttu-id="46575-241">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-241">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-242">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-242">
      - TaskPane</span></span><br><span data-ttu-id="46575-243">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-243">
      - Content</span></span> </td>
    <td><span data-ttu-id="46575-244">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-244">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-245">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-245">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-246">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-246">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-247">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-247">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-248">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-248">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-249">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-249">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-250">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-250">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-251">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-251">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-252">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-252">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-253">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-253">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-254">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-254">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-255">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-255">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-256">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-256">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-257">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-257">Office on iPad</span></span><br><span data-ttu-id="46575-258">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-258">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-259">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-259">
      - TaskPane</span></span><br><span data-ttu-id="46575-260">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-260">
      - Content</span></span> </td>
    <td><span data-ttu-id="46575-261">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-261">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="46575-262">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-262">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="46575-263">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-263">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="46575-264">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-264">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="46575-265">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-265">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="46575-266">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-266">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="46575-267">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-267">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="46575-268">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-268">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="46575-269">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="46575-269">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="46575-270">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="46575-270">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="46575-271">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="46575-271">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="46575-272">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="46575-272">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="46575-273">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="46575-273">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="46575-274">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-274">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-275">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-275">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-276">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-276">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-277">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-277">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-278">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-278">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-279">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-279">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-280">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-280">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-281">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-281">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-282">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-282">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-283">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-283">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-284">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-284">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-285">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-285">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-286">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-286">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-287">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-287">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-288">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-288">Office on Mac</span></span><br><span data-ttu-id="46575-289">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-289">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-290">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-290">
      - TaskPane</span></span><br><span data-ttu-id="46575-291">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-291">
      - Content</span></span><br><span data-ttu-id="46575-292">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="46575-292">
      - CustomFunctions</span></span><br><span data-ttu-id="46575-293">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-293">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-294">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-294">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="46575-295">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-295">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="46575-296">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-296">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="46575-297">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-297">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="46575-298">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-298">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="46575-299">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-299">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="46575-300">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-300">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="46575-301">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-301">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="46575-302">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="46575-302">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="46575-303">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="46575-303">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="46575-304">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="46575-304">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="46575-305">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="46575-305">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="46575-306">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="46575-306">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="46575-307">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-307">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-308">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-308">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-309">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-309">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-310">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-310">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="46575-311">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-311">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span></span><br><span data-ttu-id="46575-312">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-312">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="46575-313">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-313">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-314">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-314">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-315">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-315">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-316">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-316">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-317">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-317">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-318">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-318">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-319">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-319">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-320">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-320">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-321">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-321">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-322">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-322">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-323">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-323">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-324">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-324">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-325">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-325">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-326">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-326">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-327">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-327">Office 2019 on Mac</span></span><br><span data-ttu-id="46575-328">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-329">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-329">
      - TaskPane</span></span><br><span data-ttu-id="46575-330">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-330">
      - Content</span></span><br><span data-ttu-id="46575-331">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-331">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-332">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-332">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="46575-333">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-333">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="46575-334">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-334">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="46575-335">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-335">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="46575-336">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-336">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="46575-337">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-337">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="46575-338">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-338">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="46575-339">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-339">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="46575-340">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-340">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-341">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-341">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-342">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-342">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-343">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-343">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-344">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-344">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-345">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-345">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-346">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-346">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-347">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-347">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-348">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-348">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-349">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-349">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-350">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-350">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-351">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-351">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-352">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-352">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-353">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-353">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-354">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-354">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-355">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-355">Office 2016 on Mac</span></span><br><span data-ttu-id="46575-356">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-356">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-357">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-357">
      - TaskPane</span></span><br><span data-ttu-id="46575-358">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-358">
      - Content</span></span> </td>
    <td><span data-ttu-id="46575-359">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-359">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="46575-360">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-360">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-361">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-361">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-362">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-362">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-363">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-363">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-364">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-364">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-365">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-365">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-366">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-366">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-367">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-367">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-368">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-368">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-369">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-369">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-370">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-370">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-371">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-371">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-372">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-372">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-373">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-373">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-374">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-374">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="46575-375">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="46575-375">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="46575-376">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="46575-376">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="46575-377">平台</span><span class="sxs-lookup"><span data-stu-id="46575-377">Platform</span></span></th>
    <th><span data-ttu-id="46575-378">扩展点</span><span class="sxs-lookup"><span data-stu-id="46575-378">Extension points</span></span></th>
    <th><span data-ttu-id="46575-379">API 要求集</span><span class="sxs-lookup"><span data-stu-id="46575-379">API requirement sets</span></span></th>
    <th><span data-ttu-id="46575-380"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="46575-380"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-381">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="46575-381">Office on the web</span></span></td>
    <td><span data-ttu-id="46575-382">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="46575-382">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="46575-383">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-383">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="46575-384">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-384">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="46575-385">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-385">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-386">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-386">Office on Windows</span></span><br><span data-ttu-id="46575-387">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-387">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-388">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="46575-388">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="46575-389">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-389">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="46575-390">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-390">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="46575-391">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-391">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-392">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-392">Office on Mac</span></span><br><span data-ttu-id="46575-393">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-393">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-394">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="46575-394">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="46575-395">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-395">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="46575-396">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-396">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="46575-397">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-397">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="46575-398">Outlook</span><span class="sxs-lookup"><span data-stu-id="46575-398">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="46575-399">平台</span><span class="sxs-lookup"><span data-stu-id="46575-399">Platform</span></span></th>
    <th><span data-ttu-id="46575-400">扩展点</span><span class="sxs-lookup"><span data-stu-id="46575-400">Extension points</span></span></th>
    <th><span data-ttu-id="46575-401">API 要求集</span><span class="sxs-lookup"><span data-stu-id="46575-401">API requirement sets</span></span></th>
    <th><span data-ttu-id="46575-402"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="46575-402"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-403">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="46575-403">Office on the web</span></span><br><span data-ttu-id="46575-404">（新式）</span><span class="sxs-lookup"><span data-stu-id="46575-404">(modern)</span></span></td>
    <td><span data-ttu-id="46575-405">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-405">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-406">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-406">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-407">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-407">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-408">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-408">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-409">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-409">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-410">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-410">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-411">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-411">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-412">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-412">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-413">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-413">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-414">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-414">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="46575-415">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-415">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="46575-416">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-416">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="46575-417">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-417">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="46575-418">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span><span class="sxs-lookup"><span data-stu-id="46575-418">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span></span><br><span data-ttu-id="46575-419">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">邮箱 1.10</a></span><span class="sxs-lookup"><span data-stu-id="46575-419">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a></span></span><br><span data-ttu-id="46575-420">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="46575-420">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="46575-421">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-421">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-422">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="46575-422">Office on the web</span></span><br><span data-ttu-id="46575-423">（经典）</span><span class="sxs-lookup"><span data-stu-id="46575-423">(classic)</span></span></td>
    <td><span data-ttu-id="46575-424">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-424">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-425">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-425">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-426">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-426">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-427">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-427">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-428">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-428">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-429">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-429">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-430">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-430">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-431">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-431">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-432">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-432">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-433">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-433">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="46575-434">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-434">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="46575-435">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-435">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-436">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-436">Office on Windows</span></span><br><span data-ttu-id="46575-437">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-437">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-438">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-438">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-439">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-439">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-440">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-440">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-441">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-441">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-442">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="46575-442">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="46575-443">
      - <a href="../reference/manifest/extensionpoint.md#module">模块</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-443">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="46575-444">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-444">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-445">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-445">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-446">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-446">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-447">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-447">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-448">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-448">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="46575-449">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-449">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="46575-450">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-450">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="46575-451">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-451">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="46575-452">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span><span class="sxs-lookup"><span data-stu-id="46575-452">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span></span><br><span data-ttu-id="46575-453">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">邮箱 1.10</a></span><span class="sxs-lookup"><span data-stu-id="46575-453">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a></span></span><br><span data-ttu-id="46575-454">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span><span class="sxs-lookup"><span data-stu-id="46575-454">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span></span><br><span data-ttu-id="46575-455">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-455">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-456">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-457">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-457">Office 2019 on Windows</span></span><br><span data-ttu-id="46575-458">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-458">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-459">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-459">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-460">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-460">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-461">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-461">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-462">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-462">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-463">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="46575-463">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="46575-464">
      - <a href="../reference/manifest/extensionpoint.md#module">模块</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-464">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="46575-465">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-465">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-466">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-466">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-467">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-467">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-468">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-468">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-469">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-469">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="46575-470">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-470">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="46575-471">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-471">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="46575-472">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-472">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-473">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-473">Office 2016 on Windows</span></span><br><span data-ttu-id="46575-474">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-474">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-475">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-475">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-476">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-476">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-477">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-477">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-478">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-478">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-479">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="46575-479">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="46575-480">
      - <a href="../reference/manifest/extensionpoint.md#module">模块</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-480">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="46575-481">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-481">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-482">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-482">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-483">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-483">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-484">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span><span class="sxs-lookup"><span data-stu-id="46575-484">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span></span></td>
    <td><span data-ttu-id="46575-485">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-485">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-486">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="46575-486">Office 2013 on Windows</span></span><br><span data-ttu-id="46575-487">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-487">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-488">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-488">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-489">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-489">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-490">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-490">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-491">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-491">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    </td>
    <td><span data-ttu-id="46575-492">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-492">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-493">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-493">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-494">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><sup>2</sup></span><span class="sxs-lookup"><span data-stu-id="46575-494">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><sup>2</sup></span></span><br><span data-ttu-id="46575-495">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span><span class="sxs-lookup"><span data-stu-id="46575-495">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span></span></td>
    <td><span data-ttu-id="46575-496">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-497">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-497">Office on iOS</span></span><br><span data-ttu-id="46575-498">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-498">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-499">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-499">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-500">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">约会组织者（撰写）：联机会议</a></span><span class="sxs-lookup"><span data-stu-id="46575-500">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a></span></span><br><span data-ttu-id="46575-501">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-501">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-502">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-502">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-503">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-503">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-504">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-504">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-505">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-505">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-506">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-506">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="46575-507">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-507">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-508">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-508">Office on Mac</span></span><br><span data-ttu-id="46575-509">（当前 UI，</span><span class="sxs-lookup"><span data-stu-id="46575-509">(current UI,</span></span><br><span data-ttu-id="46575-510">连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-510">connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-511">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-511">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-512">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-512">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-513">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-513">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-514">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-514">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-515">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-515">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-516">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-516">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-517">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-517">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-518">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-518">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-519">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-519">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-520">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-520">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="46575-521">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-521">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="46575-522">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-522">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="46575-523">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-523">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="46575-524">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span><span class="sxs-lookup"><span data-stu-id="46575-524">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span></span><br><span data-ttu-id="46575-525">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-525">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-526">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-526">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-527">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-527">Office on Mac</span></span><br><span data-ttu-id="46575-528">（全新 UI（预览版）<sup>3</sup></span><span class="sxs-lookup"><span data-stu-id="46575-528">(new UI (preview)<sup>3</sup>,</span></span><br><span data-ttu-id="46575-529">连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-529">connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-530">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-530">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-531">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-531">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-532">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-532">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-533">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-533">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-534">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-534">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-535">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-535">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-536">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-536">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-537">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-537">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-538">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-538">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-539">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-539">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="46575-540">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="46575-540">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="46575-541">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="46575-541">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="46575-542">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="46575-542">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="46575-543">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="46575-543">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="46575-544">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-544">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-545">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-545">Office 2019 on Mac</span></span><br><span data-ttu-id="46575-546">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-546">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-547">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-547">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-548">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-548">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-549">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-549">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-550">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-550">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-551">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-551">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-552">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-552">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-553">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-553">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-554">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-554">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-555">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-555">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-556">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-556">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="46575-557">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-557">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="46575-558">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-558">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-559">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-559">Office 2016 on Mac</span></span><br><span data-ttu-id="46575-560">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-560">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-561">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-561">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-562">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="46575-562">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="46575-563">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="46575-563">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="46575-564">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="46575-564">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="46575-565">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-565">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-566">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-566">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-567">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-567">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-568">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-568">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-569">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-569">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-570">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="46575-570">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="46575-571">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-571">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="46575-572">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-572">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-573">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-573">Office on Android</span></span><br><span data-ttu-id="46575-574">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-574">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-575">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="46575-575">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="46575-576">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">约会组织者（撰写）：联机会议</a></span><span class="sxs-lookup"><span data-stu-id="46575-576">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a></span></span><br><span data-ttu-id="46575-577">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-577">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-578">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-578">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="46575-579">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-579">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="46575-580">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-580">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="46575-581">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="46575-581">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="46575-582">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-582">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="46575-583">不可用</span><span class="sxs-lookup"><span data-stu-id="46575-583">Not available</span></span></td>
  </tr>
</table>

> [!NOTE]
> <span data-ttu-id="46575-584"><sup>1</sup> 若要在加载项代码中要求标识 API 集 1.3，请通过呼叫 `isSetSupported('IdentityAPI', '1.3')` 检查其是否收到支持。</span><span class="sxs-lookup"><span data-stu-id="46575-584"><sup>1</sup> To require Identity API set 1.3 in your add-in code, check if it's supported by calling `isSetSupported('IdentityAPI', '1.3')`.</span></span> <span data-ttu-id="46575-585">声明其在加载项清单中不受支持。</span><span class="sxs-lookup"><span data-stu-id="46575-585">Declaring it in your add-in's manifest isn't supported.</span></span> <span data-ttu-id="46575-586">还可通过检查其不是 `undefined` 来确定该 API 是否受到支持。</span><span class="sxs-lookup"><span data-stu-id="46575-586">You can also determine if the API is supported by checking that it's not `undefined`.</span></span> <span data-ttu-id="46575-587">有关详细信息，请参阅 [从后续要求集中使用 API](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="46575-587">For further details, see [Using APIs from later requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span></span>
>
> <span data-ttu-id="46575-588"><sup>2</sup> 已添加发布后更新。</span><span class="sxs-lookup"><span data-stu-id="46575-588"><sup>2</sup> Added with post-release updates.</span></span>
>
> <span data-ttu-id="46575-589"><sup>3</sup> Outlook 版本 16.38.506 已提供对全新 Mac UI（预览版）的支持。</span><span class="sxs-lookup"><span data-stu-id="46575-589"><sup>3</sup> Support for the new Mac UI (preview) is available from Outlook version 16.38.506.</span></span> <span data-ttu-id="46575-590">有关详细信息，请参阅 [全新 Mac UI 上 Outlook 中的加载项支持](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview) 部分。</span><span class="sxs-lookup"><span data-stu-id="46575-590">For more information, see the [Add-in support in Outlook on new Mac UI](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview) section.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="46575-591">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="46575-591">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="46575-592">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="46575-592">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="46575-593">Word</span><span class="sxs-lookup"><span data-stu-id="46575-593">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="46575-594">平台</span><span class="sxs-lookup"><span data-stu-id="46575-594">Platform</span></span></th>
    <th><span data-ttu-id="46575-595">扩展点</span><span class="sxs-lookup"><span data-stu-id="46575-595">Extension points</span></span></th>
    <th><span data-ttu-id="46575-596">API 要求集</span><span class="sxs-lookup"><span data-stu-id="46575-596">API requirement sets</span></span></th>
    <th><span data-ttu-id="46575-597"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="46575-597"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-598">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="46575-598">Office on the web</span></span></td>
    <td><span data-ttu-id="46575-599">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-599">
      - TaskPane</span></span><br><span data-ttu-id="46575-600">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-600">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-601">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-601">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="46575-602">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-602">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="46575-603">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-603">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="46575-604">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-604">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-605">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-605">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-606">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-606">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-607">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-607">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-608">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-608">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-609">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-609">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-610">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-610">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-611">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-611">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-612">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-612">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-613">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-613">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-614">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-614">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-615">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-615">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-616">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-616">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-617">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-617">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-618">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-618">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-619">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-619">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-620">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-620">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-621">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-621">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-622">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-622">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-623">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-623">Office on Windows</span></span><br><span data-ttu-id="46575-624">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-624">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-625">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-625">
      - TaskPane</span></span><br><span data-ttu-id="46575-626">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-626">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-627">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-627">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="46575-628">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-628">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="46575-629">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-629">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="46575-630">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-630">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-631">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-631">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-632">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-632">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-633">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-633">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="46575-634">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-634">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-635">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-635">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-636">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-636">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-637">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-637">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-638">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-638">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-639">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-639">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-640">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-640">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-641">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-641">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-642">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-642">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-643">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-643">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-644">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-644">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-645">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-645">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-646">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-646">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-647">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-647">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-648">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-648">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-649">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-649">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-650">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-650">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-651">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-651">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-652">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-652">Office 2019 on Windows</span></span><br><span data-ttu-id="46575-653">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-653">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-654">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-654">
      - TaskPane</span></span><br><span data-ttu-id="46575-655">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-655">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-656">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-656">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="46575-657">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-657">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="46575-658">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-658">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="46575-659">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-659">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-660">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-660">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-661">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-661">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-662">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-662">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-663">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-663">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-664">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-664">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-665">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-665">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-666">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-666">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-667">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-667">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-668">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-668">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-669">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-669">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-670">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-670">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-671">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-671">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-672">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-672">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-673">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-673">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-674">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-674">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-675">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-675">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-676">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-676">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-677">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-677">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-678">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-678">Office 2016 on Windows</span></span><br><span data-ttu-id="46575-679">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-679">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-680">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-680">- TaskPane</span></span></td>
    <td><span data-ttu-id="46575-681">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-681">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="46575-682">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-682">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-683">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-683">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-684">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-684">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-685">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-685">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-686">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-686">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-687">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-687">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-688">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-688">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-689">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-689">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-690">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-690">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-691">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-691">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-692">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-692">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-693">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-693">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-694">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-694">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-695">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-695">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-696">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-696">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-697">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-697">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-698">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-698">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-699">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-699">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-700">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-700">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-701">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="46575-701">Office 2013 on Windows</span></span><br><span data-ttu-id="46575-702">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-702">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-703">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-703">- TaskPane</span></span></td>
    <td><span data-ttu-id="46575-704">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-704">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-705">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-705">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-706">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-706">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-707">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-707">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-708">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-708">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-709">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-709">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-710">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-710">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-711">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-711">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-712">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-712">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-713">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-713">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-714">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-714">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-715">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-715">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-716">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-716">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-717">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-717">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-718">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-718">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-719">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-719">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-720">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-720">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-721">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-721">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-722">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-722">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-723">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-723">Office on iPad</span></span><br><span data-ttu-id="46575-724">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-724">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-725">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-725">- TaskPane</span></span></td>
    <td><span data-ttu-id="46575-726">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-726">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="46575-727">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-727">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="46575-728">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-728">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="46575-729">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-729">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-730">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-730">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-731">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-731">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-732">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-732">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-733">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-733">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-734">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-734">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-735">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-735">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-736">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-736">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-737">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-737">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-738">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-738">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-739">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-739">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-740">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-740">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-741">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-741">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-742">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-742">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-743">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-743">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-744">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-744">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-745">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-745">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-746">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-746">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-747">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-747">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-748">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-748">Office on Mac</span></span><br><span data-ttu-id="46575-749">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-749">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-750">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-750">
      - TaskPane</span></span><br><span data-ttu-id="46575-751">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-751">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-752">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-752">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="46575-753">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-753">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="46575-754">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-754">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="46575-755">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-755">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-756">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-756">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-757">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-757">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-758">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-758">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="46575-759">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-759">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-760">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-760">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-761">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-761">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-762">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-762">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-763">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-763">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-764">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-764">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-765">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-765">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-766">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-766">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-767">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-767">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-768">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-768">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-769">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-769">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-770">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-770">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-771">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-771">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-772">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-772">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-773">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-773">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-774">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-774">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-775">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-775">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-776">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-776">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-777">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-777">Office 2019 on Mac</span></span><br><span data-ttu-id="46575-778">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-778">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-779">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-779">
      - TaskPane</span></span><br><span data-ttu-id="46575-780">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-780">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-781">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-781">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="46575-782">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-782">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="46575-783">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-783">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="46575-784">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-784">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-785">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-785">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-786">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-786">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-787">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-787">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-788">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-788">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-789">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-789">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-790">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-790">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-791">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-791">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-792">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-792">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-793">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-793">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-794">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-794">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-795">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-795">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-796">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-796">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-797">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-797">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-798">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-798">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-799">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-799">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-800">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-800">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-801">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-801">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-802">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-802">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-803">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-803">Office 2016 on Mac</span></span><br><span data-ttu-id="46575-804">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-804">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-805">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-805">- TaskPane</span></span></td>
    <td><span data-ttu-id="46575-806">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-806">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="46575-807">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-807">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-808">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-808">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-809">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-809">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="46575-810">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-810">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-811">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="46575-811">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="46575-812">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-812">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-813">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-813">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-814">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-814">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-815">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-815">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="46575-816">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-816">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="46575-817">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-817">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="46575-818">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-818">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-819">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-819">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-820">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">设置</a></span><span class="sxs-lookup"><span data-stu-id="46575-820">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-821">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-821">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="46575-822">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-822">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="46575-823">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="46575-823">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="46575-824">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-824">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="46575-825">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-825">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="46575-826">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="46575-826">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="46575-827">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="46575-827">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="46575-828">平台</span><span class="sxs-lookup"><span data-stu-id="46575-828">Platform</span></span></th>
    <th><span data-ttu-id="46575-829">扩展点</span><span class="sxs-lookup"><span data-stu-id="46575-829">Extension points</span></span></th>
    <th><span data-ttu-id="46575-830">API 要求集</span><span class="sxs-lookup"><span data-stu-id="46575-830">API requirement sets</span></span></th>
    <th><span data-ttu-id="46575-831"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="46575-831"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-832">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="46575-832">Office on the web</span></span></td>
    <td><span data-ttu-id="46575-833">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-833">
      - Content</span></span><br><span data-ttu-id="46575-834">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-834">
      - TaskPane</span></span><br><span data-ttu-id="46575-835">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-835">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-836">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-836">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="46575-837">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-837">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="46575-838">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-838">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-839">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-839">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-840">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-840">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-841">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-841">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="46575-842">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-842">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-843">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-843">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-844">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-844">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-845">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-845">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-846">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-846">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-847">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-847">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-848">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-848">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-849">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-849">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-850">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-850">Office on Windows</span></span><br><span data-ttu-id="46575-851">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-851">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-852">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-852">
      - Content</span></span><br><span data-ttu-id="46575-853">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-853">
      - TaskPane</span></span><br><span data-ttu-id="46575-854">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-854">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-855">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-855">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="46575-856">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-856">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="46575-857">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-857">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-858">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-858">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-859">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-859">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-860">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-860">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="46575-861">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-861">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-862">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-862">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-863">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-863">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-864">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-864">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-865">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-865">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-866">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-866">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-867">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-867">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-868">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-868">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-869">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-869">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-870">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-870">Office 2019 on Windows</span></span><br><span data-ttu-id="46575-871">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-871">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-872">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-872">
      - Content</span></span><br><span data-ttu-id="46575-873">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-873">
      - TaskPane</span></span><br><span data-ttu-id="46575-874">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-874">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-875">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-875">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-876">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-876">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-877">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-877">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-878">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-878">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-879">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-879">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-880">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-880">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-881">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-881">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-882">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-882">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-883">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-883">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-884">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-884">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-885">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-885">Office 2016 on Windows</span></span><br><span data-ttu-id="46575-886">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-886">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-887">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-887">
      - Content</span></span><br><span data-ttu-id="46575-888">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-888">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="46575-889">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-889">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-890">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-890">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-891">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-891">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-892">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-892">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-893">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-893">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-894">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-894">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-895">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-895">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-896">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-896">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-897">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-897">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-898">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-898">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-899">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="46575-899">Office 2013 on Windows</span></span><br><span data-ttu-id="46575-900">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-900">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-901">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-901">
      - Content</span></span><br><span data-ttu-id="46575-902">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-902">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="46575-903">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-903">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-904">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-904">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-905">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-905">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-906">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-906">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-907">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-907">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-908">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-908">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-909">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-909">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-910">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-910">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-911">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-911">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-912">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-912">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-913">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-913">Office on iPad</span></span><br><span data-ttu-id="46575-914">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-914">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-915">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-915">
      - Content</span></span><br><span data-ttu-id="46575-916">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-916">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="46575-917">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-917">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="46575-918">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-918">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="46575-919">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-919">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-920">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-920">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-921">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-921">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-922">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-922">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-923">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-923">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-924">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-924">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-925">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-925">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-926">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-926">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-927">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-927">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-928">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-928">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-929">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-929">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-930">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="46575-930">Office on Mac</span></span><br><span data-ttu-id="46575-931">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="46575-931">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="46575-932">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-932">
      - Content</span></span><br><span data-ttu-id="46575-933">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-933">
      - TaskPane</span></span><br><span data-ttu-id="46575-934">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-934">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-935">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-935">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="46575-936">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-936">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="46575-937">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-937">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-938">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="46575-938">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="46575-939">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-939">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="46575-940">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="46575-940">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="46575-941">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-941">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-942">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-942">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-943">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-943">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-944">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-944">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-945">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-945">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-946">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-946">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-947">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-947">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-948">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-948">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-949">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-949">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-950">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-950">Office 2019 on Mac</span></span><br><span data-ttu-id="46575-951">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-951">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-952">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-952">
      - Content</span></span><br><span data-ttu-id="46575-953">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-953">
      - TaskPane</span></span><br><span data-ttu-id="46575-954">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-954">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-955">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-955">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-956">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-956">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-957">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-957">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-958">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-958">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-959">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-959">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-960">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-960">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-961">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-961">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-962">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-962">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-963">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-963">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-964">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-964">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-965">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-965">Office 2016 on Mac</span></span><br><span data-ttu-id="46575-966">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-966">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-967">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-967">
      - Content</span></span><br><span data-ttu-id="46575-968">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-968">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="46575-969">
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="46575-969">
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="46575-970">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-970">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-971">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="46575-971">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="46575-972">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-972">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="46575-973">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-973">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-974">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span><span class="sxs-lookup"><span data-stu-id="46575-974">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="46575-975">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="46575-975">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="46575-976">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-976">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-977">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-977">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-978">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-978">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="46575-979">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="46575-979">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="46575-980">OneNote</span><span class="sxs-lookup"><span data-stu-id="46575-980">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="46575-981">平台</span><span class="sxs-lookup"><span data-stu-id="46575-981">Platform</span></span></th>
    <th><span data-ttu-id="46575-982">扩展点</span><span class="sxs-lookup"><span data-stu-id="46575-982">Extension points</span></span></th>
    <th><span data-ttu-id="46575-983">API 要求集</span><span class="sxs-lookup"><span data-stu-id="46575-983">API requirement sets</span></span></th>
    <th><span data-ttu-id="46575-984"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="46575-984"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-985">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="46575-985">Office on the web</span></span></td>
    <td><span data-ttu-id="46575-986">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="46575-986">
      - Content</span></span><br><span data-ttu-id="46575-987">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-987">
      - TaskPane</span></span><br><span data-ttu-id="46575-988">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-988">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="46575-989">
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-989">
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="46575-990">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-990">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="46575-991">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-991">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="46575-992">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="46575-992">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="46575-993">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="46575-993">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="46575-994">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="46575-994">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="46575-995">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-995">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="46575-996">Project</span><span class="sxs-lookup"><span data-stu-id="46575-996">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="46575-997">平台</span><span class="sxs-lookup"><span data-stu-id="46575-997">Platform</span></span></th>
    <th><span data-ttu-id="46575-998">扩展点</span><span class="sxs-lookup"><span data-stu-id="46575-998">Extension points</span></span></th>
    <th><span data-ttu-id="46575-999">API 要求集</span><span class="sxs-lookup"><span data-stu-id="46575-999">API requirement sets</span></span></th>
    <th><span data-ttu-id="46575-1000"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="46575-1000"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-1001">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="46575-1001">Office 2019 on Windows</span></span><br><span data-ttu-id="46575-1002">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-1002">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-1003">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-1003">- TaskPane</span></span></td>
    <td><span data-ttu-id="46575-1004">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-1004">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="46575-1005">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-1005">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-1006">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-1006">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-1007">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="46575-1007">Office 2016 on Windows</span></span><br><span data-ttu-id="46575-1008">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-1008">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-1009">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-1009">- TaskPane</span></span></td>
    <td><span data-ttu-id="46575-1010">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-1010">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="46575-1011">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-1011">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-1012">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-1012">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="46575-1013">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="46575-1013">Office 2013 on Windows</span></span><br><span data-ttu-id="46575-1014">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="46575-1014">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="46575-1015">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="46575-1015">- TaskPane</span></span></td>
    <td><span data-ttu-id="46575-1016">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="46575-1016">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="46575-1017">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="46575-1017">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="46575-1018">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="46575-1018">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="46575-1019">另请参阅</span><span class="sxs-lookup"><span data-stu-id="46575-1019">See also</span></span>

- [<span data-ttu-id="46575-1020">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="46575-1020">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="46575-1021">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="46575-1021">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="46575-1022">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="46575-1022">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="46575-1023">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="46575-1023">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="46575-1024">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="46575-1024">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="46575-1025">Microsoft 365 应用版的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="46575-1025">Update history for Microsoft 365 Apps</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="46575-1026">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="46575-1026">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="46575-1027">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="46575-1027">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="46575-1028">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="46575-1028">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="46575-1029">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="46575-1029">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="46575-1030">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="46575-1030">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="46575-1031">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="46575-1031">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
