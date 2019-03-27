---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 28a6d0e4c86d05855ed9d24461dbeb77454d2b48
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872128"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="a032a-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="a032a-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="a032a-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="a032a-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="a032a-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="a032a-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="a032a-108">Office 2019 的一次性购买的内部版本号是 16.0.10827.20150。</span><span class="sxs-lookup"><span data-stu-id="a032a-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="a032a-109">Excel</span><span class="sxs-lookup"><span data-stu-id="a032a-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="a032a-110">平台</span><span class="sxs-lookup"><span data-stu-id="a032a-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="a032a-111">扩展点</span><span class="sxs-lookup"><span data-stu-id="a032a-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="a032a-112">API 要求集</span><span class="sxs-lookup"><span data-stu-id="a032a-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="a032a-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a032a-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="a032a-114">Office Online</span></span></td>
    <td> <span data-ttu-id="a032a-115">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-115">- TaskPane</span></span><br><span data-ttu-id="a032a-116">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-116">
        - Content</span></span><br><span data-ttu-id="a032a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="a032a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="a032a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a032a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a032a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a032a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a032a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a032a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a032a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a032a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a032a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a032a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a032a-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-127">
        - BindingEvents</span></span><br><span data-ttu-id="a032a-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-128">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-129">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-130">
        - File</span></span><br><span data-ttu-id="a032a-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-131">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-133">
        - Selection</span></span><br><span data-ttu-id="a032a-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-134">
        - Settings</span></span><br><span data-ttu-id="a032a-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-135">
        - TableBindings</span></span><br><span data-ttu-id="a032a-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-136">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-137">
        - TextBindings</span></span><br><span data-ttu-id="a032a-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-139">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-140">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-140">- TaskPane</span></span><br><span data-ttu-id="a032a-141">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-141">
        - Content</span></span><br><span data-ttu-id="a032a-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="a032a-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="a032a-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a032a-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a032a-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a032a-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a032a-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a032a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a032a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a032a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a032a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a032a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a032a-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-152">
        - BindingEvents</span></span><br><span data-ttu-id="a032a-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-153">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-154">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-155">
        - File</span></span><br><span data-ttu-id="a032a-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-156">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-158">
        - Selection</span></span><br><span data-ttu-id="a032a-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-159">
        - Settings</span></span><br><span data-ttu-id="a032a-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-160">
        - TableBindings</span></span><br><span data-ttu-id="a032a-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-161">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-162">
        - TextBindings</span></span><br><span data-ttu-id="a032a-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-164">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="a032a-165">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-165">- TaskPane</span></span><br><span data-ttu-id="a032a-166">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-166">
        - Content</span></span><br><span data-ttu-id="a032a-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a032a-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a032a-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a032a-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a032a-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a032a-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a032a-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a032a-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a032a-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a032a-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a032a-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a032a-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-177">- BindingEvents</span></span><br><span data-ttu-id="a032a-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-178">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-179">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-180">
        - File</span></span><br><span data-ttu-id="a032a-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-181">
        - ImageCoercion</span></span><br><span data-ttu-id="a032a-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-182">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-184">
        - Selection</span></span><br><span data-ttu-id="a032a-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-185">
        - Settings</span></span><br><span data-ttu-id="a032a-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-186">
        - TableBindings</span></span><br><span data-ttu-id="a032a-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-187">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-188">
        - TextBindings</span></span><br><span data-ttu-id="a032a-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-190">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="a032a-191">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-191">- TaskPane</span></span><br><span data-ttu-id="a032a-192">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-192">
        - Content</span></span></td>
    <td><span data-ttu-id="a032a-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a032a-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a032a-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="a032a-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-195">- BindingEvents</span></span><br><span data-ttu-id="a032a-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-196">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-197">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-198">
        - File</span></span><br><span data-ttu-id="a032a-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-199">
        - ImageCoercion</span></span><br><span data-ttu-id="a032a-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-200">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-202">
        - Selection</span></span><br><span data-ttu-id="a032a-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-203">
        - Settings</span></span><br><span data-ttu-id="a032a-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-204">
        - TableBindings</span></span><br><span data-ttu-id="a032a-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-205">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-206">
        - TextBindings</span></span><br><span data-ttu-id="a032a-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-208">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="a032a-209">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-209">
        - TaskPane</span></span><br><span data-ttu-id="a032a-210">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="a032a-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a032a-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="a032a-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-212">
        - BindingEvents</span></span><br><span data-ttu-id="a032a-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-213">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-214">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-215">
        - File</span></span><br><span data-ttu-id="a032a-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-216">
        - ImageCoercion</span></span><br><span data-ttu-id="a032a-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-217">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-219">
        - Selection</span></span><br><span data-ttu-id="a032a-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-220">
        - Settings</span></span><br><span data-ttu-id="a032a-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-221">
        - TableBindings</span></span><br><span data-ttu-id="a032a-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-222">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-223">
        - TextBindings</span></span><br><span data-ttu-id="a032a-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-225">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="a032a-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="a032a-226">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-226">- TaskPane</span></span><br><span data-ttu-id="a032a-227">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-227">
        - Content</span></span></td>
    <td><span data-ttu-id="a032a-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a032a-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a032a-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a032a-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a032a-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a032a-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a032a-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a032a-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a032a-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a032a-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a032a-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-237">- BindingEvents</span></span><br><span data-ttu-id="a032a-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-238">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-239">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-240">
        - File</span></span><br><span data-ttu-id="a032a-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-241">
        - ImageCoercion</span></span><br><span data-ttu-id="a032a-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-242">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-244">
        - Selection</span></span><br><span data-ttu-id="a032a-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-245">
        - Settings</span></span><br><span data-ttu-id="a032a-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-246">
        - TableBindings</span></span><br><span data-ttu-id="a032a-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-247">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-248">
        - TextBindings</span></span><br><span data-ttu-id="a032a-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-250">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="a032a-251">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-251">- TaskPane</span></span><br><span data-ttu-id="a032a-252">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-252">
        - Content</span></span><br><span data-ttu-id="a032a-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a032a-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a032a-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a032a-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a032a-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a032a-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a032a-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a032a-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a032a-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a032a-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a032a-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a032a-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-263">- BindingEvents</span></span><br><span data-ttu-id="a032a-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-264">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-265">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-266">
        - File</span></span><br><span data-ttu-id="a032a-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-267">
        - ImageCoercion</span></span><br><span data-ttu-id="a032a-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-268">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-270">
        - PdfFile</span></span><br><span data-ttu-id="a032a-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-271">
        - Selection</span></span><br><span data-ttu-id="a032a-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-272">
        - Settings</span></span><br><span data-ttu-id="a032a-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-273">
        - TableBindings</span></span><br><span data-ttu-id="a032a-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-274">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-275">
        - TextBindings</span></span><br><span data-ttu-id="a032a-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-277">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="a032a-278">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-278">- TaskPane</span></span><br><span data-ttu-id="a032a-279">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-279">
        - Content</span></span><br><span data-ttu-id="a032a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a032a-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a032a-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a032a-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a032a-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a032a-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a032a-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a032a-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a032a-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a032a-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a032a-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a032a-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-290">- BindingEvents</span></span><br><span data-ttu-id="a032a-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-291">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-292">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-293">
        - File</span></span><br><span data-ttu-id="a032a-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-294">
        - ImageCoercion</span></span><br><span data-ttu-id="a032a-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-295">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-297">
        - PdfFile</span></span><br><span data-ttu-id="a032a-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-298">
        - Selection</span></span><br><span data-ttu-id="a032a-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-299">
        - Settings</span></span><br><span data-ttu-id="a032a-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-300">
        - TableBindings</span></span><br><span data-ttu-id="a032a-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-301">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-302">
        - TextBindings</span></span><br><span data-ttu-id="a032a-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-304">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="a032a-305">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-305">- TaskPane</span></span><br><span data-ttu-id="a032a-306">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-306">
        - Content</span></span></td>
    <td><span data-ttu-id="a032a-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a032a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a032a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="a032a-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-309">- BindingEvents</span></span><br><span data-ttu-id="a032a-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-310">
        - CompressedFile</span></span><br><span data-ttu-id="a032a-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-311">
        - DocumentEvents</span></span><br><span data-ttu-id="a032a-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="a032a-312">
        - File</span></span><br><span data-ttu-id="a032a-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-313">
        - ImageCoercion</span></span><br><span data-ttu-id="a032a-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-314">
        - MatrixBindings</span></span><br><span data-ttu-id="a032a-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="a032a-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-316">
        - PdfFile</span></span><br><span data-ttu-id="a032a-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-317">
        - Selection</span></span><br><span data-ttu-id="a032a-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-318">
        - Settings</span></span><br><span data-ttu-id="a032a-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-319">
        - TableBindings</span></span><br><span data-ttu-id="a032a-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-320">
        - TableCoercion</span></span><br><span data-ttu-id="a032a-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-321">
        - TextBindings</span></span><br><span data-ttu-id="a032a-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="a032a-323">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="a032a-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="a032a-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="a032a-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a032a-325">平台</span><span class="sxs-lookup"><span data-stu-id="a032a-325">Platform</span></span></th>
    <th><span data-ttu-id="a032a-326">扩展点</span><span class="sxs-lookup"><span data-stu-id="a032a-326">Extension points</span></span></th>
    <th><span data-ttu-id="a032a-327">API 要求集</span><span class="sxs-lookup"><span data-stu-id="a032a-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="a032a-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a032a-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="a032a-329">Office Online</span></span></td>
    <td> <span data-ttu-id="a032a-330">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-330">- Mail Read</span></span><br><span data-ttu-id="a032a-331">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="a032a-331">
      - Mail Compose</span></span><br><span data-ttu-id="a032a-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a032a-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a032a-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a032a-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a032a-340">不可用</span><span class="sxs-lookup"><span data-stu-id="a032a-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-341">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-342">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-342">- Mail Read</span></span><br><span data-ttu-id="a032a-343">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="a032a-343">
      - Mail Compose</span></span><br><span data-ttu-id="a032a-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a032a-345">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="a032a-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a032a-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a032a-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a032a-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a032a-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a032a-353">不可用</span><span class="sxs-lookup"><span data-stu-id="a032a-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-354">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-355">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-355">- Mail Read</span></span><br><span data-ttu-id="a032a-356">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="a032a-356">
      - Mail Compose</span></span><br><span data-ttu-id="a032a-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a032a-358">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="a032a-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a032a-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a032a-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a032a-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a032a-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a032a-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a032a-366">不可用</span><span class="sxs-lookup"><span data-stu-id="a032a-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-367">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-368">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-368">- Mail Read</span></span><br><span data-ttu-id="a032a-369">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="a032a-369">
      - Mail Compose</span></span><br><span data-ttu-id="a032a-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a032a-371">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="a032a-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a032a-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="a032a-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="a032a-376">暂无</span><span class="sxs-lookup"><span data-stu-id="a032a-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-377">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-378">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-378">- Mail Read</span></span><br><span data-ttu-id="a032a-379">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="a032a-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="a032a-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="a032a-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="a032a-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="a032a-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="a032a-384">暂无</span><span class="sxs-lookup"><span data-stu-id="a032a-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-385">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="a032a-385">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="a032a-386">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-386">- Mail Read</span></span><br><span data-ttu-id="a032a-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a032a-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="a032a-393">不可用</span><span class="sxs-lookup"><span data-stu-id="a032a-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-394">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-395">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-395">- Mail Read</span></span><br><span data-ttu-id="a032a-396">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="a032a-396">
      - Mail Compose</span></span><br><span data-ttu-id="a032a-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a032a-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a032a-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a032a-404">不可用</span><span class="sxs-lookup"><span data-stu-id="a032a-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-405">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-406">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-406">- Mail Read</span></span><br><span data-ttu-id="a032a-407">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="a032a-407">
      - Mail Compose</span></span><br><span data-ttu-id="a032a-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a032a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a032a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a032a-415">不可用</span><span class="sxs-lookup"><span data-stu-id="a032a-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-416">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-417">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-417">- Mail Read</span></span><br><span data-ttu-id="a032a-418">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="a032a-418">
      - Mail Compose</span></span><br><span data-ttu-id="a032a-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a032a-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a032a-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a032a-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a032a-426">不可用</span><span class="sxs-lookup"><span data-stu-id="a032a-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-427">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="a032a-427">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="a032a-428">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="a032a-428">- Mail Read</span></span><br><span data-ttu-id="a032a-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a032a-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a032a-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a032a-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a032a-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a032a-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a032a-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="a032a-435">不可用</span><span class="sxs-lookup"><span data-stu-id="a032a-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="a032a-436">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="a032a-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="a032a-437">Word</span><span class="sxs-lookup"><span data-stu-id="a032a-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a032a-438">平台</span><span class="sxs-lookup"><span data-stu-id="a032a-438">Platform</span></span></th>
    <th><span data-ttu-id="a032a-439">扩展点</span><span class="sxs-lookup"><span data-stu-id="a032a-439">Extension points</span></span></th>
    <th><span data-ttu-id="a032a-440">API 要求集</span><span class="sxs-lookup"><span data-stu-id="a032a-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="a032a-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a032a-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="a032a-442">Office Online</span></span></td>
    <td> <span data-ttu-id="a032a-443">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-443">- TaskPane</span></span><br><span data-ttu-id="a032a-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a032a-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a032a-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a032a-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-449">- BindingEvents</span></span><br><span data-ttu-id="a032a-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-451">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-452">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-452">
         - File</span></span><br><span data-ttu-id="a032a-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-454">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-455">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-458">
         - PdfFile</span></span><br><span data-ttu-id="a032a-459">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-459">
         - Selection</span></span><br><span data-ttu-id="a032a-460">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-460">
         - Settings</span></span><br><span data-ttu-id="a032a-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-461">
         - TableBindings</span></span><br><span data-ttu-id="a032a-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-462">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-463">
         - TextBindings</span></span><br><span data-ttu-id="a032a-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-464">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-466">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-467">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-467">- TaskPane</span></span><br><span data-ttu-id="a032a-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a032a-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a032a-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a032a-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-473">- BindingEvents</span></span><br><span data-ttu-id="a032a-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-474">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-476">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-477">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-477">
         - File</span></span><br><span data-ttu-id="a032a-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-479">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-480">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-483">
         - PdfFile</span></span><br><span data-ttu-id="a032a-484">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-484">
         - Selection</span></span><br><span data-ttu-id="a032a-485">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-485">
         - Settings</span></span><br><span data-ttu-id="a032a-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-486">
         - TableBindings</span></span><br><span data-ttu-id="a032a-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-487">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-488">
         - TextBindings</span></span><br><span data-ttu-id="a032a-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-489">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-491">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-492">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-492">- TaskPane</span></span><br><span data-ttu-id="a032a-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a032a-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a032a-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a032a-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-498">- BindingEvents</span></span><br><span data-ttu-id="a032a-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-499">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-501">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-502">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-502">
         - File</span></span><br><span data-ttu-id="a032a-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-504">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-505">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-508">
         - PdfFile</span></span><br><span data-ttu-id="a032a-509">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-509">
         - Selection</span></span><br><span data-ttu-id="a032a-510">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-510">
         - Settings</span></span><br><span data-ttu-id="a032a-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-511">
         - TableBindings</span></span><br><span data-ttu-id="a032a-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-512">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-513">
         - TextBindings</span></span><br><span data-ttu-id="a032a-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-514">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-516">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-517">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a032a-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a032a-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="a032a-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-520">- BindingEvents</span></span><br><span data-ttu-id="a032a-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-521">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-523">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-524">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-524">
         - File</span></span><br><span data-ttu-id="a032a-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-526">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-527">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-530">
         - PdfFile</span></span><br><span data-ttu-id="a032a-531">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-531">
         - Selection</span></span><br><span data-ttu-id="a032a-532">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-532">
         - Settings</span></span><br><span data-ttu-id="a032a-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-533">
         - TableBindings</span></span><br><span data-ttu-id="a032a-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-534">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-535">
         - TextBindings</span></span><br><span data-ttu-id="a032a-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-536">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-538">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-539">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a032a-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="a032a-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-541">- BindingEvents</span></span><br><span data-ttu-id="a032a-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-542">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-544">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-545">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-545">
         - File</span></span><br><span data-ttu-id="a032a-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-547">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-548">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-551">
         - PdfFile</span></span><br><span data-ttu-id="a032a-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-552">
         - Selection</span></span><br><span data-ttu-id="a032a-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-553">
         - Settings</span></span><br><span data-ttu-id="a032a-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-554">
         - TableBindings</span></span><br><span data-ttu-id="a032a-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-555">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-556">
         - TextBindings</span></span><br><span data-ttu-id="a032a-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-557">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-559">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="a032a-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="a032a-560">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a032a-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a032a-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a032a-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a032a-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a032a-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-565">- BindingEvents</span></span><br><span data-ttu-id="a032a-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-566">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-568">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-569">
         - File</span></span><br><span data-ttu-id="a032a-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-571">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-572">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-575">
         - PdfFile</span></span><br><span data-ttu-id="a032a-576">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-576">
         - Selection</span></span><br><span data-ttu-id="a032a-577">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-577">
         - Settings</span></span><br><span data-ttu-id="a032a-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-578">
         - TableBindings</span></span><br><span data-ttu-id="a032a-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-579">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-580">
         - TextBindings</span></span><br><span data-ttu-id="a032a-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-581">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-583">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-584">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-584">- TaskPane</span></span><br><span data-ttu-id="a032a-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a032a-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a032a-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a032a-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a032a-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a032a-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-590">- BindingEvents</span></span><br><span data-ttu-id="a032a-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-591">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-593">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-594">
         - File</span></span><br><span data-ttu-id="a032a-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-596">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-597">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-600">
         - PdfFile</span></span><br><span data-ttu-id="a032a-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-601">
         - Selection</span></span><br><span data-ttu-id="a032a-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-602">
         - Settings</span></span><br><span data-ttu-id="a032a-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-603">
         - TableBindings</span></span><br><span data-ttu-id="a032a-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-604">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-605">
         - TextBindings</span></span><br><span data-ttu-id="a032a-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-606">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-608">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-609">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-609">- TaskPane</span></span><br><span data-ttu-id="a032a-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a032a-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a032a-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a032a-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a032a-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a032a-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a032a-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a032a-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-615">- BindingEvents</span></span><br><span data-ttu-id="a032a-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-616">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-618">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-619">
         - File</span></span><br><span data-ttu-id="a032a-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-621">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-622">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-625">
         - PdfFile</span></span><br><span data-ttu-id="a032a-626">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-626">
         - Selection</span></span><br><span data-ttu-id="a032a-627">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-627">
         - Settings</span></span><br><span data-ttu-id="a032a-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-628">
         - TableBindings</span></span><br><span data-ttu-id="a032a-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-629">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-630">
         - TextBindings</span></span><br><span data-ttu-id="a032a-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-631">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-633">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-634">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a032a-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a032a-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="a032a-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-637">- BindingEvents</span></span><br><span data-ttu-id="a032a-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-638">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a032a-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="a032a-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-640">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-641">
         - File</span></span><br><span data-ttu-id="a032a-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-643">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-644">
         - MatrixBindings</span></span><br><span data-ttu-id="a032a-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="a032a-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a032a-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-647">
         - PdfFile</span></span><br><span data-ttu-id="a032a-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-648">
         - Selection</span></span><br><span data-ttu-id="a032a-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-649">
         - Settings</span></span><br><span data-ttu-id="a032a-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-650">
         - TableBindings</span></span><br><span data-ttu-id="a032a-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-651">
         - TableCoercion</span></span><br><span data-ttu-id="a032a-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a032a-652">
         - TextBindings</span></span><br><span data-ttu-id="a032a-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-653">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a032a-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="a032a-655">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="a032a-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="a032a-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a032a-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a032a-657">平台</span><span class="sxs-lookup"><span data-stu-id="a032a-657">Platform</span></span></th>
    <th><span data-ttu-id="a032a-658">扩展点</span><span class="sxs-lookup"><span data-stu-id="a032a-658">Extension points</span></span></th>
    <th><span data-ttu-id="a032a-659">API 要求集</span><span class="sxs-lookup"><span data-stu-id="a032a-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="a032a-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a032a-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="a032a-661">Office Online</span></span></td>
    <td> <span data-ttu-id="a032a-662">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-662">- Content</span></span><br><span data-ttu-id="a032a-663">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-663">
         - TaskPane</span></span><br><span data-ttu-id="a032a-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-666">- ActiveView</span></span><br><span data-ttu-id="a032a-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-667">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-668">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-669">
         - File</span></span><br><span data-ttu-id="a032a-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-670">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-671">
         - PdfFile</span></span><br><span data-ttu-id="a032a-672">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-672">
         - Selection</span></span><br><span data-ttu-id="a032a-673">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-673">
         - Settings</span></span><br><span data-ttu-id="a032a-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-675">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-676">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-676">- Content</span></span><br><span data-ttu-id="a032a-677">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-677">
         - TaskPane</span></span><br><span data-ttu-id="a032a-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-680">- ActiveView</span></span><br><span data-ttu-id="a032a-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-681">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-682">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-683">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-683">
         - File</span></span><br><span data-ttu-id="a032a-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-684">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-685">
         - PdfFile</span></span><br><span data-ttu-id="a032a-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-686">
         - Selection</span></span><br><span data-ttu-id="a032a-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-687">
         - Settings</span></span><br><span data-ttu-id="a032a-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-689">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-690">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-690">- Content</span></span><br><span data-ttu-id="a032a-691">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-691">
         - TaskPane</span></span><br><span data-ttu-id="a032a-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-694">- ActiveView</span></span><br><span data-ttu-id="a032a-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-695">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-696">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-697">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-697">
         - File</span></span><br><span data-ttu-id="a032a-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-698">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-699">
         - PdfFile</span></span><br><span data-ttu-id="a032a-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-700">
         - Selection</span></span><br><span data-ttu-id="a032a-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-701">
         - Settings</span></span><br><span data-ttu-id="a032a-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-703">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-704">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-704">- Content</span></span><br><span data-ttu-id="a032a-705">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a032a-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="a032a-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-707">- ActiveView</span></span><br><span data-ttu-id="a032a-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-708">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-709">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-710">
         - File</span></span><br><span data-ttu-id="a032a-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-711">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-712">
         - PdfFile</span></span><br><span data-ttu-id="a032a-713">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-713">
         - Selection</span></span><br><span data-ttu-id="a032a-714">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-714">
         - Settings</span></span><br><span data-ttu-id="a032a-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-716">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-717">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-717">- Content</span></span><br><span data-ttu-id="a032a-718">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="a032a-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a032a-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="a032a-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-720">- ActiveView</span></span><br><span data-ttu-id="a032a-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-721">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-722">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-723">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-723">
         - File</span></span><br><span data-ttu-id="a032a-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-724">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-725">
         - PdfFile</span></span><br><span data-ttu-id="a032a-726">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-726">
         - Selection</span></span><br><span data-ttu-id="a032a-727">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-727">
         - Settings</span></span><br><span data-ttu-id="a032a-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-729">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="a032a-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="a032a-730">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-730">- Content</span></span><br><span data-ttu-id="a032a-731">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="a032a-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-733">- ActiveView</span></span><br><span data-ttu-id="a032a-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-734">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-735">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-736">
         - File</span></span><br><span data-ttu-id="a032a-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-737">
         - PdfFile</span></span><br><span data-ttu-id="a032a-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-738">
         - Selection</span></span><br><span data-ttu-id="a032a-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-739">
         - Settings</span></span><br><span data-ttu-id="a032a-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-740">
         - TextCoercion</span></span><br><span data-ttu-id="a032a-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-742">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-743">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-743">- Content</span></span><br><span data-ttu-id="a032a-744">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-744">
         - TaskPane</span></span><br><span data-ttu-id="a032a-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-747">- ActiveView</span></span><br><span data-ttu-id="a032a-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-748">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-749">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-750">
         - File</span></span><br><span data-ttu-id="a032a-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-751">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-752">
         - PdfFile</span></span><br><span data-ttu-id="a032a-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-753">
         - Selection</span></span><br><span data-ttu-id="a032a-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-754">
         - Settings</span></span><br><span data-ttu-id="a032a-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-756">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-757">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-757">- Content</span></span><br><span data-ttu-id="a032a-758">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-758">
         - TaskPane</span></span><br><span data-ttu-id="a032a-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-761">- ActiveView</span></span><br><span data-ttu-id="a032a-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-762">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-763">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-764">
         - File</span></span><br><span data-ttu-id="a032a-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-765">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-766">
         - PdfFile</span></span><br><span data-ttu-id="a032a-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-767">
         - Selection</span></span><br><span data-ttu-id="a032a-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-768">
         - Settings</span></span><br><span data-ttu-id="a032a-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-770">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a032a-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="a032a-771">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-771">- Content</span></span><br><span data-ttu-id="a032a-772">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a032a-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="a032a-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a032a-774">- ActiveView</span></span><br><span data-ttu-id="a032a-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a032a-775">
         - CompressedFile</span></span><br><span data-ttu-id="a032a-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-776">
         - DocumentEvents</span></span><br><span data-ttu-id="a032a-777">
         - File</span><span class="sxs-lookup"><span data-stu-id="a032a-777">
         - File</span></span><br><span data-ttu-id="a032a-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-778">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a032a-779">
         - PdfFile</span></span><br><span data-ttu-id="a032a-780">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-780">
         - Selection</span></span><br><span data-ttu-id="a032a-781">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-781">
         - Settings</span></span><br><span data-ttu-id="a032a-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="a032a-783">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="a032a-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="a032a-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="a032a-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a032a-785">平台</span><span class="sxs-lookup"><span data-stu-id="a032a-785">Platform</span></span></th>
    <th><span data-ttu-id="a032a-786">扩展点</span><span class="sxs-lookup"><span data-stu-id="a032a-786">Extension points</span></span></th>
    <th><span data-ttu-id="a032a-787">API 要求集</span><span class="sxs-lookup"><span data-stu-id="a032a-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="a032a-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a032a-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="a032a-789">Office Online</span></span></td>
    <td> <span data-ttu-id="a032a-790">- 内容</span><span class="sxs-lookup"><span data-stu-id="a032a-790">- Content</span></span><br><span data-ttu-id="a032a-791">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-791">
         - TaskPane</span></span><br><span data-ttu-id="a032a-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="a032a-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a032a-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="a032a-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a032a-795">- DocumentEvents</span></span><br><span data-ttu-id="a032a-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="a032a-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-797">
         - ImageCoercion</span></span><br><span data-ttu-id="a032a-798">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a032a-798">
         - Settings</span></span><br><span data-ttu-id="a032a-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="a032a-800">项目</span><span class="sxs-lookup"><span data-stu-id="a032a-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a032a-801">平台</span><span class="sxs-lookup"><span data-stu-id="a032a-801">Platform</span></span></th>
    <th><span data-ttu-id="a032a-802">扩展点</span><span class="sxs-lookup"><span data-stu-id="a032a-802">Extension points</span></span></th>
    <th><span data-ttu-id="a032a-803">API 要求集</span><span class="sxs-lookup"><span data-stu-id="a032a-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="a032a-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a032a-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-805">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-806">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-808">- Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-808">- Selection</span></span><br><span data-ttu-id="a032a-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-810">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-811">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-813">- Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-813">- Selection</span></span><br><span data-ttu-id="a032a-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a032a-815">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="a032a-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="a032a-816">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="a032a-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a032a-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a032a-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a032a-818">- Selection</span><span class="sxs-lookup"><span data-stu-id="a032a-818">- Selection</span></span><br><span data-ttu-id="a032a-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a032a-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="a032a-820">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a032a-820">See also</span></span>

- [<span data-ttu-id="a032a-821">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="a032a-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="a032a-822">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="a032a-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="a032a-823">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="a032a-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="a032a-824">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="a032a-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
