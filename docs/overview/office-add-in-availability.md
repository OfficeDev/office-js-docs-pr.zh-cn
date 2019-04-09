---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477591"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="fb4de-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="fb4de-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="fb4de-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="fb4de-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="fb4de-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="fb4de-106">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="fb4de-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="fb4de-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="fb4de-108">Excel</span><span class="sxs-lookup"><span data-stu-id="fb4de-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="fb4de-109">平台</span><span class="sxs-lookup"><span data-stu-id="fb4de-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="fb4de-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="fb4de-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="fb4de-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fb4de-111">API requirement sets</span></span></th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="fb4de-112">通用 API</span><span class="sxs-lookup"><span data-stu-id="fb4de-112">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="fb4de-113">Office Online</span></span></td>
    <td> - <span data-ttu-id="fb4de-114">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-114">TaskPane</span></span><br>
        - <span data-ttu-id="fb4de-115">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-115">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-116">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-116">Add-in Commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-117">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-117">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-118">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-118">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-119">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-119">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-120">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-120">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-121">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-122">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-122">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-123">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-123">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-124">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="fb4de-124">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-125">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-125">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="fb4de-126">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-126">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-127">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-127">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-128">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-128">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-129">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-129">File</span></span><br>
        - <span data-ttu-id="fb4de-130">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-130">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-131">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-131">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-132">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-132">Selection</span></span><br>
        - <span data-ttu-id="fb4de-133">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-133">Settings</span></span><br>
        - <span data-ttu-id="fb4de-134">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-134">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-135">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-135">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-136">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-136">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-137">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-137">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-138">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-138">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-139">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-139">TaskPane</span></span><br>
        - <span data-ttu-id="fb4de-140">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-140">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-141">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-141">Add-in Commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-142">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-142">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-143">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-143">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-144">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-144">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-145">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-145">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-146">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-146">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-147">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-147">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-148">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-148">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-149">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="fb4de-149">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-150">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-150">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="fb4de-151">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-151">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-152">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-152">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-153">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-153">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-154">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-154">File</span></span><br>
        - <span data-ttu-id="fb4de-155">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-155">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-156">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-156">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-157">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-157">Selection</span></span><br>
        - <span data-ttu-id="fb4de-158">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-158">Settings</span></span><br>
        - <span data-ttu-id="fb4de-159">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-159">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-160">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-160">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-161">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-161">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-162">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-162">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-163">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-163">Office 2019 for Windows</span></span></td>
    <td>- <span data-ttu-id="fb4de-164">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-164"> Taskpane</span></span><br>
        - <span data-ttu-id="fb4de-165">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-165">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-166">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-166">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-167">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-167">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-168">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-168">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-169">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-169">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-170">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-170">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-171">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-171">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-172">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-172">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-173">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-173">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-174">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="fb4de-174">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-175">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-175">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="fb4de-176">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-176">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-177">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-177">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-178">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-178">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-179">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-179">File</span></span><br>
        - <span data-ttu-id="fb4de-180">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-180">ImageCoercion</span></span><br>
        - <span data-ttu-id="fb4de-181">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-181">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-182">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-182">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-183">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-183">Selection</span></span><br>
        - <span data-ttu-id="fb4de-184">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-184">Settings</span></span><br>
        - <span data-ttu-id="fb4de-185">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-185">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-186">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-186">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-187">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-187">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-188">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-188">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-189">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-189">Office 2016 for Windows</span></span></td>
    <td>- <span data-ttu-id="fb4de-190">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-190">TaskPane</span></span><br>
        - <span data-ttu-id="fb4de-191">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-191">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-192">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-192">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-193">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-193">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="fb4de-194">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-194">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-195">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-195">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-196">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-196">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-197">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-197">File</span></span><br>
        - <span data-ttu-id="fb4de-198">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-198">ImageCoercion</span></span><br>
        - <span data-ttu-id="fb4de-199">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-199">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-200">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-200">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-201">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-201">Selection</span></span><br>
        - <span data-ttu-id="fb4de-202">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-202">Settings</span></span><br>
        - <span data-ttu-id="fb4de-203">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-203">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-204">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-204">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-205">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-205">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-206">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-206">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-207">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-207">Office 2013 for Windows</span></span></td>
    <td>
        - <span data-ttu-id="fb4de-208">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-208">TaskPane</span></span><br>
        - <span data-ttu-id="fb4de-209">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-209">Content</span></span></td>
    <td>  - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-210">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-210">DialogApi 1.1</span></span></a>*</td>
    <td>
        - <span data-ttu-id="fb4de-211">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-211">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-212">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-212">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-213">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-213">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-214">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-214">File</span></span><br>
        - <span data-ttu-id="fb4de-215">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-215">ImageCoercion</span></span><br>
        - <span data-ttu-id="fb4de-216">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-216">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-217">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-217">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-218">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-218">Selection</span></span><br>
        - <span data-ttu-id="fb4de-219">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-219">Settings</span></span><br>
        - <span data-ttu-id="fb4de-220">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-220">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-221">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-221">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-222">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-222">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-223">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-223">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-224">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="fb4de-224">Office 365 for iPad</span></span></td>
    <td>- <span data-ttu-id="fb4de-225">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-225">TaskPane</span></span><br>
        - <span data-ttu-id="fb4de-226">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-226">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-227">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-227">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-228">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-228">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-229">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-229">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-230">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-230">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-231">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-231">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-232">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-232">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-233">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-233">ExcelApi 1.7</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-234">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="fb4de-234">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-235">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-235">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="fb4de-236">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-236">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-237">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-237">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-238">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-238">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-239">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-239">File</span></span><br>
        - <span data-ttu-id="fb4de-240">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-240">ImageCoercion</span></span><br>
        - <span data-ttu-id="fb4de-241">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-241">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-242">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-242">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-243">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-243">Selection</span></span><br>
        - <span data-ttu-id="fb4de-244">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-244">Settings</span></span><br>
        - <span data-ttu-id="fb4de-245">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-245">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-246">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-246">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-247">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-247">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-248">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-248">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-249">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-249">Office 365 for Mac</span></span></td>
    <td>- <span data-ttu-id="fb4de-250">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-250">TaskPane</span></span><br>
        - <span data-ttu-id="fb4de-251">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-251">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-252">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-252">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-253">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-253">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-254">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-254">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-255">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-255">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-256">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-256">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-257">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-257">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-258">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-258">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-259">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-259">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-260">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="fb4de-260">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-261">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-261">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="fb4de-262">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-262">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-263">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-263">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-264">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-264">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-265">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-265">File</span></span><br>
        - <span data-ttu-id="fb4de-266">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-266">ImageCoercion</span></span><br>
        - <span data-ttu-id="fb4de-267">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-267">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-268">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-268">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-269">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-269">PdfFile</span></span><br>
        - <span data-ttu-id="fb4de-270">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-270">Selection</span></span><br>
        - <span data-ttu-id="fb4de-271">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-271">Settings</span></span><br>
        - <span data-ttu-id="fb4de-272">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-272">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-273">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-273">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-274">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-274">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-275">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-275">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-276">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-276">Office 2019 for Mac</span></span></td>
    <td>- <span data-ttu-id="fb4de-277">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-277">TaskPane</span></span><br>
        - <span data-ttu-id="fb4de-278">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-278">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-279">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-279">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-280">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-280">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-281">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-281">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-282">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-282">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-283">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-283">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-284">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-284">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-285">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-285">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-286">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-286">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-287">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="fb4de-287">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-288">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-288">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="fb4de-289">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-289">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-290">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-290">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-291">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-291">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-292">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-292">File</span></span><br>
        - <span data-ttu-id="fb4de-293">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-293">ImageCoercion</span></span><br>
        - <span data-ttu-id="fb4de-294">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-294">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-295">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-295">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-296">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-296">PdfFile</span></span><br>
        - <span data-ttu-id="fb4de-297">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-297">Selection</span></span><br>
        - <span data-ttu-id="fb4de-298">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-298">Settings</span></span><br>
        - <span data-ttu-id="fb4de-299">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-299">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-300">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-300">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-301">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-301">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-302">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-302">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-303">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-303">Office 2016 for Mac</span></span></td>
    <td>- <span data-ttu-id="fb4de-304">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-304"> Taskpane</span></span><br>
        - <span data-ttu-id="fb4de-305">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-305">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="fb4de-306">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-306">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-307">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-307">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="fb4de-308">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-308">BindingEvents</span></span><br>
        - <span data-ttu-id="fb4de-309">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-309">CompressedFile</span></span><br>
        - <span data-ttu-id="fb4de-310">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-310">DocumentEvents</span></span><br>
        - <span data-ttu-id="fb4de-311">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-311">File</span></span><br>
        - <span data-ttu-id="fb4de-312">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-312">ImageCoercion</span></span><br>
        - <span data-ttu-id="fb4de-313">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-313">MatrixBindings</span></span><br>
        - <span data-ttu-id="fb4de-314">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-314">MatrixCoercion</span></span><br>
        - <span data-ttu-id="fb4de-315">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-315">PdfFile</span></span><br>
        - <span data-ttu-id="fb4de-316">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-316">Selection</span></span><br>
        - <span data-ttu-id="fb4de-317">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-317">Settings</span></span><br>
        - <span data-ttu-id="fb4de-318">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-318">TableBindings</span></span><br>
        - <span data-ttu-id="fb4de-319">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-319">TableCoercion</span></span><br>
        - <span data-ttu-id="fb4de-320">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-320">TextBindings</span></span><br>
        - <span data-ttu-id="fb4de-321">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-321">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="fb4de-322">&ast; - 已添加发布后更新。</span><span class="sxs-lookup"><span data-stu-id="fb4de-322">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="outlook"></a><span data-ttu-id="fb4de-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="fb4de-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fb4de-324">平台</span><span class="sxs-lookup"><span data-stu-id="fb4de-324">Platform</span></span></th>
    <th><span data-ttu-id="fb4de-325">扩展点</span><span class="sxs-lookup"><span data-stu-id="fb4de-325">Extension points</span></span></th>
    <th><span data-ttu-id="fb4de-326">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fb4de-326">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="fb4de-327">通用 API</span><span class="sxs-lookup"><span data-stu-id="fb4de-327">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="fb4de-328">Office Online</span></span></td>
    <td> - <span data-ttu-id="fb4de-329">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-329">Mail Read</span></span><br>
      - <span data-ttu-id="fb4de-330">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="fb4de-330">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-331">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-331">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-332">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-332">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-333">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-333">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-334">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-334">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-335">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-335">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="fb4de-336">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-336">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="fb4de-337">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-337">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="fb4de-338">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-338">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="fb4de-339">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-340">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-340">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-341">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-341">Mail Read</span></span><br>
      - <span data-ttu-id="fb4de-342">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="fb4de-342">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-343">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-343">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="fb4de-344">模块</span><span class="sxs-lookup"><span data-stu-id="fb4de-344">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-345">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-345">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-346">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-346">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-347">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-347">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-348">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-348">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="fb4de-349">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-349">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="fb4de-350">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-350">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="fb4de-351">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-351">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="fb4de-352">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-353">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-353">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-354">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-354">Mail Read</span></span><br>
      - <span data-ttu-id="fb4de-355">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="fb4de-355">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-356">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-356">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="fb4de-357">模块</span><span class="sxs-lookup"><span data-stu-id="fb4de-357">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-358">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-358">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-359">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-359">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-360">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-360">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-361">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-361">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="fb4de-362">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-362">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="fb4de-363">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-363">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="fb4de-364">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="fb4de-364">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="fb4de-365">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-366">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-366">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-367">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-367">Mail Read</span></span><br>
      - <span data-ttu-id="fb4de-368">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="fb4de-368">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-369">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-369">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="fb4de-370">模块</span><span class="sxs-lookup"><span data-stu-id="fb4de-370">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-371">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-371">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-372">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-372">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-373">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-373">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-374">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-374">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="fb4de-375">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-376">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-376">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-377">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-377">Mail Read</span></span><br>
      - <span data-ttu-id="fb4de-378">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="fb4de-378">Mail Compose</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-379">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-379">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-380">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-380">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-381">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-381">Mailbox 1.3</span></span></a>*<br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-382">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-382">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="fb4de-383">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-384">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="fb4de-384">Office 365 for iOS</span></span></td>
    <td> - <span data-ttu-id="fb4de-385">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-385">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-386">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-386">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-387">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-387">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-388">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-388">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-389">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-389">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-390">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-390">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="fb4de-391">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-391">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="fb4de-392">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-393">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-393">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-394">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-394">Mail Read</span></span><br>
      - <span data-ttu-id="fb4de-395">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="fb4de-395">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-396">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-396">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-397">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-397">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-398">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-398">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-399">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-399">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-400">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-400">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="fb4de-401">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-401">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="fb4de-402">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-402">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="fb4de-403">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-404">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-404">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-405">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-405">Mail Read</span></span><br>
      - <span data-ttu-id="fb4de-406">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="fb4de-406">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-407">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-407">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-408">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-408">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-409">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-409">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-410">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-410">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-411">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-411">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="fb4de-412">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-412">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="fb4de-413">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-413">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="fb4de-414">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-415">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-415">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-416">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-416">Mail Read</span></span><br>
      - <span data-ttu-id="fb4de-417">邮件撰写</span><span class="sxs-lookup"><span data-stu-id="fb4de-417">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-418">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-418">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-419">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-419">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-420">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-420">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-421">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-421">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-422">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-422">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="fb4de-423">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-423">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="fb4de-424">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="fb4de-424">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="fb4de-425">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-426">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="fb4de-426">Office 365 for Android</span></span></td>
    <td> - <span data-ttu-id="fb4de-427">邮件阅读</span><span class="sxs-lookup"><span data-stu-id="fb4de-427">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-428">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-428">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="fb4de-429">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-429">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="fb4de-430">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-430">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="fb4de-431">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-431">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="fb4de-432">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="fb4de-432">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="fb4de-433">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="fb4de-433">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="fb4de-434">不可用</span><span class="sxs-lookup"><span data-stu-id="fb4de-434">Not available</span></span></td>
  </tr>
</table>

*<span data-ttu-id="fb4de-435">&ast; - 已添加发布后更新。</span><span class="sxs-lookup"><span data-stu-id="fb4de-435">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="word"></a><span data-ttu-id="fb4de-436">Word</span><span class="sxs-lookup"><span data-stu-id="fb4de-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fb4de-437">平台</span><span class="sxs-lookup"><span data-stu-id="fb4de-437">Platform</span></span></th>
    <th><span data-ttu-id="fb4de-438">扩展点</span><span class="sxs-lookup"><span data-stu-id="fb4de-438">Extension points</span></span></th>
    <th><span data-ttu-id="fb4de-439">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fb4de-439">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="fb4de-440">通用 API</span><span class="sxs-lookup"><span data-stu-id="fb4de-440">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="fb4de-441">Office Online</span></span></td>
    <td> - <span data-ttu-id="fb4de-442">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-442">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-443">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-443">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-444">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-444">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-445">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-445">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-446">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-446">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-447">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-447">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-448">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-448">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-449">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-449">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-450">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-450">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-451">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-451">File</span></span><br>
         - <span data-ttu-id="fb4de-452">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-452">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-453">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-453">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-454">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-454">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-455">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-455">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-456">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-456">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-457">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-457">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-458">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-458">Selection</span></span><br>
         - <span data-ttu-id="fb4de-459">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-459">Settings</span></span><br>
         - <span data-ttu-id="fb4de-460">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-460">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-461">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-461">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-462">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-462">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-463">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-463">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-464">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-464">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-465">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-465">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-466">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-466">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-467">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-467">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-468">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-468">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-469">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-469">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-470">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-470">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-471">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-471">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-472">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-472">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-473">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-473">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-474">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-474">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-475">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-475">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-476">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-476">File</span></span><br>
         - <span data-ttu-id="fb4de-477">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-477">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-478">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-478">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-479">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-479">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-480">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-480">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-481">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-481">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-482">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-482">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-483">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-483">Selection</span></span><br>
         - <span data-ttu-id="fb4de-484">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-484">Settings</span></span><br>
         - <span data-ttu-id="fb4de-485">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-485">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-486">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-486">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-487">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-487">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-488">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-488">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-489">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-489">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-490">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-490">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-491">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-491"> Taskpane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-492">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-492">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-493">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-493">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-494">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-494">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-495">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-495">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-496">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-496">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-497">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-497">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-498">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-498">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-499">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-499">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-500">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-500">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-501">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-501">File</span></span><br>
         - <span data-ttu-id="fb4de-502">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-502">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-503">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-503">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-504">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-504">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-505">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-505">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-506">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-506">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-507">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-507">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-508">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-508">Selection</span></span><br>
         - <span data-ttu-id="fb4de-509">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-509">Settings</span></span><br>
         - <span data-ttu-id="fb4de-510">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-510">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-511">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-511">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-512">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-512">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-513">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-513">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-514">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-514">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-515">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-515">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-516">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-516"> Taskpane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-517">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-517">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-518">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-518">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="fb4de-519">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-519">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-520">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-520">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-521">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-521">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-522">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-522">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-523">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-523">File</span></span><br>
         - <span data-ttu-id="fb4de-524">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-524">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-525">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-525">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-526">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-526">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-527">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-527">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-528">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-528">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-529">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-529">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-530">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-530">Selection</span></span><br>
         - <span data-ttu-id="fb4de-531">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-531">Settings</span></span><br>
         - <span data-ttu-id="fb4de-532">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-532">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-533">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-533">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-534">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-534">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-535">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-535">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-536">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-536">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-537">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-537">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-538">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-538">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-539">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-539">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="fb4de-540">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-540">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-541">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-541">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-542">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-542">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-543">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-543">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-544">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-544">File</span></span><br>
         - <span data-ttu-id="fb4de-545">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-545">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-546">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-546">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-547">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-547">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-548">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-548">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-549">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-549">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-550">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-550">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-551">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-551">Selection</span></span><br>
         - <span data-ttu-id="fb4de-552">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-552">Settings</span></span><br>
         - <span data-ttu-id="fb4de-553">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-553">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-554">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-554">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-555">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-555">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-556">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-556">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-557">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-557">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-558">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="fb4de-558">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="fb4de-559">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-559">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-560">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-560">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-561">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-561">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-562">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-562">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-563">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-563">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="fb4de-564">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-564">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-565">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-565">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-566">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-566">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-567">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-567">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-568">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-568">File</span></span><br>
         - <span data-ttu-id="fb4de-569">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-569">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-570">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-570">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-571">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-571">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-572">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-572">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-573">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-573">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-574">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-574">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-575">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-575">Selection</span></span><br>
         - <span data-ttu-id="fb4de-576">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-576">Settings</span></span><br>
         - <span data-ttu-id="fb4de-577">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-577">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-578">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-578">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-579">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-579">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-580">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-580">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-581">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-581">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-582">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-582">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-583">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-583">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-584">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-584">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-585">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-585">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-586">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-586">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-587">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-587">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-588">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-588">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="fb4de-589">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-589">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-590">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-590">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-591">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-591">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-592">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-592">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-593">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-593">File</span></span><br>
         - <span data-ttu-id="fb4de-594">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-594">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-595">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-595">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-596">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-596">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-597">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-597">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-598">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-598">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-599">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-599">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-600">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-600">Selection</span></span><br>
         - <span data-ttu-id="fb4de-601">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-601">Settings</span></span><br>
         - <span data-ttu-id="fb4de-602">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-602">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-603">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-603">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-604">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-604">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-605">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-605">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-606">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-606">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-607">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-607">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-608">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-608">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-609">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-609">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-610">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-610">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-611">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="fb4de-611">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-612">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="fb4de-612">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-613">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-613">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="fb4de-614">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-614">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-615">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-615">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-616">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-616">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-617">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-617">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-618">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-618">File</span></span><br>
         - <span data-ttu-id="fb4de-619">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-619">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-620">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-620">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-621">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-621">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-622">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-622">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-623">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-623">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-624">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-624">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-625">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-625">Selection</span></span><br>
         - <span data-ttu-id="fb4de-626">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-626">Settings</span></span><br>
         - <span data-ttu-id="fb4de-627">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-627">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-628">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-628">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-629">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-629">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-630">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-630">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-631">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-631">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-632">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-632">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-633">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-633"> Taskpane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="fb4de-634">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-634">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-635">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-635">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="fb4de-636">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-636">BindingEvents</span></span><br>
         - <span data-ttu-id="fb4de-637">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-637">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-638">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fb4de-638">CustomXmlParts</span></span><br>
         - <span data-ttu-id="fb4de-639">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-639">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-640">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-640">File</span></span><br>
         - <span data-ttu-id="fb4de-641">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-641">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-642">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-642">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-643">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-643">MatrixBindings</span></span><br>
         - <span data-ttu-id="fb4de-644">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-644">MatrixCoercion</span></span><br>
         - <span data-ttu-id="fb4de-645">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-645">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-646">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-646">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-647">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-647">Selection</span></span><br>
         - <span data-ttu-id="fb4de-648">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-648">Settings</span></span><br>
         - <span data-ttu-id="fb4de-649">TableBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-649">TableBindings</span></span><br>
         - <span data-ttu-id="fb4de-650">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-650">TableCoercion</span></span><br>
         - <span data-ttu-id="fb4de-651">TextBindings</span><span class="sxs-lookup"><span data-stu-id="fb4de-651">TextBindings</span></span><br>
         - <span data-ttu-id="fb4de-652">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-652">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-653">TextFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-653">TextFile</span></span> </td>
  </tr>
</table>

*<span data-ttu-id="fb4de-654">&ast; - 已添加发布后更新。</span><span class="sxs-lookup"><span data-stu-id="fb4de-654">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="powerpoint"></a><span data-ttu-id="fb4de-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="fb4de-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fb4de-656">平台</span><span class="sxs-lookup"><span data-stu-id="fb4de-656">Platform</span></span></th>
    <th><span data-ttu-id="fb4de-657">扩展点</span><span class="sxs-lookup"><span data-stu-id="fb4de-657">Extension points</span></span></th>
    <th><span data-ttu-id="fb4de-658">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fb4de-658">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="fb4de-659">通用 API</span><span class="sxs-lookup"><span data-stu-id="fb4de-659">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="fb4de-660">Office Online</span></span></td>
    <td> - <span data-ttu-id="fb4de-661">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-661">Content</span></span><br>
         - <span data-ttu-id="fb4de-662">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-662">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-663">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-663">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-664">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-664">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-665">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-665">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-666">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-666">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-667">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-667">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-668">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-668">File</span></span><br>
         - <span data-ttu-id="fb4de-669">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-669">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-670">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-670">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-671">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-671">Selection</span></span><br>
         - <span data-ttu-id="fb4de-672">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-672">Settings</span></span><br>
         - <span data-ttu-id="fb4de-673">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-673">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-674">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-674">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-675">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-675">Content</span></span><br>
         - <span data-ttu-id="fb4de-676">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-676">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-677">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-677">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-678">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-678">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-679">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-679">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-680">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-680">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-681">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-681">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-682">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-682">File</span></span><br>
         - <span data-ttu-id="fb4de-683">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-683">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-684">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-684">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-685">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-685">Selection</span></span><br>
         - <span data-ttu-id="fb4de-686">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-686">Settings</span></span><br>
         - <span data-ttu-id="fb4de-687">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-687">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-688">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-688">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-689">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-689">Content</span></span><br>
         - <span data-ttu-id="fb4de-690">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-690">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-691">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-691">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-692">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-692">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-693">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-693">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-694">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-694">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-695">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-695">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-696">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-696">File</span></span><br>
         - <span data-ttu-id="fb4de-697">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-697">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-698">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-698">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-699">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-699">Selection</span></span><br>
         - <span data-ttu-id="fb4de-700">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-700">Settings</span></span><br>
         - <span data-ttu-id="fb4de-701">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-701">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-702">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-702">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-703">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-703">Content</span></span><br>
         - <span data-ttu-id="fb4de-704">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-704">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-705">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-705">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="fb4de-706">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-706">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-707">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-707">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-708">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-708">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-709">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-709">File</span></span><br>
         - <span data-ttu-id="fb4de-710">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-710">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-711">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-711">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-712">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-712">Selection</span></span><br>
         - <span data-ttu-id="fb4de-713">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-713">Settings</span></span><br>
         - <span data-ttu-id="fb4de-714">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-714">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-715">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-715">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-716">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-716">Content</span></span><br>
         - <span data-ttu-id="fb4de-717">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-717">TaskPane</span></span><br>
    </td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-718">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-718">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="fb4de-719">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-719">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-720">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-720">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-721">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-721">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-722">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-722">File</span></span><br>
         - <span data-ttu-id="fb4de-723">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-723">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-724">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-724">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-725">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-725">Selection</span></span><br>
         - <span data-ttu-id="fb4de-726">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-726">Settings</span></span><br>
         - <span data-ttu-id="fb4de-727">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-727">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-728">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="fb4de-728">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="fb4de-729">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-729">Content</span></span><br>
         - <span data-ttu-id="fb4de-730">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-730">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-731">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-731">DialogApi 1.1</span></span></a></td>
     <td> - <span data-ttu-id="fb4de-732">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-732">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-733">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-733">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-734">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-734">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-735">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-735">File</span></span><br>
         - <span data-ttu-id="fb4de-736">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-736">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-737">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-737">Selection</span></span><br>
         - <span data-ttu-id="fb4de-738">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-738">Settings</span></span><br>
         - <span data-ttu-id="fb4de-739">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-739">TextCoercion</span></span><br>
         - <span data-ttu-id="fb4de-740">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-740">ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-741">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-741">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-742">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-742">Content</span></span><br>
         - <span data-ttu-id="fb4de-743">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-743">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-744">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-744">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-745">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-745">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-746">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-746">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-747">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-747">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-748">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-748">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-749">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-749">File</span></span><br>
         - <span data-ttu-id="fb4de-750">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-750">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-751">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-751">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-752">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-752">Selection</span></span><br>
         - <span data-ttu-id="fb4de-753">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-753">Settings</span></span><br>
         - <span data-ttu-id="fb4de-754">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-754">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-755">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-755">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-756">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-756">Content</span></span><br>
         - <span data-ttu-id="fb4de-757">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-757">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-758">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-758">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-759">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-759">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-760">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-760">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-761">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-761">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-762">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-762">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-763">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-763">File</span></span><br>
         - <span data-ttu-id="fb4de-764">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-764">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-765">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-765">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-766">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-766">Selection</span></span><br>
         - <span data-ttu-id="fb4de-767">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-767">Settings</span></span><br>
         - <span data-ttu-id="fb4de-768">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-768">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-769">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="fb4de-769">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="fb4de-770">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-770">Content</span></span><br>
         - <span data-ttu-id="fb4de-771">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-771">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-772">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-772">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="fb4de-773">ActiveView</span><span class="sxs-lookup"><span data-stu-id="fb4de-773">ActiveView</span></span><br>
         - <span data-ttu-id="fb4de-774">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-774">CompressedFile</span></span><br>
         - <span data-ttu-id="fb4de-775">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-775">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-776">File</span><span class="sxs-lookup"><span data-stu-id="fb4de-776">File</span></span><br>
         - <span data-ttu-id="fb4de-777">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-777">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-778">PdfFile</span><span class="sxs-lookup"><span data-stu-id="fb4de-778">PdfFile</span></span><br>
         - <span data-ttu-id="fb4de-779">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-779">Selection</span></span><br>
         - <span data-ttu-id="fb4de-780">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-780">Settings</span></span><br>
         - <span data-ttu-id="fb4de-781">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-781">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="fb4de-782">&ast; - 已添加发布后更新。</span><span class="sxs-lookup"><span data-stu-id="fb4de-782">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="onenote"></a><span data-ttu-id="fb4de-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="fb4de-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fb4de-784">平台</span><span class="sxs-lookup"><span data-stu-id="fb4de-784">Platform</span></span></th>
    <th><span data-ttu-id="fb4de-785">扩展点</span><span class="sxs-lookup"><span data-stu-id="fb4de-785">Extension points</span></span></th>
    <th><span data-ttu-id="fb4de-786">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fb4de-786">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="fb4de-787">通用 API</span><span class="sxs-lookup"><span data-stu-id="fb4de-787">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="fb4de-788">Office Online</span></span></td>
    <td> - <span data-ttu-id="fb4de-789">内容</span><span class="sxs-lookup"><span data-stu-id="fb4de-789">Content</span></span><br>
         - <span data-ttu-id="fb4de-790">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-790">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="fb4de-791">加载项命令</span><span class="sxs-lookup"><span data-stu-id="fb4de-791">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets"><span data-ttu-id="fb4de-792">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-792">OneNoteApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-793">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-793">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-794">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fb4de-794">DocumentEvents</span></span><br>
         - <span data-ttu-id="fb4de-795">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-795">HtmlCoercion</span></span><br>
         - <span data-ttu-id="fb4de-796">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-796">ImageCoercion</span></span><br>
         - <span data-ttu-id="fb4de-797">Settings</span><span class="sxs-lookup"><span data-stu-id="fb4de-797">Settings</span></span><br>
         - <span data-ttu-id="fb4de-798">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-798">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="fb4de-799">Project</span><span class="sxs-lookup"><span data-stu-id="fb4de-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fb4de-800">平台</span><span class="sxs-lookup"><span data-stu-id="fb4de-800">Platform</span></span></th>
    <th><span data-ttu-id="fb4de-801">扩展点</span><span class="sxs-lookup"><span data-stu-id="fb4de-801">Extension points</span></span></th>
    <th><span data-ttu-id="fb4de-802">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fb4de-802">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="fb4de-803">通用 API</span><span class="sxs-lookup"><span data-stu-id="fb4de-803">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-804">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-804">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-805">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-805">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-806">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-806">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-807">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-807">Selection</span></span><br>
         - <span data-ttu-id="fb4de-808">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-808">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-809">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-809">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-810">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-810">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-811">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-811">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-812">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-812">Selection</span></span><br>
         - <span data-ttu-id="fb4de-813">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-813">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fb4de-814">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="fb4de-814">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="fb4de-815">任务窗格</span><span class="sxs-lookup"><span data-stu-id="fb4de-815">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="fb4de-816">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb4de-816">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="fb4de-817">Selection</span><span class="sxs-lookup"><span data-stu-id="fb4de-817">Selection</span></span><br>
         - <span data-ttu-id="fb4de-818">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fb4de-818">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="fb4de-819">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fb4de-819">See also</span></span>

- [<span data-ttu-id="fb4de-820">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="fb4de-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="fb4de-821">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="fb4de-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="fb4de-822">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="fb4de-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="fb4de-823">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="fb4de-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="fb4de-824">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="fb4de-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="fb4de-825">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="fb4de-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="fb4de-826">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="fb4de-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="fb4de-827">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="fb4de-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)