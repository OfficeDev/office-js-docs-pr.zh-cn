---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 06/13/2019
localization_priority: Priority
ms.openlocfilehash: 82c276c802cab66ae4f5443d0d556bc42ee57841
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128620"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="86cac-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="86cac-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="86cac-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="86cac-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="86cac-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="86cac-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="86cac-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="86cac-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="86cac-108">Excel</span><span class="sxs-lookup"><span data-stu-id="86cac-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="86cac-109">平台</span><span class="sxs-lookup"><span data-stu-id="86cac-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="86cac-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="86cac-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="86cac-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="86cac-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="86cac-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="86cac-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="86cac-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-114">- TaskPane</span></span><br><span data-ttu-id="86cac-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-115">
        - Content</span></span><br><span data-ttu-id="86cac-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="86cac-116">
        - Custom Functions</span></span><br><span data-ttu-id="86cac-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="86cac-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="86cac-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="86cac-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="86cac-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="86cac-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="86cac-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="86cac-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="86cac-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="86cac-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="86cac-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="86cac-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="86cac-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="86cac-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="86cac-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-128">
        - BindingEvents</span></span><br><span data-ttu-id="86cac-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-129">
        - CompressedFile</span></span><br><span data-ttu-id="86cac-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-130">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-131">
        - File</span></span><br><span data-ttu-id="86cac-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-132">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-134">
        - Selection</span></span><br><span data-ttu-id="86cac-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-135">
        - Settings</span></span><br><span data-ttu-id="86cac-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-136">
        - TableBindings</span></span><br><span data-ttu-id="86cac-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-137">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-138">
        - TextBindings</span></span><br><span data-ttu-id="86cac-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-140">Office on Windows</span></span><br><span data-ttu-id="86cac-141">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-142">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-142">- TaskPane</span></span><br><span data-ttu-id="86cac-143">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-143">
        - Content</span></span><br><span data-ttu-id="86cac-144">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="86cac-144">
        - Custom Functions</span></span><br><span data-ttu-id="86cac-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="86cac-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="86cac-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="86cac-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="86cac-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="86cac-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="86cac-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="86cac-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="86cac-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="86cac-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="86cac-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="86cac-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="86cac-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="86cac-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="86cac-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-156">
        - BindingEvents</span></span><br><span data-ttu-id="86cac-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-157">
        - CompressedFile</span></span><br><span data-ttu-id="86cac-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-158">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-159">
        - File</span></span><br><span data-ttu-id="86cac-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-160">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-162">
        - Selection</span></span><br><span data-ttu-id="86cac-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-163">
        - Settings</span></span><br><span data-ttu-id="86cac-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-164">
        - TableBindings</span></span><br><span data-ttu-id="86cac-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-165">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-166">
        - TextBindings</span></span><br><span data-ttu-id="86cac-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-168">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-168">Office 2019 on Windows</span></span><br><span data-ttu-id="86cac-169">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="86cac-170">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-170">- TaskPane</span></span><br><span data-ttu-id="86cac-171">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-171">
        - Content</span></span><br><span data-ttu-id="86cac-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="86cac-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="86cac-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="86cac-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="86cac-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="86cac-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="86cac-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="86cac-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="86cac-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="86cac-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="86cac-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="86cac-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-182">- BindingEvents</span></span><br><span data-ttu-id="86cac-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-183">
        - CompressedFile</span></span><br><span data-ttu-id="86cac-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-184">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-185">
        - File</span></span><br><span data-ttu-id="86cac-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-186">
        - ImageCoercion</span></span><br><span data-ttu-id="86cac-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-187">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-189">
        - Selection</span></span><br><span data-ttu-id="86cac-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-190">
        - Settings</span></span><br><span data-ttu-id="86cac-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-191">
        - TableBindings</span></span><br><span data-ttu-id="86cac-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-192">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-193">
        - TextBindings</span></span><br><span data-ttu-id="86cac-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-195">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-195">Office 2016 on Windows</span></span><br><span data-ttu-id="86cac-196">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="86cac-197">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-197">- TaskPane</span></span><br><span data-ttu-id="86cac-198">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-198">
        - Content</span></span></td>
    <td><span data-ttu-id="86cac-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="86cac-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="86cac-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="86cac-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-201">- BindingEvents</span></span><br><span data-ttu-id="86cac-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-202">
        - CompressedFile</span></span><br><span data-ttu-id="86cac-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-203">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-204">
        - File</span></span><br><span data-ttu-id="86cac-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-205">
        - ImageCoercion</span></span><br><span data-ttu-id="86cac-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-206">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-208">
        - Selection</span></span><br><span data-ttu-id="86cac-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-209">
        - Settings</span></span><br><span data-ttu-id="86cac-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-210">
        - TableBindings</span></span><br><span data-ttu-id="86cac-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-211">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-212">
        - TextBindings</span></span><br><span data-ttu-id="86cac-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-214">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="86cac-214">Office 2013 on Windows</span></span><br><span data-ttu-id="86cac-215">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="86cac-216">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-216">
        - TaskPane</span></span><br><span data-ttu-id="86cac-217">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="86cac-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="86cac-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="86cac-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-219">
        - BindingEvents</span></span><br><span data-ttu-id="86cac-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-220">
        - CompressedFile</span></span><br><span data-ttu-id="86cac-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-221">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-222">
        - File</span></span><br><span data-ttu-id="86cac-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-223">
        - ImageCoercion</span></span><br><span data-ttu-id="86cac-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-224">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-226">
        - Selection</span></span><br><span data-ttu-id="86cac-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-227">
        - Settings</span></span><br><span data-ttu-id="86cac-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-228">
        - TableBindings</span></span><br><span data-ttu-id="86cac-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-229">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-230">
        - TextBindings</span></span><br><span data-ttu-id="86cac-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-232">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-232">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="86cac-233">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-233">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="86cac-234">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-234">- TaskPane</span></span><br><span data-ttu-id="86cac-235">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-235">
        - Content</span></span><br><span data-ttu-id="86cac-236">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="86cac-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="86cac-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="86cac-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="86cac-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="86cac-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="86cac-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="86cac-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="86cac-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="86cac-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="86cac-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="86cac-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="86cac-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="86cac-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="86cac-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-247">- BindingEvents</span></span><br><span data-ttu-id="86cac-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-248">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-249">
        - File</span></span><br><span data-ttu-id="86cac-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-250">
        - ImageCoercion</span></span><br><span data-ttu-id="86cac-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-251">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-253">
        - Selection</span></span><br><span data-ttu-id="86cac-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-254">
        - Settings</span></span><br><span data-ttu-id="86cac-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-255">
        - TableBindings</span></span><br><span data-ttu-id="86cac-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-256">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-257">
        - TextBindings</span></span><br><span data-ttu-id="86cac-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-259">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-259">Office apps on Mac</span></span><br><span data-ttu-id="86cac-260">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-260">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="86cac-261">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-261">- TaskPane</span></span><br><span data-ttu-id="86cac-262">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-262">
        - Content</span></span><br><span data-ttu-id="86cac-263">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="86cac-263">
        - Custom Functions</span></span><br><span data-ttu-id="86cac-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="86cac-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="86cac-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="86cac-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="86cac-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="86cac-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="86cac-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="86cac-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="86cac-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="86cac-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="86cac-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="86cac-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="86cac-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="86cac-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-275">- BindingEvents</span></span><br><span data-ttu-id="86cac-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-276">
        - CompressedFile</span></span><br><span data-ttu-id="86cac-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-277">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-278">
        - File</span></span><br><span data-ttu-id="86cac-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-279">
        - ImageCoercion</span></span><br><span data-ttu-id="86cac-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-280">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-282">
        - PdfFile</span></span><br><span data-ttu-id="86cac-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-283">
        - Selection</span></span><br><span data-ttu-id="86cac-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-284">
        - Settings</span></span><br><span data-ttu-id="86cac-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-285">
        - TableBindings</span></span><br><span data-ttu-id="86cac-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-286">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-287">
        - TextBindings</span></span><br><span data-ttu-id="86cac-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-289">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-289">Office 2019 for Mac</span></span><br><span data-ttu-id="86cac-290">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="86cac-291">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-291">- TaskPane</span></span><br><span data-ttu-id="86cac-292">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-292">
        - Content</span></span><br><span data-ttu-id="86cac-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="86cac-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="86cac-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="86cac-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="86cac-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="86cac-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="86cac-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="86cac-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="86cac-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="86cac-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="86cac-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="86cac-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-303">- BindingEvents</span></span><br><span data-ttu-id="86cac-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-304">
        - CompressedFile</span></span><br><span data-ttu-id="86cac-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-305">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-306">
        - File</span></span><br><span data-ttu-id="86cac-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-307">
        - ImageCoercion</span></span><br><span data-ttu-id="86cac-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-308">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-310">
        - PdfFile</span></span><br><span data-ttu-id="86cac-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-311">
        - Selection</span></span><br><span data-ttu-id="86cac-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-312">
        - Settings</span></span><br><span data-ttu-id="86cac-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-313">
        - TableBindings</span></span><br><span data-ttu-id="86cac-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-314">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-315">
        - TextBindings</span></span><br><span data-ttu-id="86cac-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-317">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-317">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="86cac-318">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="86cac-319">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-319">- TaskPane</span></span><br><span data-ttu-id="86cac-320">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-320">
        - Content</span></span></td>
    <td><span data-ttu-id="86cac-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="86cac-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="86cac-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="86cac-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-323">- BindingEvents</span></span><br><span data-ttu-id="86cac-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-324">
        - CompressedFile</span></span><br><span data-ttu-id="86cac-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-325">
        - DocumentEvents</span></span><br><span data-ttu-id="86cac-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="86cac-326">
        - File</span></span><br><span data-ttu-id="86cac-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-327">
        - ImageCoercion</span></span><br><span data-ttu-id="86cac-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-328">
        - MatrixBindings</span></span><br><span data-ttu-id="86cac-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="86cac-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-330">
        - PdfFile</span></span><br><span data-ttu-id="86cac-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-331">
        - Selection</span></span><br><span data-ttu-id="86cac-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-332">
        - Settings</span></span><br><span data-ttu-id="86cac-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-333">
        - TableBindings</span></span><br><span data-ttu-id="86cac-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-334">
        - TableCoercion</span></span><br><span data-ttu-id="86cac-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-335">
        - TextBindings</span></span><br><span data-ttu-id="86cac-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="86cac-337">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="86cac-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="86cac-338">自定义函数</span><span class="sxs-lookup"><span data-stu-id="86cac-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="86cac-339">平台</span><span class="sxs-lookup"><span data-stu-id="86cac-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="86cac-340">扩展点</span><span class="sxs-lookup"><span data-stu-id="86cac-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="86cac-341">API 要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="86cac-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="86cac-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-343">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="86cac-343">Office on the web</span></span></td>
    <td><span data-ttu-id="86cac-344">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="86cac-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="86cac-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-346">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-346">Office on Windows</span></span><br><span data-ttu-id="86cac-347">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-347">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="86cac-348">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="86cac-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="86cac-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-350">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="86cac-350">Office for Mac</span></span><br><span data-ttu-id="86cac-351">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="86cac-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="86cac-352">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="86cac-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="86cac-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="86cac-354">Outlook</span><span class="sxs-lookup"><span data-stu-id="86cac-354">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="86cac-355">平台</span><span class="sxs-lookup"><span data-stu-id="86cac-355">Platform</span></span></th>
    <th><span data-ttu-id="86cac-356">扩展点</span><span class="sxs-lookup"><span data-stu-id="86cac-356">Extension points</span></span></th>
    <th><span data-ttu-id="86cac-357">API 要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-357">API requirement sets</span></span></th>
    <th><span data-ttu-id="86cac-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="86cac-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-359">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="86cac-359">Office on the web</span></span><br><span data-ttu-id="86cac-360">（新）</span><span class="sxs-lookup"><span data-stu-id="86cac-360">New</span></span></td>
    <td> <span data-ttu-id="86cac-361">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-361">- Mail Read</span></span><br><span data-ttu-id="86cac-362">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-362">
      - Mail Compose</span></span><br><span data-ttu-id="86cac-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="86cac-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="86cac-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="86cac-371">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-372">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="86cac-372">Office on the web</span></span><br><span data-ttu-id="86cac-373">（经典）</span><span class="sxs-lookup"><span data-stu-id="86cac-373">Classic.</span></span></td>
    <td> <span data-ttu-id="86cac-374">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-374">- Mail Read</span></span><br><span data-ttu-id="86cac-375">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-375">
      - Mail Compose</span></span><br><span data-ttu-id="86cac-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="86cac-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="86cac-383">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-384">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-384">Office on Windows</span></span><br><span data-ttu-id="86cac-385">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-385">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-386">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-386">- Mail Read</span></span><br><span data-ttu-id="86cac-387">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-387">
      - Mail Compose</span></span><br><span data-ttu-id="86cac-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="86cac-389">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="86cac-389">
      - Modules</span></span></td>
    <td> <span data-ttu-id="86cac-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="86cac-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="86cac-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="86cac-397">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-397">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-398">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-398">Office 2019 on Windows</span></span><br><span data-ttu-id="86cac-399">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-399">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-400">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-400">- Mail Read</span></span><br><span data-ttu-id="86cac-401">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-401">
      - Mail Compose</span></span><br><span data-ttu-id="86cac-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="86cac-403">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="86cac-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="86cac-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="86cac-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="86cac-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="86cac-411">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-411">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-412">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-412">Office 2016 on Windows</span></span><br><span data-ttu-id="86cac-413">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-413">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-414">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-414">- Mail Read</span></span><br><span data-ttu-id="86cac-415">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-415">
      - Mail Compose</span></span><br><span data-ttu-id="86cac-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="86cac-417">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="86cac-417">
      - Modules</span></span></td>
    <td> <span data-ttu-id="86cac-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="86cac-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="86cac-422">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-423">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="86cac-423">Office 2013 on Windows</span></span><br><span data-ttu-id="86cac-424">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-424">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-425">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-425">- Mail Read</span></span><br><span data-ttu-id="86cac-426">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-426">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="86cac-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="86cac-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="86cac-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="86cac-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="86cac-431">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-432">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-432">Office apps on iOS</span></span><br><span data-ttu-id="86cac-433">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-433">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-434">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-434">- Mail Read</span></span><br><span data-ttu-id="86cac-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="86cac-441">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-442">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-442">Office apps on Mac</span></span><br><span data-ttu-id="86cac-443">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-443">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-444">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-444">- Mail Read</span></span><br><span data-ttu-id="86cac-445">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-445">
      - Mail Compose</span></span><br><span data-ttu-id="86cac-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="86cac-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="86cac-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="86cac-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="86cac-454">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-454">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-455">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-455">Office 2019 for Mac</span></span><br><span data-ttu-id="86cac-456">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-456">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-457">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-457">- Mail Read</span></span><br><span data-ttu-id="86cac-458">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-458">
      - Mail Compose</span></span><br><span data-ttu-id="86cac-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="86cac-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="86cac-466">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-467">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-467">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="86cac-468">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-468">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-469">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-469">- Mail Read</span></span><br><span data-ttu-id="86cac-470">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="86cac-470">
      - Mail Compose</span></span><br><span data-ttu-id="86cac-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="86cac-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="86cac-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="86cac-478">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-479">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-479">Office apps on Android</span></span><br><span data-ttu-id="86cac-480">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-480">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-481">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="86cac-481">- Mail Read</span></span><br><span data-ttu-id="86cac-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="86cac-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="86cac-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="86cac-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="86cac-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="86cac-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="86cac-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="86cac-488">不可用</span><span class="sxs-lookup"><span data-stu-id="86cac-488">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="86cac-489">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="86cac-489">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="86cac-490">Word</span><span class="sxs-lookup"><span data-stu-id="86cac-490">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="86cac-491">平台</span><span class="sxs-lookup"><span data-stu-id="86cac-491">Platform</span></span></th>
    <th><span data-ttu-id="86cac-492">扩展点</span><span class="sxs-lookup"><span data-stu-id="86cac-492">Extension points</span></span></th>
    <th><span data-ttu-id="86cac-493">API 要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-493">API requirement sets</span></span></th>
    <th><span data-ttu-id="86cac-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="86cac-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-495">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="86cac-495">Office on the web</span></span></td>
    <td> <span data-ttu-id="86cac-496">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-496">- TaskPane</span></span><br><span data-ttu-id="86cac-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="86cac-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="86cac-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="86cac-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-502">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-502">- BindingEvents</span></span><br><span data-ttu-id="86cac-503">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-503">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-504">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-504">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-505">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-505">
         - File</span></span><br><span data-ttu-id="86cac-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-506">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-507">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-508">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-508">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-509">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-509">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-510">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-510">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-511">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-511">
         - PdfFile</span></span><br><span data-ttu-id="86cac-512">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-512">
         - Selection</span></span><br><span data-ttu-id="86cac-513">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-513">
         - Settings</span></span><br><span data-ttu-id="86cac-514">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-514">
         - TableBindings</span></span><br><span data-ttu-id="86cac-515">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-515">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-516">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-516">
         - TextBindings</span></span><br><span data-ttu-id="86cac-517">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-517">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-518">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-518">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-519">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-519">Office on Windows</span></span><br><span data-ttu-id="86cac-520">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-520">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-521">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-521">- TaskPane</span></span><br><span data-ttu-id="86cac-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="86cac-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="86cac-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="86cac-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-527">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-527">- BindingEvents</span></span><br><span data-ttu-id="86cac-528">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-528">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-529">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-529">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-530">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-530">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-531">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-531">
         - File</span></span><br><span data-ttu-id="86cac-532">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-532">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-533">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-533">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-534">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-534">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-535">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-535">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-536">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-536">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-537">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-537">
         - PdfFile</span></span><br><span data-ttu-id="86cac-538">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-538">
         - Selection</span></span><br><span data-ttu-id="86cac-539">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-539">
         - Settings</span></span><br><span data-ttu-id="86cac-540">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-540">
         - TableBindings</span></span><br><span data-ttu-id="86cac-541">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-541">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-542">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-542">
         - TextBindings</span></span><br><span data-ttu-id="86cac-543">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-543">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-544">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-544">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-545">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-545">Office 2019 on Windows</span></span><br><span data-ttu-id="86cac-546">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-546">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-547">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-547">- TaskPane</span></span><br><span data-ttu-id="86cac-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="86cac-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="86cac-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="86cac-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-553">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-553">- BindingEvents</span></span><br><span data-ttu-id="86cac-554">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-554">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-555">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-555">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-556">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-557">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-557">
         - File</span></span><br><span data-ttu-id="86cac-558">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-558">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-559">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-559">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-560">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-560">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-561">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-561">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-562">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-562">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-563">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-563">
         - PdfFile</span></span><br><span data-ttu-id="86cac-564">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-564">
         - Selection</span></span><br><span data-ttu-id="86cac-565">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-565">
         - Settings</span></span><br><span data-ttu-id="86cac-566">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-566">
         - TableBindings</span></span><br><span data-ttu-id="86cac-567">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-567">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-568">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-568">
         - TextBindings</span></span><br><span data-ttu-id="86cac-569">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-569">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-570">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-570">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-571">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-571">Office 2016 on Windows</span></span><br><span data-ttu-id="86cac-572">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-572">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-573">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-573">- TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="86cac-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="86cac-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="86cac-576">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-576">- BindingEvents</span></span><br><span data-ttu-id="86cac-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-577">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-578">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-578">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-579">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-580">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-580">
         - File</span></span><br><span data-ttu-id="86cac-581">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-581">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-582">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-583">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-583">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-584">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-584">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-585">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-585">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-586">
         - PdfFile</span></span><br><span data-ttu-id="86cac-587">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-587">
         - Selection</span></span><br><span data-ttu-id="86cac-588">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-588">
         - Settings</span></span><br><span data-ttu-id="86cac-589">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-589">
         - TableBindings</span></span><br><span data-ttu-id="86cac-590">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-590">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-591">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-591">
         - TextBindings</span></span><br><span data-ttu-id="86cac-592">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-592">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-593">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-593">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-594">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="86cac-594">Office 2013 on Windows</span></span><br><span data-ttu-id="86cac-595">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-595">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-596">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-596">- TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="86cac-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="86cac-598">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-598">- BindingEvents</span></span><br><span data-ttu-id="86cac-599">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-599">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-600">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-600">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-601">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-601">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-602">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-602">
         - File</span></span><br><span data-ttu-id="86cac-603">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-603">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-604">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-604">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-605">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-605">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-606">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-606">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-607">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-607">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-608">
         - PdfFile</span></span><br><span data-ttu-id="86cac-609">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-609">
         - Selection</span></span><br><span data-ttu-id="86cac-610">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-610">
         - Settings</span></span><br><span data-ttu-id="86cac-611">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-611">
         - TableBindings</span></span><br><span data-ttu-id="86cac-612">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-612">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-613">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-613">
         - TextBindings</span></span><br><span data-ttu-id="86cac-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-614">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-615">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-615">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-616">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-616">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="86cac-617">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-617">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-618">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-618">- TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="86cac-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="86cac-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="86cac-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="86cac-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="86cac-623">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-623">- BindingEvents</span></span><br><span data-ttu-id="86cac-624">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-624">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-625">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-625">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-626">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-627">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-627">
         - File</span></span><br><span data-ttu-id="86cac-628">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-628">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-629">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-629">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-630">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-630">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-631">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-631">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-632">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-632">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-633">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-633">
         - PdfFile</span></span><br><span data-ttu-id="86cac-634">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-634">
         - Selection</span></span><br><span data-ttu-id="86cac-635">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-635">
         - Settings</span></span><br><span data-ttu-id="86cac-636">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-636">
         - TableBindings</span></span><br><span data-ttu-id="86cac-637">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-637">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-638">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-638">
         - TextBindings</span></span><br><span data-ttu-id="86cac-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-639">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-640">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-640">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-641">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-641">Office apps on Mac</span></span><br><span data-ttu-id="86cac-642">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-642">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-643">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-643">- TaskPane</span></span><br><span data-ttu-id="86cac-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="86cac-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="86cac-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="86cac-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="86cac-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="86cac-649">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-649">- BindingEvents</span></span><br><span data-ttu-id="86cac-650">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-650">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-651">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-651">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-652">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-652">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-653">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-653">
         - File</span></span><br><span data-ttu-id="86cac-654">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-654">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-655">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-655">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-656">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-656">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-657">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-657">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-658">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-658">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-659">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-659">
         - PdfFile</span></span><br><span data-ttu-id="86cac-660">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-660">
         - Selection</span></span><br><span data-ttu-id="86cac-661">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-661">
         - Settings</span></span><br><span data-ttu-id="86cac-662">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-662">
         - TableBindings</span></span><br><span data-ttu-id="86cac-663">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-663">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-664">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-664">
         - TextBindings</span></span><br><span data-ttu-id="86cac-665">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-665">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-666">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-666">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-667">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-667">Office 2019 for Mac</span></span><br><span data-ttu-id="86cac-668">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-668">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-669">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-669">- TaskPane</span></span><br><span data-ttu-id="86cac-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="86cac-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="86cac-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="86cac-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="86cac-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="86cac-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="86cac-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="86cac-675">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-675">- BindingEvents</span></span><br><span data-ttu-id="86cac-676">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-676">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-677">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-677">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-678">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-678">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-679">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-679">
         - File</span></span><br><span data-ttu-id="86cac-680">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-680">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-681">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-681">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-682">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-682">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-683">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-683">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-684">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-684">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-685">
         - PdfFile</span></span><br><span data-ttu-id="86cac-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-686">
         - Selection</span></span><br><span data-ttu-id="86cac-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-687">
         - Settings</span></span><br><span data-ttu-id="86cac-688">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-688">
         - TableBindings</span></span><br><span data-ttu-id="86cac-689">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-689">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-690">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-690">
         - TextBindings</span></span><br><span data-ttu-id="86cac-691">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-691">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-692">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-692">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-693">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-693">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="86cac-694">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-694">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-695">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-695">- TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="86cac-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="86cac-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="86cac-698">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-698">- BindingEvents</span></span><br><span data-ttu-id="86cac-699">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-699">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-700">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="86cac-700">
         - CustomXmlParts</span></span><br><span data-ttu-id="86cac-701">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-701">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-702">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-702">
         - File</span></span><br><span data-ttu-id="86cac-703">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-703">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-704">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-704">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-705">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-705">
         - MatrixBindings</span></span><br><span data-ttu-id="86cac-706">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-706">
         - MatrixCoercion</span></span><br><span data-ttu-id="86cac-707">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-707">
         - OoxmlCoercion</span></span><br><span data-ttu-id="86cac-708">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-708">
         - PdfFile</span></span><br><span data-ttu-id="86cac-709">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-709">
         - Selection</span></span><br><span data-ttu-id="86cac-710">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-710">
         - Settings</span></span><br><span data-ttu-id="86cac-711">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-711">
         - TableBindings</span></span><br><span data-ttu-id="86cac-712">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-712">
         - TableCoercion</span></span><br><span data-ttu-id="86cac-713">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="86cac-713">
         - TextBindings</span></span><br><span data-ttu-id="86cac-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-714">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-715">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="86cac-715">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="86cac-716">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="86cac-716">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="86cac-717">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="86cac-717">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="86cac-718">平台</span><span class="sxs-lookup"><span data-stu-id="86cac-718">Platform</span></span></th>
    <th><span data-ttu-id="86cac-719">扩展点</span><span class="sxs-lookup"><span data-stu-id="86cac-719">Extension points</span></span></th>
    <th><span data-ttu-id="86cac-720">API 要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-720">API requirement sets</span></span></th>
    <th><span data-ttu-id="86cac-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="86cac-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-722">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="86cac-722">Office on the web</span></span></td>
    <td> <span data-ttu-id="86cac-723">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-723">- Content</span></span><br><span data-ttu-id="86cac-724">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-724">
         - TaskPane</span></span><br><span data-ttu-id="86cac-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-727">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-727">- ActiveView</span></span><br><span data-ttu-id="86cac-728">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-728">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-729">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-729">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-730">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-730">
         - File</span></span><br><span data-ttu-id="86cac-731">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-731">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-732">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-732">
         - PdfFile</span></span><br><span data-ttu-id="86cac-733">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-733">
         - Selection</span></span><br><span data-ttu-id="86cac-734">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-734">
         - Settings</span></span><br><span data-ttu-id="86cac-735">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-735">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-736">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-736">Office on Windows</span></span><br><span data-ttu-id="86cac-737">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-737">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-738">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-738">- Content</span></span><br><span data-ttu-id="86cac-739">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-739">
         - TaskPane</span></span><br><span data-ttu-id="86cac-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-742">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-742">- ActiveView</span></span><br><span data-ttu-id="86cac-743">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-743">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-744">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-744">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-745">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-745">
         - File</span></span><br><span data-ttu-id="86cac-746">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-746">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-747">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-747">
         - PdfFile</span></span><br><span data-ttu-id="86cac-748">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-748">
         - Selection</span></span><br><span data-ttu-id="86cac-749">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-749">
         - Settings</span></span><br><span data-ttu-id="86cac-750">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-750">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-751">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-751">Office 2019 on Windows</span></span><br><span data-ttu-id="86cac-752">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-752">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-753">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-753">- Content</span></span><br><span data-ttu-id="86cac-754">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-754">
         - TaskPane</span></span><br><span data-ttu-id="86cac-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-757">- ActiveView</span></span><br><span data-ttu-id="86cac-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-758">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-759">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-760">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-760">
         - File</span></span><br><span data-ttu-id="86cac-761">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-761">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-762">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-762">
         - PdfFile</span></span><br><span data-ttu-id="86cac-763">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-763">
         - Selection</span></span><br><span data-ttu-id="86cac-764">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-764">
         - Settings</span></span><br><span data-ttu-id="86cac-765">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-765">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-766">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-766">Office 2016 on Windows</span></span><br><span data-ttu-id="86cac-767">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-767">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-768">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-768">- Content</span></span><br><span data-ttu-id="86cac-769">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-769">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="86cac-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="86cac-771">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-771">- ActiveView</span></span><br><span data-ttu-id="86cac-772">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-772">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-773">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-773">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-774">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-774">
         - File</span></span><br><span data-ttu-id="86cac-775">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-775">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-776">
         - PdfFile</span></span><br><span data-ttu-id="86cac-777">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-777">
         - Selection</span></span><br><span data-ttu-id="86cac-778">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-778">
         - Settings</span></span><br><span data-ttu-id="86cac-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-780">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="86cac-780">Office 2013 on Windows</span></span><br><span data-ttu-id="86cac-781">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-782">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-782">- Content</span></span><br><span data-ttu-id="86cac-783">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-783">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="86cac-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="86cac-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="86cac-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-785">- ActiveView</span></span><br><span data-ttu-id="86cac-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-786">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-787">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-788">
         - File</span></span><br><span data-ttu-id="86cac-789">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-789">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-790">
         - PdfFile</span></span><br><span data-ttu-id="86cac-791">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-791">
         - Selection</span></span><br><span data-ttu-id="86cac-792">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-792">
         - Settings</span></span><br><span data-ttu-id="86cac-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-794">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-794">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="86cac-795">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-795">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-796">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-796">- Content</span></span><br><span data-ttu-id="86cac-797">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-797">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-799">- ActiveView</span></span><br><span data-ttu-id="86cac-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-800">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-801">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-802">
         - File</span></span><br><span data-ttu-id="86cac-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-803">
         - PdfFile</span></span><br><span data-ttu-id="86cac-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-804">
         - Selection</span></span><br><span data-ttu-id="86cac-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-805">
         - Settings</span></span><br><span data-ttu-id="86cac-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-806">
         - TextCoercion</span></span><br><span data-ttu-id="86cac-807">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-807">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-808">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="86cac-808">Office apps on Mac</span></span><br><span data-ttu-id="86cac-809">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="86cac-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="86cac-810">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-810">- Content</span></span><br><span data-ttu-id="86cac-811">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-811">
         - TaskPane</span></span><br><span data-ttu-id="86cac-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-814">- ActiveView</span></span><br><span data-ttu-id="86cac-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-815">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-816">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-817">
         - File</span></span><br><span data-ttu-id="86cac-818">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-818">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-819">
         - PdfFile</span></span><br><span data-ttu-id="86cac-820">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-820">
         - Selection</span></span><br><span data-ttu-id="86cac-821">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-821">
         - Settings</span></span><br><span data-ttu-id="86cac-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-823">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-823">Office 2019 for Mac</span></span><br><span data-ttu-id="86cac-824">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-824">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-825">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-825">- Content</span></span><br><span data-ttu-id="86cac-826">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-826">
         - TaskPane</span></span><br><span data-ttu-id="86cac-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-829">- ActiveView</span></span><br><span data-ttu-id="86cac-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-830">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-831">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-832">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-832">
         - File</span></span><br><span data-ttu-id="86cac-833">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-833">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-834">
         - PdfFile</span></span><br><span data-ttu-id="86cac-835">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-835">
         - Selection</span></span><br><span data-ttu-id="86cac-836">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-836">
         - Settings</span></span><br><span data-ttu-id="86cac-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-838">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-838">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="86cac-839">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-840">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-840">- Content</span></span><br><span data-ttu-id="86cac-841">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-841">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="86cac-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="86cac-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="86cac-843">- ActiveView</span></span><br><span data-ttu-id="86cac-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="86cac-844">
         - CompressedFile</span></span><br><span data-ttu-id="86cac-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-845">
         - DocumentEvents</span></span><br><span data-ttu-id="86cac-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="86cac-846">
         - File</span></span><br><span data-ttu-id="86cac-847">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-847">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="86cac-848">
         - PdfFile</span></span><br><span data-ttu-id="86cac-849">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-849">
         - Selection</span></span><br><span data-ttu-id="86cac-850">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-850">
         - Settings</span></span><br><span data-ttu-id="86cac-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-851">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="86cac-852">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="86cac-852">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="86cac-853">OneNote</span><span class="sxs-lookup"><span data-stu-id="86cac-853">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="86cac-854">平台</span><span class="sxs-lookup"><span data-stu-id="86cac-854">Platform</span></span></th>
    <th><span data-ttu-id="86cac-855">扩展点</span><span class="sxs-lookup"><span data-stu-id="86cac-855">Extension points</span></span></th>
    <th><span data-ttu-id="86cac-856">API 要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-856">API requirement sets</span></span></th>
    <th><span data-ttu-id="86cac-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="86cac-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-858">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="86cac-858">Office on the web</span></span></td>
    <td> <span data-ttu-id="86cac-859">- 内容</span><span class="sxs-lookup"><span data-stu-id="86cac-859">- Content</span></span><br><span data-ttu-id="86cac-860">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-860">
         - TaskPane</span></span><br><span data-ttu-id="86cac-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="86cac-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="86cac-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="86cac-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-864">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="86cac-864">- DocumentEvents</span></span><br><span data-ttu-id="86cac-865">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-865">
         - HtmlCoercion</span></span><br><span data-ttu-id="86cac-866">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-866">
         - ImageCoercion</span></span><br><span data-ttu-id="86cac-867">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="86cac-867">
         - Settings</span></span><br><span data-ttu-id="86cac-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="86cac-869">项目</span><span class="sxs-lookup"><span data-stu-id="86cac-869">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="86cac-870">平台</span><span class="sxs-lookup"><span data-stu-id="86cac-870">Platform</span></span></th>
    <th><span data-ttu-id="86cac-871">扩展点</span><span class="sxs-lookup"><span data-stu-id="86cac-871">Extension points</span></span></th>
    <th><span data-ttu-id="86cac-872">API 要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-872">API requirement sets</span></span></th>
    <th><span data-ttu-id="86cac-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="86cac-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-874">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="86cac-874">Office 2019 on Windows</span></span><br><span data-ttu-id="86cac-875">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-875">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-876">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-876">- TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-878">- Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-878">- Selection</span></span><br><span data-ttu-id="86cac-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-879">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-880">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="86cac-880">Office 2016 on Windows</span></span><br><span data-ttu-id="86cac-881">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-881">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-882">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-882">- TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-884">- Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-884">- Selection</span></span><br><span data-ttu-id="86cac-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-885">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="86cac-886">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="86cac-886">Office 2013 on Windows</span></span><br><span data-ttu-id="86cac-887">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="86cac-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="86cac-888">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="86cac-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="86cac-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="86cac-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="86cac-890">- Selection</span><span class="sxs-lookup"><span data-stu-id="86cac-890">- Selection</span></span><br><span data-ttu-id="86cac-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="86cac-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="86cac-892">另请参阅</span><span class="sxs-lookup"><span data-stu-id="86cac-892">See also</span></span>

- [<span data-ttu-id="86cac-893">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="86cac-893">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="86cac-894">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-894">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="86cac-895">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-895">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="86cac-896">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="86cac-896">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="86cac-897">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="86cac-897">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="86cac-898">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="86cac-898">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="86cac-899">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="86cac-899">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="86cac-900">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="86cac-900">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="86cac-901">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="86cac-901">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="86cac-902">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="86cac-902">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="86cac-903">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="86cac-903">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
