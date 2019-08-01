---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 07/26/2019
localization_priority: Priority
ms.openlocfilehash: 7039ca59af22f1101bdff7b6bcd4506497d6c9cd
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940834"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="16035-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="16035-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="16035-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="16035-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="16035-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="16035-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="16035-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="16035-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="16035-108">Excel</span><span class="sxs-lookup"><span data-stu-id="16035-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="16035-109">平台</span><span class="sxs-lookup"><span data-stu-id="16035-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="16035-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="16035-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="16035-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="16035-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="16035-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="16035-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="16035-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="16035-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-114">- TaskPane</span></span><br><span data-ttu-id="16035-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-115">
        - Content</span></span><br><span data-ttu-id="16035-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="16035-116">
        - Custom Functions</span></span><br><span data-ttu-id="16035-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="16035-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="16035-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16035-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16035-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16035-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16035-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16035-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16035-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16035-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16035-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16035-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="16035-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="16035-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="16035-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-130">
        - BindingEvents</span></span><br><span data-ttu-id="16035-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-131">
        - CompressedFile</span></span><br><span data-ttu-id="16035-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-132">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-133">
        - File</span></span><br><span data-ttu-id="16035-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-134">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-136">
        - Selection</span></span><br><span data-ttu-id="16035-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-137">
        - Settings</span></span><br><span data-ttu-id="16035-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-138">
        - TableBindings</span></span><br><span data-ttu-id="16035-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-139">
        - TableCoercion</span></span><br><span data-ttu-id="16035-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-140">
        - TextBindings</span></span><br><span data-ttu-id="16035-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-142">Office on Windows</span></span><br><span data-ttu-id="16035-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-144">- TaskPane</span></span><br><span data-ttu-id="16035-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-145">
        - Content</span></span><br><span data-ttu-id="16035-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="16035-146">
        - Custom Functions</span></span><br><span data-ttu-id="16035-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="16035-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="16035-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16035-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16035-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16035-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16035-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16035-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16035-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16035-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16035-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16035-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="16035-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="16035-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="16035-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-160">
        - BindingEvents</span></span><br><span data-ttu-id="16035-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-161">
        - CompressedFile</span></span><br><span data-ttu-id="16035-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-162">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-163">
        - File</span></span><br><span data-ttu-id="16035-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-164">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-166">
        - Selection</span></span><br><span data-ttu-id="16035-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-167">
        - Settings</span></span><br><span data-ttu-id="16035-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-168">
        - TableBindings</span></span><br><span data-ttu-id="16035-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-169">
        - TableCoercion</span></span><br><span data-ttu-id="16035-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-170">
        - TextBindings</span></span><br><span data-ttu-id="16035-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-172">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-172">Office 2019 on Windows</span></span><br><span data-ttu-id="16035-173">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16035-174">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-174">- TaskPane</span></span><br><span data-ttu-id="16035-175">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-175">
        - Content</span></span><br><span data-ttu-id="16035-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="16035-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16035-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16035-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16035-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16035-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16035-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16035-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16035-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16035-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16035-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16035-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-187">- BindingEvents</span></span><br><span data-ttu-id="16035-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-188">
        - CompressedFile</span></span><br><span data-ttu-id="16035-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-189">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-190">
        - File</span></span><br><span data-ttu-id="16035-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-191">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-193">
        - Selection</span></span><br><span data-ttu-id="16035-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-194">
        - Settings</span></span><br><span data-ttu-id="16035-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-195">
        - TableBindings</span></span><br><span data-ttu-id="16035-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-196">
        - TableCoercion</span></span><br><span data-ttu-id="16035-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-197">
        - TextBindings</span></span><br><span data-ttu-id="16035-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-199">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-199">Office 2016 on Windows</span></span><br><span data-ttu-id="16035-200">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16035-201">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-201">- TaskPane</span></span><br><span data-ttu-id="16035-202">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-202">
        - Content</span></span></td>
    <td><span data-ttu-id="16035-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16035-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="16035-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="16035-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16035-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-206">- BindingEvents</span></span><br><span data-ttu-id="16035-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-207">
        - CompressedFile</span></span><br><span data-ttu-id="16035-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-208">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-209">
        - File</span></span><br><span data-ttu-id="16035-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-210">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-212">
        - Selection</span></span><br><span data-ttu-id="16035-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-213">
        - Settings</span></span><br><span data-ttu-id="16035-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-214">
        - TableBindings</span></span><br><span data-ttu-id="16035-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-215">
        - TableCoercion</span></span><br><span data-ttu-id="16035-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-216">
        - TextBindings</span></span><br><span data-ttu-id="16035-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-218">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="16035-218">Office 2013 on Windows</span></span><br><span data-ttu-id="16035-219">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16035-220">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-220">
        - TaskPane</span></span><br><span data-ttu-id="16035-221">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="16035-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16035-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16035-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16035-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-224">
        - BindingEvents</span></span><br><span data-ttu-id="16035-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-225">
        - CompressedFile</span></span><br><span data-ttu-id="16035-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-226">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-227">
        - File</span></span><br><span data-ttu-id="16035-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-228">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-230">
        - Selection</span></span><br><span data-ttu-id="16035-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-231">
        - Settings</span></span><br><span data-ttu-id="16035-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-232">
        - TableBindings</span></span><br><span data-ttu-id="16035-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-233">
        - TableCoercion</span></span><br><span data-ttu-id="16035-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-234">
        - TextBindings</span></span><br><span data-ttu-id="16035-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-236">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="16035-237">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="16035-238">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-238">- TaskPane</span></span><br><span data-ttu-id="16035-239">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-239">
        - Content</span></span><br><span data-ttu-id="16035-240">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="16035-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="16035-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16035-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16035-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16035-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16035-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16035-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16035-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16035-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16035-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16035-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="16035-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="16035-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16035-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-252">- BindingEvents</span></span><br><span data-ttu-id="16035-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-253">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-254">
        - File</span></span><br><span data-ttu-id="16035-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-255">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-257">
        - Selection</span></span><br><span data-ttu-id="16035-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-258">
        - Settings</span></span><br><span data-ttu-id="16035-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-259">
        - TableBindings</span></span><br><span data-ttu-id="16035-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-260">
        - TableCoercion</span></span><br><span data-ttu-id="16035-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-261">
        - TextBindings</span></span><br><span data-ttu-id="16035-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-263">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-263">Office apps on Mac</span></span><br><span data-ttu-id="16035-264">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="16035-265">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-265">- TaskPane</span></span><br><span data-ttu-id="16035-266">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-266">
        - Content</span></span><br><span data-ttu-id="16035-267">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="16035-267">
        - Custom Functions</span></span><br><span data-ttu-id="16035-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="16035-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16035-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16035-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16035-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16035-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16035-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16035-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16035-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16035-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16035-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="16035-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="16035-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="16035-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-281">- BindingEvents</span></span><br><span data-ttu-id="16035-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-282">
        - CompressedFile</span></span><br><span data-ttu-id="16035-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-283">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-284">
        - File</span></span><br><span data-ttu-id="16035-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-285">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-287">
        - PdfFile</span></span><br><span data-ttu-id="16035-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-288">
        - Selection</span></span><br><span data-ttu-id="16035-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-289">
        - Settings</span></span><br><span data-ttu-id="16035-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-290">
        - TableBindings</span></span><br><span data-ttu-id="16035-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-291">
        - TableCoercion</span></span><br><span data-ttu-id="16035-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-292">
        - TextBindings</span></span><br><span data-ttu-id="16035-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-294">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-294">Office 2019 for Mac</span></span><br><span data-ttu-id="16035-295">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16035-296">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-296">- TaskPane</span></span><br><span data-ttu-id="16035-297">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-297">
        - Content</span></span><br><span data-ttu-id="16035-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="16035-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16035-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="16035-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="16035-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="16035-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="16035-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="16035-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="16035-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="16035-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="16035-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16035-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-309">- BindingEvents</span></span><br><span data-ttu-id="16035-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-310">
        - CompressedFile</span></span><br><span data-ttu-id="16035-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-311">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-312">
        - File</span></span><br><span data-ttu-id="16035-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-313">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-315">
        - PdfFile</span></span><br><span data-ttu-id="16035-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-316">
        - Selection</span></span><br><span data-ttu-id="16035-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-317">
        - Settings</span></span><br><span data-ttu-id="16035-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-318">
        - TableBindings</span></span><br><span data-ttu-id="16035-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-319">
        - TableCoercion</span></span><br><span data-ttu-id="16035-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-320">
        - TextBindings</span></span><br><span data-ttu-id="16035-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-322">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="16035-323">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="16035-324">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-324">- TaskPane</span></span><br><span data-ttu-id="16035-325">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="16035-325">
        - Content</span></span></td>
    <td><span data-ttu-id="16035-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="16035-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="16035-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="16035-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="16035-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-329">- BindingEvents</span></span><br><span data-ttu-id="16035-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-330">
        - CompressedFile</span></span><br><span data-ttu-id="16035-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-331">
        - DocumentEvents</span></span><br><span data-ttu-id="16035-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="16035-332">
        - File</span></span><br><span data-ttu-id="16035-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-333">
        - MatrixBindings</span></span><br><span data-ttu-id="16035-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="16035-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-335">
        - PdfFile</span></span><br><span data-ttu-id="16035-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-336">
        - Selection</span></span><br><span data-ttu-id="16035-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-337">
        - Settings</span></span><br><span data-ttu-id="16035-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-338">
        - TableBindings</span></span><br><span data-ttu-id="16035-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-339">
        - TableCoercion</span></span><br><span data-ttu-id="16035-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-340">
        - TextBindings</span></span><br><span data-ttu-id="16035-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="16035-342">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="16035-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="16035-343">自定义函数</span><span class="sxs-lookup"><span data-stu-id="16035-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="16035-344">平台</span><span class="sxs-lookup"><span data-stu-id="16035-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="16035-345">扩展点</span><span class="sxs-lookup"><span data-stu-id="16035-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="16035-346">API 要求集</span><span class="sxs-lookup"><span data-stu-id="16035-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="16035-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="16035-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-348">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="16035-348">Office on the web</span></span></td>
    <td><span data-ttu-id="16035-349">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="16035-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="16035-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-351">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-351">Office on Windows</span></span><br><span data-ttu-id="16035-352">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="16035-353">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="16035-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="16035-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-355">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="16035-355">Office for Mac</span></span><br><span data-ttu-id="16035-356">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="16035-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="16035-357">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="16035-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="16035-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="16035-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="16035-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16035-360">平台</span><span class="sxs-lookup"><span data-stu-id="16035-360">Platform</span></span></th>
    <th><span data-ttu-id="16035-361">扩展点</span><span class="sxs-lookup"><span data-stu-id="16035-361">Extension points</span></span></th>
    <th><span data-ttu-id="16035-362">API 要求集</span><span class="sxs-lookup"><span data-stu-id="16035-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="16035-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="16035-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-364">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="16035-364">Office on the web</span></span><br><span data-ttu-id="16035-365">（新式）</span><span class="sxs-lookup"><span data-stu-id="16035-365">Modern</span></span></td>
    <td> <span data-ttu-id="16035-366">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-366">- Mail Read</span></span><br><span data-ttu-id="16035-367">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-367">
      - Mail Compose</span></span><br><span data-ttu-id="16035-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16035-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="16035-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="16035-376">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-377">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="16035-377">Office on the web</span></span><br><span data-ttu-id="16035-378">（经典）</span><span class="sxs-lookup"><span data-stu-id="16035-378">Classic.</span></span></td>
    <td> <span data-ttu-id="16035-379">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-379">- Mail Read</span></span><br><span data-ttu-id="16035-380">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-380">
      - Mail Compose</span></span><br><span data-ttu-id="16035-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16035-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="16035-388">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-389">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-389">Office on Windows</span></span><br><span data-ttu-id="16035-390">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-391">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-391">- Mail Read</span></span><br><span data-ttu-id="16035-392">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-392">
      - Mail Compose</span></span><br><span data-ttu-id="16035-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="16035-394">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="16035-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="16035-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16035-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="16035-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="16035-402">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-403">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-403">Office 2019 on Windows</span></span><br><span data-ttu-id="16035-404">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-405">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-405">- Mail Read</span></span><br><span data-ttu-id="16035-406">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-406">
      - Mail Compose</span></span><br><span data-ttu-id="16035-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="16035-408">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="16035-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="16035-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16035-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="16035-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="16035-416">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-417">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-417">Office 2016 on Windows</span></span><br><span data-ttu-id="16035-418">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-419">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-419">- Mail Read</span></span><br><span data-ttu-id="16035-420">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-420">
      - Mail Compose</span></span><br><span data-ttu-id="16035-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="16035-422">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="16035-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="16035-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="16035-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="16035-427">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-428">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="16035-428">Office 2013 on Windows</span></span><br><span data-ttu-id="16035-429">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-430">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-430">- Mail Read</span></span><br><span data-ttu-id="16035-431">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="16035-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="16035-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="16035-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="16035-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="16035-436">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-437">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-437">Office apps on iOS</span></span><br><span data-ttu-id="16035-438">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-439">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-439">- Mail Read</span></span><br><span data-ttu-id="16035-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="16035-446">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-447">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-447">Office apps on Mac</span></span><br><span data-ttu-id="16035-448">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-449">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-449">- Mail Read</span></span><br><span data-ttu-id="16035-450">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-450">
      - Mail Compose</span></span><br><span data-ttu-id="16035-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16035-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="16035-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="16035-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="16035-459">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-460">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-460">Office 2019 for Mac</span></span><br><span data-ttu-id="16035-461">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-462">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-462">- Mail Read</span></span><br><span data-ttu-id="16035-463">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-463">
      - Mail Compose</span></span><br><span data-ttu-id="16035-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16035-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="16035-471">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-472">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="16035-473">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-474">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-474">- Mail Read</span></span><br><span data-ttu-id="16035-475">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="16035-475">
      - Mail Compose</span></span><br><span data-ttu-id="16035-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="16035-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="16035-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="16035-483">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-484">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-484">Office apps on Android</span></span><br><span data-ttu-id="16035-485">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-486">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="16035-486">- Mail Read</span></span><br><span data-ttu-id="16035-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="16035-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="16035-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="16035-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="16035-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="16035-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="16035-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="16035-493">不可用</span><span class="sxs-lookup"><span data-stu-id="16035-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="16035-494">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="16035-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="16035-495">Word</span><span class="sxs-lookup"><span data-stu-id="16035-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16035-496">平台</span><span class="sxs-lookup"><span data-stu-id="16035-496">Platform</span></span></th>
    <th><span data-ttu-id="16035-497">扩展点</span><span class="sxs-lookup"><span data-stu-id="16035-497">Extension points</span></span></th>
    <th><span data-ttu-id="16035-498">API 要求集</span><span class="sxs-lookup"><span data-stu-id="16035-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="16035-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="16035-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-500">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="16035-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="16035-501">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-501">- TaskPane</span></span><br><span data-ttu-id="16035-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16035-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16035-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16035-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16035-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-509">- BindingEvents</span></span><br><span data-ttu-id="16035-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-511">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-512">
         - File</span></span><br><span data-ttu-id="16035-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-514">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-517">
         - PdfFile</span></span><br><span data-ttu-id="16035-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-518">
         - Selection</span></span><br><span data-ttu-id="16035-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-519">
         - Settings</span></span><br><span data-ttu-id="16035-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-520">
         - TableBindings</span></span><br><span data-ttu-id="16035-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-521">
         - TableCoercion</span></span><br><span data-ttu-id="16035-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-522">
         - TextBindings</span></span><br><span data-ttu-id="16035-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-523">
         - TextCoercion</span></span><br><span data-ttu-id="16035-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-525">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-525">Office on Windows</span></span><br><span data-ttu-id="16035-526">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-527">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-527">- TaskPane</span></span><br><span data-ttu-id="16035-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16035-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16035-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16035-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16035-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-535">- BindingEvents</span></span><br><span data-ttu-id="16035-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-536">
         - CompressedFile</span></span><br><span data-ttu-id="16035-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-538">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-539">
         - File</span></span><br><span data-ttu-id="16035-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-541">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-544">
         - PdfFile</span></span><br><span data-ttu-id="16035-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-545">
         - Selection</span></span><br><span data-ttu-id="16035-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-546">
         - Settings</span></span><br><span data-ttu-id="16035-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-547">
         - TableBindings</span></span><br><span data-ttu-id="16035-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-548">
         - TableCoercion</span></span><br><span data-ttu-id="16035-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-549">
         - TextBindings</span></span><br><span data-ttu-id="16035-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-550">
         - TextCoercion</span></span><br><span data-ttu-id="16035-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-552">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-552">Office 2019 on Windows</span></span><br><span data-ttu-id="16035-553">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-554">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-554">- TaskPane</span></span><br><span data-ttu-id="16035-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16035-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16035-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16035-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-561">- BindingEvents</span></span><br><span data-ttu-id="16035-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-562">
         - CompressedFile</span></span><br><span data-ttu-id="16035-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-564">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-565">
         - File</span></span><br><span data-ttu-id="16035-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-567">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-570">
         - PdfFile</span></span><br><span data-ttu-id="16035-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-571">
         - Selection</span></span><br><span data-ttu-id="16035-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-572">
         - Settings</span></span><br><span data-ttu-id="16035-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-573">
         - TableBindings</span></span><br><span data-ttu-id="16035-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-574">
         - TableCoercion</span></span><br><span data-ttu-id="16035-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-575">
         - TextBindings</span></span><br><span data-ttu-id="16035-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-576">
         - TextCoercion</span></span><br><span data-ttu-id="16035-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-578">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-578">Office 2016 on Windows</span></span><br><span data-ttu-id="16035-579">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-580">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16035-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="16035-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="16035-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-584">- BindingEvents</span></span><br><span data-ttu-id="16035-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-585">
         - CompressedFile</span></span><br><span data-ttu-id="16035-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-587">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-588">
         - File</span></span><br><span data-ttu-id="16035-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-590">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-593">
         - PdfFile</span></span><br><span data-ttu-id="16035-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-594">
         - Selection</span></span><br><span data-ttu-id="16035-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-595">
         - Settings</span></span><br><span data-ttu-id="16035-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-596">
         - TableBindings</span></span><br><span data-ttu-id="16035-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-597">
         - TableCoercion</span></span><br><span data-ttu-id="16035-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-598">
         - TextBindings</span></span><br><span data-ttu-id="16035-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-599">
         - TextCoercion</span></span><br><span data-ttu-id="16035-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-601">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="16035-601">Office 2013 on Windows</span></span><br><span data-ttu-id="16035-602">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-603">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16035-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16035-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-606">- BindingEvents</span></span><br><span data-ttu-id="16035-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-607">
         - CompressedFile</span></span><br><span data-ttu-id="16035-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-609">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-610">
         - File</span></span><br><span data-ttu-id="16035-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-612">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-615">
         - PdfFile</span></span><br><span data-ttu-id="16035-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-616">
         - Selection</span></span><br><span data-ttu-id="16035-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-617">
         - Settings</span></span><br><span data-ttu-id="16035-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-618">
         - TableBindings</span></span><br><span data-ttu-id="16035-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-619">
         - TableCoercion</span></span><br><span data-ttu-id="16035-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-620">
         - TextBindings</span></span><br><span data-ttu-id="16035-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-621">
         - TextCoercion</span></span><br><span data-ttu-id="16035-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-623">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="16035-624">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-625">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16035-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16035-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16035-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="16035-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-631">- BindingEvents</span></span><br><span data-ttu-id="16035-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-632">
         - CompressedFile</span></span><br><span data-ttu-id="16035-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-634">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-635">
         - File</span></span><br><span data-ttu-id="16035-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-637">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-640">
         - PdfFile</span></span><br><span data-ttu-id="16035-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-641">
         - Selection</span></span><br><span data-ttu-id="16035-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-642">
         - Settings</span></span><br><span data-ttu-id="16035-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-643">
         - TableBindings</span></span><br><span data-ttu-id="16035-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-644">
         - TableCoercion</span></span><br><span data-ttu-id="16035-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-645">
         - TextBindings</span></span><br><span data-ttu-id="16035-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-646">
         - TextCoercion</span></span><br><span data-ttu-id="16035-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-648">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-648">Office apps on Mac</span></span><br><span data-ttu-id="16035-649">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-650">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-650">- TaskPane</span></span><br><span data-ttu-id="16035-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16035-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16035-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16035-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="16035-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-658">- BindingEvents</span></span><br><span data-ttu-id="16035-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-659">
         - CompressedFile</span></span><br><span data-ttu-id="16035-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-661">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-662">
         - File</span></span><br><span data-ttu-id="16035-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-664">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-667">
         - PdfFile</span></span><br><span data-ttu-id="16035-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-668">
         - Selection</span></span><br><span data-ttu-id="16035-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-669">
         - Settings</span></span><br><span data-ttu-id="16035-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-670">
         - TableBindings</span></span><br><span data-ttu-id="16035-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-671">
         - TableCoercion</span></span><br><span data-ttu-id="16035-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-672">
         - TextBindings</span></span><br><span data-ttu-id="16035-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-673">
         - TextCoercion</span></span><br><span data-ttu-id="16035-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-675">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-675">Office 2019 for Mac</span></span><br><span data-ttu-id="16035-676">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-677">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-677">- TaskPane</span></span><br><span data-ttu-id="16035-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16035-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="16035-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="16035-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="16035-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="16035-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-684">- BindingEvents</span></span><br><span data-ttu-id="16035-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-685">
         - CompressedFile</span></span><br><span data-ttu-id="16035-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-687">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-688">
         - File</span></span><br><span data-ttu-id="16035-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-690">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-693">
         - PdfFile</span></span><br><span data-ttu-id="16035-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-694">
         - Selection</span></span><br><span data-ttu-id="16035-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-695">
         - Settings</span></span><br><span data-ttu-id="16035-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-696">
         - TableBindings</span></span><br><span data-ttu-id="16035-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-697">
         - TableCoercion</span></span><br><span data-ttu-id="16035-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-698">
         - TextBindings</span></span><br><span data-ttu-id="16035-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-699">
         - TextCoercion</span></span><br><span data-ttu-id="16035-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-701">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="16035-702">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-703">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="16035-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="16035-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="16035-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16035-707">- BindingEvents</span></span><br><span data-ttu-id="16035-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-708">
         - CompressedFile</span></span><br><span data-ttu-id="16035-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16035-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="16035-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-710">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-711">
         - File</span></span><br><span data-ttu-id="16035-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16035-713">
         - MatrixBindings</span></span><br><span data-ttu-id="16035-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="16035-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="16035-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-716">
         - PdfFile</span></span><br><span data-ttu-id="16035-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-717">
         - Selection</span></span><br><span data-ttu-id="16035-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-718">
         - Settings</span></span><br><span data-ttu-id="16035-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="16035-719">
         - TableBindings</span></span><br><span data-ttu-id="16035-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-720">
         - TableCoercion</span></span><br><span data-ttu-id="16035-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="16035-721">
         - TextBindings</span></span><br><span data-ttu-id="16035-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-722">
         - TextCoercion</span></span><br><span data-ttu-id="16035-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="16035-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="16035-724">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="16035-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="16035-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="16035-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16035-726">平台</span><span class="sxs-lookup"><span data-stu-id="16035-726">Platform</span></span></th>
    <th><span data-ttu-id="16035-727">扩展点</span><span class="sxs-lookup"><span data-stu-id="16035-727">Extension points</span></span></th>
    <th><span data-ttu-id="16035-728">API 要求集</span><span class="sxs-lookup"><span data-stu-id="16035-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="16035-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="16035-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-730">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="16035-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="16035-731">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-731">- Content</span></span><br><span data-ttu-id="16035-732">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-732">
         - TaskPane</span></span><br><span data-ttu-id="16035-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="16035-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16035-738">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-738">- ActiveView</span></span><br><span data-ttu-id="16035-739">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-739">
         - CompressedFile</span></span><br><span data-ttu-id="16035-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-740">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-741">
         - File</span></span><br><span data-ttu-id="16035-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-742">
         - PdfFile</span></span><br><span data-ttu-id="16035-743">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-743">
         - Selection</span></span><br><span data-ttu-id="16035-744">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-744">
         - Settings</span></span><br><span data-ttu-id="16035-745">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-745">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-746">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-746">Office on Windows</span></span><br><span data-ttu-id="16035-747">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-747">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-748">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-748">- Content</span></span><br><span data-ttu-id="16035-749">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-749">
         - TaskPane</span></span><br><span data-ttu-id="16035-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="16035-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16035-755">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-755">- ActiveView</span></span><br><span data-ttu-id="16035-756">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-756">
         - CompressedFile</span></span><br><span data-ttu-id="16035-757">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-757">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-758">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-758">
         - File</span></span><br><span data-ttu-id="16035-759">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-759">
         - PdfFile</span></span><br><span data-ttu-id="16035-760">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-760">
         - Selection</span></span><br><span data-ttu-id="16035-761">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-761">
         - Settings</span></span><br><span data-ttu-id="16035-762">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-762">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-763">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-763">Office 2019 on Windows</span></span><br><span data-ttu-id="16035-764">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-764">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-765">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-765">- Content</span></span><br><span data-ttu-id="16035-766">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-766">
         - TaskPane</span></span><br><span data-ttu-id="16035-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-770">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-770">- ActiveView</span></span><br><span data-ttu-id="16035-771">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-771">
         - CompressedFile</span></span><br><span data-ttu-id="16035-772">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-772">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-773">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-773">
         - File</span></span><br><span data-ttu-id="16035-774">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-774">
         - PdfFile</span></span><br><span data-ttu-id="16035-775">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-775">
         - Selection</span></span><br><span data-ttu-id="16035-776">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-776">
         - Settings</span></span><br><span data-ttu-id="16035-777">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-777">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-778">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-778">Office 2016 on Windows</span></span><br><span data-ttu-id="16035-779">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-779">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-780">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-780">- Content</span></span><br><span data-ttu-id="16035-781">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-781">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16035-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16035-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-784">- ActiveView</span></span><br><span data-ttu-id="16035-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-785">
         - CompressedFile</span></span><br><span data-ttu-id="16035-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-786">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-787">
         - File</span></span><br><span data-ttu-id="16035-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-788">
         - PdfFile</span></span><br><span data-ttu-id="16035-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-789">
         - Selection</span></span><br><span data-ttu-id="16035-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-790">
         - Settings</span></span><br><span data-ttu-id="16035-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-792">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="16035-792">Office 2013 on Windows</span></span><br><span data-ttu-id="16035-793">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-794">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-794">- Content</span></span><br><span data-ttu-id="16035-795">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-795">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="16035-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16035-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16035-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-798">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-798">- ActiveView</span></span><br><span data-ttu-id="16035-799">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-799">
         - CompressedFile</span></span><br><span data-ttu-id="16035-800">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-800">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-801">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-801">
         - File</span></span><br><span data-ttu-id="16035-802">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-802">
         - PdfFile</span></span><br><span data-ttu-id="16035-803">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-803">
         - Selection</span></span><br><span data-ttu-id="16035-804">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-804">
         - Settings</span></span><br><span data-ttu-id="16035-805">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-805">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-806">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-806">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="16035-807">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-807">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-808">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-808">- Content</span></span><br><span data-ttu-id="16035-809">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-809">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="16035-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-813">- ActiveView</span></span><br><span data-ttu-id="16035-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-814">
         - CompressedFile</span></span><br><span data-ttu-id="16035-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-815">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-816">
         - File</span></span><br><span data-ttu-id="16035-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-817">
         - PdfFile</span></span><br><span data-ttu-id="16035-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-818">
         - Selection</span></span><br><span data-ttu-id="16035-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-819">
         - Settings</span></span><br><span data-ttu-id="16035-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-821">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="16035-821">Office apps on Mac</span></span><br><span data-ttu-id="16035-822">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="16035-822">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="16035-823">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-823">- Content</span></span><br><span data-ttu-id="16035-824">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-824">
         - TaskPane</span></span><br><span data-ttu-id="16035-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="16035-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="16035-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="16035-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="16035-830">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-830">- ActiveView</span></span><br><span data-ttu-id="16035-831">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-831">
         - CompressedFile</span></span><br><span data-ttu-id="16035-832">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-832">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-833">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-833">
         - File</span></span><br><span data-ttu-id="16035-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-834">
         - PdfFile</span></span><br><span data-ttu-id="16035-835">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-835">
         - Selection</span></span><br><span data-ttu-id="16035-836">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-836">
         - Settings</span></span><br><span data-ttu-id="16035-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-838">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-838">Office 2019 for Mac</span></span><br><span data-ttu-id="16035-839">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-840">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-840">- Content</span></span><br><span data-ttu-id="16035-841">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-841">
         - TaskPane</span></span><br><span data-ttu-id="16035-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-845">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-845">- ActiveView</span></span><br><span data-ttu-id="16035-846">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-846">
         - CompressedFile</span></span><br><span data-ttu-id="16035-847">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-847">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-848">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-848">
         - File</span></span><br><span data-ttu-id="16035-849">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-849">
         - PdfFile</span></span><br><span data-ttu-id="16035-850">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-850">
         - Selection</span></span><br><span data-ttu-id="16035-851">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-851">
         - Settings</span></span><br><span data-ttu-id="16035-852">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-852">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-853">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-853">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="16035-854">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-854">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-855">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-855">- Content</span></span><br><span data-ttu-id="16035-856">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-856">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="16035-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="16035-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="16035-859">- ActiveView</span></span><br><span data-ttu-id="16035-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16035-860">
         - CompressedFile</span></span><br><span data-ttu-id="16035-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-861">
         - DocumentEvents</span></span><br><span data-ttu-id="16035-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="16035-862">
         - File</span></span><br><span data-ttu-id="16035-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="16035-863">
         - PdfFile</span></span><br><span data-ttu-id="16035-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="16035-864">
         - Selection</span></span><br><span data-ttu-id="16035-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-865">
         - Settings</span></span><br><span data-ttu-id="16035-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-866">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="16035-867">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="16035-867">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="16035-868">OneNote</span><span class="sxs-lookup"><span data-stu-id="16035-868">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16035-869">平台</span><span class="sxs-lookup"><span data-stu-id="16035-869">Platform</span></span></th>
    <th><span data-ttu-id="16035-870">扩展点</span><span class="sxs-lookup"><span data-stu-id="16035-870">Extension points</span></span></th>
    <th><span data-ttu-id="16035-871">API 要求集</span><span class="sxs-lookup"><span data-stu-id="16035-871">API requirement sets</span></span></th>
    <th><span data-ttu-id="16035-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="16035-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-873">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="16035-873">Office on the web</span></span></td>
    <td> <span data-ttu-id="16035-874">- 内容</span><span class="sxs-lookup"><span data-stu-id="16035-874">- Content</span></span><br><span data-ttu-id="16035-875">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-875">
         - TaskPane</span></span><br><span data-ttu-id="16035-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="16035-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="16035-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="16035-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="16035-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-880">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16035-880">- DocumentEvents</span></span><br><span data-ttu-id="16035-881">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-881">
         - HtmlCoercion</span></span><br><span data-ttu-id="16035-882">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="16035-882">
         - Settings</span></span><br><span data-ttu-id="16035-883">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-883">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="16035-884">项目</span><span class="sxs-lookup"><span data-stu-id="16035-884">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16035-885">平台</span><span class="sxs-lookup"><span data-stu-id="16035-885">Platform</span></span></th>
    <th><span data-ttu-id="16035-886">扩展点</span><span class="sxs-lookup"><span data-stu-id="16035-886">Extension points</span></span></th>
    <th><span data-ttu-id="16035-887">API 要求集</span><span class="sxs-lookup"><span data-stu-id="16035-887">API requirement sets</span></span></th>
    <th><span data-ttu-id="16035-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="16035-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-889">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="16035-889">Office 2019 on Windows</span></span><br><span data-ttu-id="16035-890">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-890">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-891">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-891">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-893">- Selection</span><span class="sxs-lookup"><span data-stu-id="16035-893">- Selection</span></span><br><span data-ttu-id="16035-894">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-894">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-895">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="16035-895">Office 2016 on Windows</span></span><br><span data-ttu-id="16035-896">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-896">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-897">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-897">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-899">- Selection</span><span class="sxs-lookup"><span data-stu-id="16035-899">- Selection</span></span><br><span data-ttu-id="16035-900">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-900">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16035-901">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="16035-901">Office 2013 on Windows</span></span><br><span data-ttu-id="16035-902">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="16035-902">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="16035-903">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="16035-903">- TaskPane</span></span></td>
    <td> <span data-ttu-id="16035-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="16035-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="16035-905">- Selection</span><span class="sxs-lookup"><span data-stu-id="16035-905">- Selection</span></span><br><span data-ttu-id="16035-906">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16035-906">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="16035-907">另请参阅</span><span class="sxs-lookup"><span data-stu-id="16035-907">See also</span></span>

- [<span data-ttu-id="16035-908">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="16035-908">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="16035-909">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="16035-909">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="16035-910">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="16035-910">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="16035-911">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="16035-911">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="16035-912">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="16035-912">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="16035-913">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="16035-913">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="16035-914">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="16035-914">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="16035-915">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="16035-915">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="16035-916">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="16035-916">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="16035-917">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="16035-917">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="16035-918">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="16035-918">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
