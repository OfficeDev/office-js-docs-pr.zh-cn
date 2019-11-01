---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 10/30/2019
localization_priority: Priority
ms.openlocfilehash: 3621236ea86410d70d17655450e1f6d32a212823
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901946"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4ed14-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="4ed14-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4ed14-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="4ed14-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="4ed14-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="4ed14-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="4ed14-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="4ed14-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="4ed14-108">Excel</span><span class="sxs-lookup"><span data-stu-id="4ed14-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4ed14-109">平台</span><span class="sxs-lookup"><span data-stu-id="4ed14-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4ed14-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="4ed14-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4ed14-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4ed14-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4ed14-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="4ed14-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="4ed14-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-114">- TaskPane</span></span><br><span data-ttu-id="4ed14-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-115">
        - Content</span></span><br><span data-ttu-id="4ed14-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="4ed14-116">
        - Custom Functions</span></span><br><span data-ttu-id="4ed14-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="4ed14-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4ed14-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ed14-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ed14-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ed14-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ed14-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ed14-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ed14-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4ed14-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4ed14-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4ed14-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4ed14-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-128">
        - BindingEvents</span></span><br><span data-ttu-id="4ed14-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-129">
        - CompressedFile</span></span><br><span data-ttu-id="4ed14-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-130">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-131">
        - File</span></span><br><span data-ttu-id="4ed14-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-132">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-134">
        - Selection</span></span><br><span data-ttu-id="4ed14-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-135">
        - Settings</span></span><br><span data-ttu-id="4ed14-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-136">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-137">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-138">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-140">Office on Windows</span></span><br><span data-ttu-id="4ed14-141">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-142">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-142">- TaskPane</span></span><br><span data-ttu-id="4ed14-143">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-143">
        - Content</span></span><br><span data-ttu-id="4ed14-144">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="4ed14-144">
        - Custom Functions</span></span><br><span data-ttu-id="4ed14-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="4ed14-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4ed14-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ed14-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ed14-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ed14-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ed14-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ed14-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ed14-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4ed14-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4ed14-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4ed14-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4ed14-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="4ed14-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-158">
        - BindingEvents</span></span><br><span data-ttu-id="4ed14-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-159">
        - CompressedFile</span></span><br><span data-ttu-id="4ed14-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-160">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-161">
        - File</span></span><br><span data-ttu-id="4ed14-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-162">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-164">
        - Selection</span></span><br><span data-ttu-id="4ed14-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-165">
        - Settings</span></span><br><span data-ttu-id="4ed14-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-166">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-167">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-168">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-170">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-170">Office 2019 on Windows</span></span><br><span data-ttu-id="4ed14-171">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4ed14-172">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-172">- TaskPane</span></span><br><span data-ttu-id="4ed14-173">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-173">
        - Content</span></span><br><span data-ttu-id="4ed14-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4ed14-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ed14-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ed14-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ed14-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ed14-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ed14-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ed14-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4ed14-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4ed14-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4ed14-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-185">- BindingEvents</span></span><br><span data-ttu-id="4ed14-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-186">
        - CompressedFile</span></span><br><span data-ttu-id="4ed14-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-187">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-188">
        - File</span></span><br><span data-ttu-id="4ed14-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-189">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-191">
        - Selection</span></span><br><span data-ttu-id="4ed14-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-192">
        - Settings</span></span><br><span data-ttu-id="4ed14-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-193">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-194">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-195">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-197">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-197">Office 2016 on Windows</span></span><br><span data-ttu-id="4ed14-198">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4ed14-199">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-199">- TaskPane</span></span><br><span data-ttu-id="4ed14-200">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-200">
        - Content</span></span></td>
    <td><span data-ttu-id="4ed14-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ed14-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4ed14-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4ed14-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4ed14-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-204">- BindingEvents</span></span><br><span data-ttu-id="4ed14-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-205">
        - CompressedFile</span></span><br><span data-ttu-id="4ed14-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-206">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-207">
        - File</span></span><br><span data-ttu-id="4ed14-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-208">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-210">
        - Selection</span></span><br><span data-ttu-id="4ed14-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-211">
        - Settings</span></span><br><span data-ttu-id="4ed14-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-212">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-213">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-214">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-216">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4ed14-216">Office 2013 on Windows</span></span><br><span data-ttu-id="4ed14-217">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4ed14-218">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-218">
        - TaskPane</span></span><br><span data-ttu-id="4ed14-219">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4ed14-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4ed14-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4ed14-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4ed14-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-222">
        - BindingEvents</span></span><br><span data-ttu-id="4ed14-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-223">
        - CompressedFile</span></span><br><span data-ttu-id="4ed14-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-224">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-225">
        - File</span></span><br><span data-ttu-id="4ed14-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-226">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-228">
        - Selection</span></span><br><span data-ttu-id="4ed14-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-229">
        - Settings</span></span><br><span data-ttu-id="4ed14-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-230">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-231">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-232">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-234">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-234">Office on iPad</span></span><br><span data-ttu-id="4ed14-235">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="4ed14-236">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-236">- TaskPane</span></span><br><span data-ttu-id="4ed14-237">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-237">
        - Content</span></span></td>
    <td><span data-ttu-id="4ed14-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ed14-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ed14-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ed14-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ed14-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ed14-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ed14-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4ed14-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4ed14-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4ed14-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4ed14-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-249">- BindingEvents</span></span><br><span data-ttu-id="4ed14-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-250">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-251">
        - File</span></span><br><span data-ttu-id="4ed14-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-252">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-254">
        - Selection</span></span><br><span data-ttu-id="4ed14-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-255">
        - Settings</span></span><br><span data-ttu-id="4ed14-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-256">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-257">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-258">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-260">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-260">Office on Mac</span></span><br><span data-ttu-id="4ed14-261">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="4ed14-262">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-262">- TaskPane</span></span><br><span data-ttu-id="4ed14-263">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-263">
        - Content</span></span><br><span data-ttu-id="4ed14-264">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="4ed14-264">
        - Custom Functions</span></span><br><span data-ttu-id="4ed14-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4ed14-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ed14-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ed14-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ed14-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ed14-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ed14-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ed14-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4ed14-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4ed14-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4ed14-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4ed14-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="4ed14-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-278">- BindingEvents</span></span><br><span data-ttu-id="4ed14-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-279">
        - CompressedFile</span></span><br><span data-ttu-id="4ed14-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-280">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-281">
        - File</span></span><br><span data-ttu-id="4ed14-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-282">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-284">
        - PdfFile</span></span><br><span data-ttu-id="4ed14-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-285">
        - Selection</span></span><br><span data-ttu-id="4ed14-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-286">
        - Settings</span></span><br><span data-ttu-id="4ed14-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-287">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-288">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-289">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-291">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-291">Office 2019 on Mac</span></span><br><span data-ttu-id="4ed14-292">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4ed14-293">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-293">- TaskPane</span></span><br><span data-ttu-id="4ed14-294">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-294">
        - Content</span></span><br><span data-ttu-id="4ed14-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4ed14-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ed14-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ed14-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ed14-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ed14-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ed14-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ed14-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4ed14-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4ed14-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4ed14-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-306">- BindingEvents</span></span><br><span data-ttu-id="4ed14-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-307">
        - CompressedFile</span></span><br><span data-ttu-id="4ed14-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-308">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-309">
        - File</span></span><br><span data-ttu-id="4ed14-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-310">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-312">
        - PdfFile</span></span><br><span data-ttu-id="4ed14-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-313">
        - Selection</span></span><br><span data-ttu-id="4ed14-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-314">
        - Settings</span></span><br><span data-ttu-id="4ed14-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-315">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-316">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-317">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-319">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-319">Office 2016 on Mac</span></span><br><span data-ttu-id="4ed14-320">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4ed14-321">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-321">- TaskPane</span></span><br><span data-ttu-id="4ed14-322">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-322">
        - Content</span></span></td>
    <td><span data-ttu-id="4ed14-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ed14-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4ed14-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4ed14-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4ed14-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-326">- BindingEvents</span></span><br><span data-ttu-id="4ed14-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-327">
        - CompressedFile</span></span><br><span data-ttu-id="4ed14-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-328">
        - DocumentEvents</span></span><br><span data-ttu-id="4ed14-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-329">
        - File</span></span><br><span data-ttu-id="4ed14-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-330">
        - MatrixBindings</span></span><br><span data-ttu-id="4ed14-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-332">
        - PdfFile</span></span><br><span data-ttu-id="4ed14-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-333">
        - Selection</span></span><br><span data-ttu-id="4ed14-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-334">
        - Settings</span></span><br><span data-ttu-id="4ed14-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-335">
        - TableBindings</span></span><br><span data-ttu-id="4ed14-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-336">
        - TableCoercion</span></span><br><span data-ttu-id="4ed14-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-337">
        - TextBindings</span></span><br><span data-ttu-id="4ed14-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4ed14-339">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="4ed14-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="4ed14-340">自定义函数</span><span class="sxs-lookup"><span data-stu-id="4ed14-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4ed14-341">平台</span><span class="sxs-lookup"><span data-stu-id="4ed14-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4ed14-342">扩展点</span><span class="sxs-lookup"><span data-stu-id="4ed14-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4ed14-343">API 要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4ed14-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4ed14-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-345">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="4ed14-345">Office on the web</span></span></td>
    <td><span data-ttu-id="4ed14-346">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="4ed14-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="4ed14-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-348">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-348">Office on Windows</span></span><br><span data-ttu-id="4ed14-349">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="4ed14-350">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="4ed14-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="4ed14-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-352">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="4ed14-352">Office for Mac</span></span><br><span data-ttu-id="4ed14-353">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="4ed14-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="4ed14-354">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="4ed14-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="4ed14-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="4ed14-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="4ed14-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ed14-357">平台</span><span class="sxs-lookup"><span data-stu-id="4ed14-357">Platform</span></span></th>
    <th><span data-ttu-id="4ed14-358">扩展点</span><span class="sxs-lookup"><span data-stu-id="4ed14-358">Extension points</span></span></th>
    <th><span data-ttu-id="4ed14-359">API 要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ed14-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4ed14-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-361">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="4ed14-361">Office on the web</span></span><br><span data-ttu-id="4ed14-362">（新式）</span><span class="sxs-lookup"><span data-stu-id="4ed14-362">(modern)</span></span></td>
    <td> <span data-ttu-id="4ed14-363">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-363">- Mail Read</span></span><br><span data-ttu-id="4ed14-364">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-364">
      - Mail Compose</span></span><br><span data-ttu-id="4ed14-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ed14-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4ed14-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4ed14-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="4ed14-374">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-375">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="4ed14-375">Office on the web</span></span><br><span data-ttu-id="4ed14-376">（经典）</span><span class="sxs-lookup"><span data-stu-id="4ed14-376">(classic)</span></span></td>
    <td> <span data-ttu-id="4ed14-377">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-377">- Mail Read</span></span><br><span data-ttu-id="4ed14-378">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-378">
      - Mail Compose</span></span><br><span data-ttu-id="4ed14-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ed14-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4ed14-386">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-387">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-387">Office on Windows</span></span><br><span data-ttu-id="4ed14-388">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-389">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-389">- Mail Read</span></span><br><span data-ttu-id="4ed14-390">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-390">
      - Mail Compose</span></span><br><span data-ttu-id="4ed14-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4ed14-392">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="4ed14-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4ed14-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ed14-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4ed14-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4ed14-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="4ed14-401">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-401">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-402">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-402">Office 2019 on Windows</span></span><br><span data-ttu-id="4ed14-403">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-403">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-404">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-404">- Mail Read</span></span><br><span data-ttu-id="4ed14-405">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-405">
      - Mail Compose</span></span><br><span data-ttu-id="4ed14-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4ed14-407">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="4ed14-407">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4ed14-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ed14-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4ed14-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4ed14-415">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-416">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-416">Office 2016 on Windows</span></span><br><span data-ttu-id="4ed14-417">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-418">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-418">- Mail Read</span></span><br><span data-ttu-id="4ed14-419">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-419">
      - Mail Compose</span></span><br><span data-ttu-id="4ed14-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4ed14-421">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="4ed14-421">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4ed14-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="4ed14-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="4ed14-426">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-427">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4ed14-427">Office 2013 on Windows</span></span><br><span data-ttu-id="4ed14-428">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-428">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-429">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-429">- Mail Read</span></span><br><span data-ttu-id="4ed14-430">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-430">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="4ed14-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="4ed14-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="4ed14-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="4ed14-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="4ed14-435">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-435">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-436">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-436">Office on iOS</span></span><br><span data-ttu-id="4ed14-437">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-437">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-438">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-438">- Mail Read</span></span><br><span data-ttu-id="4ed14-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4ed14-445">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-446">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-446">Office on Mac</span></span><br><span data-ttu-id="4ed14-447">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-447">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-448">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-448">- Mail Read</span></span><br><span data-ttu-id="4ed14-449">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-449">
      - Mail Compose</span></span><br><span data-ttu-id="4ed14-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ed14-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4ed14-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4ed14-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="4ed14-459">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-460">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-460">Office 2019 on Mac</span></span><br><span data-ttu-id="4ed14-461">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-462">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-462">- Mail Read</span></span><br><span data-ttu-id="4ed14-463">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-463">
      - Mail Compose</span></span><br><span data-ttu-id="4ed14-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ed14-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4ed14-471">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-472">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-472">Office 2016 on Mac</span></span><br><span data-ttu-id="4ed14-473">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-474">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-474">- Mail Read</span></span><br><span data-ttu-id="4ed14-475">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="4ed14-475">
      - Mail Compose</span></span><br><span data-ttu-id="4ed14-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ed14-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4ed14-483">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-484">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-484">Office on Android</span></span><br><span data-ttu-id="4ed14-485">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-486">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="4ed14-486">- Mail Read</span></span><br><span data-ttu-id="4ed14-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ed14-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ed14-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ed14-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ed14-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4ed14-493">不可用</span><span class="sxs-lookup"><span data-stu-id="4ed14-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="4ed14-494">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="4ed14-494">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4ed14-495">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="4ed14-495">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="4ed14-496">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="4ed14-496">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="4ed14-497">Word</span><span class="sxs-lookup"><span data-stu-id="4ed14-497">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ed14-498">平台</span><span class="sxs-lookup"><span data-stu-id="4ed14-498">Platform</span></span></th>
    <th><span data-ttu-id="4ed14-499">扩展点</span><span class="sxs-lookup"><span data-stu-id="4ed14-499">Extension points</span></span></th>
    <th><span data-ttu-id="4ed14-500">API 要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-500">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ed14-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4ed14-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-502">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="4ed14-502">Office on the web</span></span></td>
    <td> <span data-ttu-id="4ed14-503">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-503">- TaskPane</span></span><br><span data-ttu-id="4ed14-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4ed14-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4ed14-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4ed14-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4ed14-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4ed14-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-511">- BindingEvents</span></span><br><span data-ttu-id="4ed14-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-513">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-514">
         - File</span></span><br><span data-ttu-id="4ed14-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-516">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-519">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-520">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-520">
         - Selection</span></span><br><span data-ttu-id="4ed14-521">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-521">
         - Settings</span></span><br><span data-ttu-id="4ed14-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-522">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-523">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-524">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-525">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-526">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-527">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-527">Office on Windows</span></span><br><span data-ttu-id="4ed14-528">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-528">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-529">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-529">- TaskPane</span></span><br><span data-ttu-id="4ed14-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4ed14-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4ed14-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4ed14-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4ed14-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4ed14-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-537">- BindingEvents</span></span><br><span data-ttu-id="4ed14-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-538">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-540">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-541">
         - File</span></span><br><span data-ttu-id="4ed14-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-543">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-546">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-547">
         - Selection</span></span><br><span data-ttu-id="4ed14-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-548">
         - Settings</span></span><br><span data-ttu-id="4ed14-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-549">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-550">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-551">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-552">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-553">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-554">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-554">Office 2019 on Windows</span></span><br><span data-ttu-id="4ed14-555">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-555">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-556">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-556">- TaskPane</span></span><br><span data-ttu-id="4ed14-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4ed14-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4ed14-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4ed14-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-563">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-563">- BindingEvents</span></span><br><span data-ttu-id="4ed14-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-564">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-565">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-565">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-566">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-567">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-567">
         - File</span></span><br><span data-ttu-id="4ed14-568">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-568">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-569">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-569">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-570">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-570">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-571">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-571">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-572">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-573">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-573">
         - Selection</span></span><br><span data-ttu-id="4ed14-574">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-574">
         - Settings</span></span><br><span data-ttu-id="4ed14-575">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-575">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-576">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-576">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-577">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-577">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-578">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-578">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-579">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-579">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-580">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-580">Office 2016 on Windows</span></span><br><span data-ttu-id="4ed14-581">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-581">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-582">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-582">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4ed14-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4ed14-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4ed14-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-586">- BindingEvents</span></span><br><span data-ttu-id="4ed14-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-587">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-589">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-590">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-590">
         - File</span></span><br><span data-ttu-id="4ed14-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-592">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-595">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-596">
         - Selection</span></span><br><span data-ttu-id="4ed14-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-597">
         - Settings</span></span><br><span data-ttu-id="4ed14-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-598">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-599">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-600">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-601">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-603">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4ed14-603">Office 2013 on Windows</span></span><br><span data-ttu-id="4ed14-604">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-605">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4ed14-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4ed14-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-608">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-608">- BindingEvents</span></span><br><span data-ttu-id="4ed14-609">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-609">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-610">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-610">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-611">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-611">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-612">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-612">
         - File</span></span><br><span data-ttu-id="4ed14-613">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-613">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-614">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-614">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-615">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-615">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-616">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-616">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-617">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-617">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-618">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-618">
         - Selection</span></span><br><span data-ttu-id="4ed14-619">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-619">
         - Settings</span></span><br><span data-ttu-id="4ed14-620">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-620">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-621">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-621">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-622">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-622">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-623">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-624">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-624">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-625">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-625">Office on iPad</span></span><br><span data-ttu-id="4ed14-626">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-626">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-627">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-627">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4ed14-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4ed14-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4ed14-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="4ed14-633">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-633">- BindingEvents</span></span><br><span data-ttu-id="4ed14-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-634">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-635">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-635">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-636">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-636">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-637">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-637">
         - File</span></span><br><span data-ttu-id="4ed14-638">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-638">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-639">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-639">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-640">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-640">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-641">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-641">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-642">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-642">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-643">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-643">
         - Selection</span></span><br><span data-ttu-id="4ed14-644">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-644">
         - Settings</span></span><br><span data-ttu-id="4ed14-645">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-645">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-646">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-646">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-647">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-647">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-648">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-648">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-649">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-649">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-650">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-650">Office on Mac</span></span><br><span data-ttu-id="4ed14-651">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-651">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-652">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-652">- TaskPane</span></span><br><span data-ttu-id="4ed14-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4ed14-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4ed14-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4ed14-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4ed14-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="4ed14-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-660">- BindingEvents</span></span><br><span data-ttu-id="4ed14-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-661">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-663">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-664">
         - File</span></span><br><span data-ttu-id="4ed14-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-666">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-669">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-670">
         - Selection</span></span><br><span data-ttu-id="4ed14-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-671">
         - Settings</span></span><br><span data-ttu-id="4ed14-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-672">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-673">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-674">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-675">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-677">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-677">Office 2019 on Mac</span></span><br><span data-ttu-id="4ed14-678">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-678">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-679">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-679">- TaskPane</span></span><br><span data-ttu-id="4ed14-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4ed14-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4ed14-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4ed14-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="4ed14-686">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-686">- BindingEvents</span></span><br><span data-ttu-id="4ed14-687">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-687">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-688">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-688">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-689">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-689">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-690">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-690">
         - File</span></span><br><span data-ttu-id="4ed14-691">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-691">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-692">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-692">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-693">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-693">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-694">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-694">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-695">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-695">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-696">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-696">
         - Selection</span></span><br><span data-ttu-id="4ed14-697">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-697">
         - Settings</span></span><br><span data-ttu-id="4ed14-698">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-698">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-699">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-699">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-700">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-700">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-701">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-702">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-702">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-703">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-703">Office 2016 on Mac</span></span><br><span data-ttu-id="4ed14-704">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-704">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-705">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-705">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4ed14-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4ed14-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4ed14-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-709">- BindingEvents</span></span><br><span data-ttu-id="4ed14-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-710">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ed14-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="4ed14-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-712">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-713">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-713">
         - File</span></span><br><span data-ttu-id="4ed14-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-715">
         - MatrixBindings</span></span><br><span data-ttu-id="4ed14-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="4ed14-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4ed14-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-718">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-719">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-719">
         - Selection</span></span><br><span data-ttu-id="4ed14-720">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-720">
         - Settings</span></span><br><span data-ttu-id="4ed14-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-721">
         - TableBindings</span></span><br><span data-ttu-id="4ed14-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-722">
         - TableCoercion</span></span><br><span data-ttu-id="4ed14-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ed14-723">
         - TextBindings</span></span><br><span data-ttu-id="4ed14-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-724">
         - TextCoercion</span></span><br><span data-ttu-id="4ed14-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-725">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="4ed14-726">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="4ed14-726">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4ed14-727">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4ed14-727">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ed14-728">平台</span><span class="sxs-lookup"><span data-stu-id="4ed14-728">Platform</span></span></th>
    <th><span data-ttu-id="4ed14-729">扩展点</span><span class="sxs-lookup"><span data-stu-id="4ed14-729">Extension points</span></span></th>
    <th><span data-ttu-id="4ed14-730">API 要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-730">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ed14-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4ed14-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-732">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="4ed14-732">Office on the web</span></span></td>
    <td> <span data-ttu-id="4ed14-733">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-733">- Content</span></span><br><span data-ttu-id="4ed14-734">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-734">
         - TaskPane</span></span><br><span data-ttu-id="4ed14-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4ed14-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4ed14-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4ed14-740">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-740">- ActiveView</span></span><br><span data-ttu-id="4ed14-741">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-741">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-742">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-742">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-743">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-743">
         - File</span></span><br><span data-ttu-id="4ed14-744">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-744">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-745">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-745">
         - Selection</span></span><br><span data-ttu-id="4ed14-746">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-746">
         - Settings</span></span><br><span data-ttu-id="4ed14-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-747">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-748">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-748">Office on Windows</span></span><br><span data-ttu-id="4ed14-749">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-749">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-750">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-750">- Content</span></span><br><span data-ttu-id="4ed14-751">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-751">
         - TaskPane</span></span><br><span data-ttu-id="4ed14-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4ed14-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4ed14-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4ed14-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-757">- ActiveView</span></span><br><span data-ttu-id="4ed14-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-758">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-759">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-760">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-760">
         - File</span></span><br><span data-ttu-id="4ed14-761">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-761">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-762">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-762">
         - Selection</span></span><br><span data-ttu-id="4ed14-763">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-763">
         - Settings</span></span><br><span data-ttu-id="4ed14-764">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-764">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-765">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-765">Office 2019 on Windows</span></span><br><span data-ttu-id="4ed14-766">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-766">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-767">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-767">- Content</span></span><br><span data-ttu-id="4ed14-768">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-768">
         - TaskPane</span></span><br><span data-ttu-id="4ed14-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-772">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-772">- ActiveView</span></span><br><span data-ttu-id="4ed14-773">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-773">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-774">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-774">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-775">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-775">
         - File</span></span><br><span data-ttu-id="4ed14-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-776">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-777">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-777">
         - Selection</span></span><br><span data-ttu-id="4ed14-778">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-778">
         - Settings</span></span><br><span data-ttu-id="4ed14-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-780">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-780">Office 2016 on Windows</span></span><br><span data-ttu-id="4ed14-781">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-782">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-782">- Content</span></span><br><span data-ttu-id="4ed14-783">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-783">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4ed14-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4ed14-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-786">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-786">- ActiveView</span></span><br><span data-ttu-id="4ed14-787">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-787">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-788">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-788">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-789">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-789">
         - File</span></span><br><span data-ttu-id="4ed14-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-790">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-791">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-791">
         - Selection</span></span><br><span data-ttu-id="4ed14-792">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-792">
         - Settings</span></span><br><span data-ttu-id="4ed14-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-794">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4ed14-794">Office 2013 on Windows</span></span><br><span data-ttu-id="4ed14-795">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-795">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-796">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-796">- Content</span></span><br><span data-ttu-id="4ed14-797">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-797">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="4ed14-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4ed14-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4ed14-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-800">- ActiveView</span></span><br><span data-ttu-id="4ed14-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-801">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-802">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-803">
         - File</span></span><br><span data-ttu-id="4ed14-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-804">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-805">
         - Selection</span></span><br><span data-ttu-id="4ed14-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-806">
         - Settings</span></span><br><span data-ttu-id="4ed14-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-808">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-808">Office on iPad</span></span><br><span data-ttu-id="4ed14-809">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-810">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-810">- Content</span></span><br><span data-ttu-id="4ed14-811">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4ed14-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-815">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-815">- ActiveView</span></span><br><span data-ttu-id="4ed14-816">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-816">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-817">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-817">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-818">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-818">
         - File</span></span><br><span data-ttu-id="4ed14-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-819">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-820">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-820">
         - Selection</span></span><br><span data-ttu-id="4ed14-821">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-821">
         - Settings</span></span><br><span data-ttu-id="4ed14-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-823">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="4ed14-823">Office on Mac</span></span><br><span data-ttu-id="4ed14-824">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4ed14-824">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4ed14-825">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-825">- Content</span></span><br><span data-ttu-id="4ed14-826">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-826">
         - TaskPane</span></span><br><span data-ttu-id="4ed14-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4ed14-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4ed14-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4ed14-832">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-832">- ActiveView</span></span><br><span data-ttu-id="4ed14-833">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-833">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-834">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-834">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-835">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-835">
         - File</span></span><br><span data-ttu-id="4ed14-836">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-836">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-837">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-837">
         - Selection</span></span><br><span data-ttu-id="4ed14-838">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-838">
         - Settings</span></span><br><span data-ttu-id="4ed14-839">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-839">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-840">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-840">Office 2019 on Mac</span></span><br><span data-ttu-id="4ed14-841">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-841">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-842">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-842">- Content</span></span><br><span data-ttu-id="4ed14-843">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-843">
         - TaskPane</span></span><br><span data-ttu-id="4ed14-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-847">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-847">- ActiveView</span></span><br><span data-ttu-id="4ed14-848">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-848">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-849">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-849">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-850">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-850">
         - File</span></span><br><span data-ttu-id="4ed14-851">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-851">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-852">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-852">
         - Selection</span></span><br><span data-ttu-id="4ed14-853">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-853">
         - Settings</span></span><br><span data-ttu-id="4ed14-854">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-854">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-855">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-855">Office 2016 on Mac</span></span><br><span data-ttu-id="4ed14-856">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-856">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-857">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-857">- Content</span></span><br><span data-ttu-id="4ed14-858">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-858">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4ed14-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4ed14-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-861">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ed14-861">- ActiveView</span></span><br><span data-ttu-id="4ed14-862">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-862">
         - CompressedFile</span></span><br><span data-ttu-id="4ed14-863">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-863">
         - DocumentEvents</span></span><br><span data-ttu-id="4ed14-864">
         - File</span><span class="sxs-lookup"><span data-stu-id="4ed14-864">
         - File</span></span><br><span data-ttu-id="4ed14-865">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ed14-865">
         - PdfFile</span></span><br><span data-ttu-id="4ed14-866">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-866">
         - Selection</span></span><br><span data-ttu-id="4ed14-867">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-867">
         - Settings</span></span><br><span data-ttu-id="4ed14-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4ed14-869">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="4ed14-869">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="4ed14-870">OneNote</span><span class="sxs-lookup"><span data-stu-id="4ed14-870">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ed14-871">平台</span><span class="sxs-lookup"><span data-stu-id="4ed14-871">Platform</span></span></th>
    <th><span data-ttu-id="4ed14-872">扩展点</span><span class="sxs-lookup"><span data-stu-id="4ed14-872">Extension points</span></span></th>
    <th><span data-ttu-id="4ed14-873">API 要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-873">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ed14-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4ed14-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-875">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="4ed14-875">Office on the web</span></span></td>
    <td> <span data-ttu-id="4ed14-876">- 内容</span><span class="sxs-lookup"><span data-stu-id="4ed14-876">- Content</span></span><br><span data-ttu-id="4ed14-877">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-877">
         - TaskPane</span></span><br><span data-ttu-id="4ed14-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ed14-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4ed14-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4ed14-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-882">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ed14-882">- DocumentEvents</span></span><br><span data-ttu-id="4ed14-883">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-883">
         - HtmlCoercion</span></span><br><span data-ttu-id="4ed14-884">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4ed14-884">
         - Settings</span></span><br><span data-ttu-id="4ed14-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-885">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="4ed14-886">项目</span><span class="sxs-lookup"><span data-stu-id="4ed14-886">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ed14-887">平台</span><span class="sxs-lookup"><span data-stu-id="4ed14-887">Platform</span></span></th>
    <th><span data-ttu-id="4ed14-888">扩展点</span><span class="sxs-lookup"><span data-stu-id="4ed14-888">Extension points</span></span></th>
    <th><span data-ttu-id="4ed14-889">API 要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-889">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ed14-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4ed14-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-891">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4ed14-891">Office 2019 on Windows</span></span><br><span data-ttu-id="4ed14-892">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-893">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-895">- Selection</span></span><br><span data-ttu-id="4ed14-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-897">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4ed14-897">Office 2016 on Windows</span></span><br><span data-ttu-id="4ed14-898">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-899">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-901">- Selection</span></span><br><span data-ttu-id="4ed14-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-902">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ed14-903">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4ed14-903">Office 2013 on Windows</span></span><br><span data-ttu-id="4ed14-904">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4ed14-904">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4ed14-905">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="4ed14-905">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4ed14-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ed14-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ed14-907">- Selection</span><span class="sxs-lookup"><span data-stu-id="4ed14-907">- Selection</span></span><br><span data-ttu-id="4ed14-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ed14-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4ed14-909">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4ed14-909">See also</span></span>

- [<span data-ttu-id="4ed14-910">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="4ed14-910">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4ed14-911">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-911">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="4ed14-912">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-912">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="4ed14-913">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="4ed14-913">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="4ed14-914">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="4ed14-914">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="4ed14-915">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="4ed14-915">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="4ed14-916">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="4ed14-916">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="4ed14-917">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="4ed14-917">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="4ed14-918">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="4ed14-918">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="4ed14-919">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="4ed14-919">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="4ed14-920">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="4ed14-920">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
