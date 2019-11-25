---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: ecb906e595c08b973b5146416a5317d59547ed39
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757483"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="aaa28-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="aaa28-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="aaa28-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="aaa28-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="aaa28-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="aaa28-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="aaa28-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="aaa28-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="aaa28-108">Excel</span><span class="sxs-lookup"><span data-stu-id="aaa28-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="aaa28-109">平台</span><span class="sxs-lookup"><span data-stu-id="aaa28-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="aaa28-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="aaa28-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="aaa28-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="aaa28-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="aaa28-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="aaa28-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="aaa28-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-114">- TaskPane</span></span><br><span data-ttu-id="aaa28-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-115">
        - Content</span></span><br><span data-ttu-id="aaa28-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="aaa28-116">
        - Custom Functions</span></span><br><span data-ttu-id="aaa28-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="aaa28-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="aaa28-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aaa28-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aaa28-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aaa28-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aaa28-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aaa28-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aaa28-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="aaa28-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="aaa28-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="aaa28-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="aaa28-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="aaa28-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="aaa28-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-130">
        - BindingEvents</span></span><br><span data-ttu-id="aaa28-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-131">
        - CompressedFile</span></span><br><span data-ttu-id="aaa28-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-132">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-133">
        - File</span></span><br><span data-ttu-id="aaa28-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-134">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-136">
        - Selection</span></span><br><span data-ttu-id="aaa28-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-137">
        - Settings</span></span><br><span data-ttu-id="aaa28-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-138">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-139">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-140">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-142">Office on Windows</span></span><br><span data-ttu-id="aaa28-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-144">- TaskPane</span></span><br><span data-ttu-id="aaa28-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-145">
        - Content</span></span><br><span data-ttu-id="aaa28-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="aaa28-146">
        - Custom Functions</span></span><br><span data-ttu-id="aaa28-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="aaa28-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="aaa28-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aaa28-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aaa28-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aaa28-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aaa28-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aaa28-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aaa28-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="aaa28-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="aaa28-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="aaa28-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="aaa28-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="aaa28-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="aaa28-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-161">
        - BindingEvents</span></span><br><span data-ttu-id="aaa28-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-162">
        - CompressedFile</span></span><br><span data-ttu-id="aaa28-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-163">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-164">
        - File</span></span><br><span data-ttu-id="aaa28-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-165">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-167">
        - Selection</span></span><br><span data-ttu-id="aaa28-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-168">
        - Settings</span></span><br><span data-ttu-id="aaa28-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-169">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-170">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-171">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-173">Office 2019 on Windows</span></span><br><span data-ttu-id="aaa28-174">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="aaa28-175">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-175">- TaskPane</span></span><br><span data-ttu-id="aaa28-176">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-176">
        - Content</span></span><br><span data-ttu-id="aaa28-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="aaa28-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aaa28-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aaa28-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aaa28-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aaa28-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aaa28-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aaa28-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="aaa28-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="aaa28-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="aaa28-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-188">- BindingEvents</span></span><br><span data-ttu-id="aaa28-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-189">
        - CompressedFile</span></span><br><span data-ttu-id="aaa28-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-190">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-191">
        - File</span></span><br><span data-ttu-id="aaa28-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-192">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-194">
        - Selection</span></span><br><span data-ttu-id="aaa28-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-195">
        - Settings</span></span><br><span data-ttu-id="aaa28-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-196">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-197">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-198">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-200">Office 2016 on Windows</span></span><br><span data-ttu-id="aaa28-201">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="aaa28-202">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-202">- TaskPane</span></span><br><span data-ttu-id="aaa28-203">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-203">
        - Content</span></span></td>
    <td><span data-ttu-id="aaa28-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aaa28-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="aaa28-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="aaa28-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="aaa28-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-207">- BindingEvents</span></span><br><span data-ttu-id="aaa28-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-208">
        - CompressedFile</span></span><br><span data-ttu-id="aaa28-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-209">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-210">
        - File</span></span><br><span data-ttu-id="aaa28-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-211">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-213">
        - Selection</span></span><br><span data-ttu-id="aaa28-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-214">
        - Settings</span></span><br><span data-ttu-id="aaa28-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-215">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-216">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-217">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="aaa28-219">Office 2013 on Windows</span></span><br><span data-ttu-id="aaa28-220">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="aaa28-221">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-221">
        - TaskPane</span></span><br><span data-ttu-id="aaa28-222">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="aaa28-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="aaa28-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="aaa28-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="aaa28-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-225">
        - BindingEvents</span></span><br><span data-ttu-id="aaa28-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-226">
        - CompressedFile</span></span><br><span data-ttu-id="aaa28-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-227">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-228">
        - File</span></span><br><span data-ttu-id="aaa28-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-229">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-231">
        - Selection</span></span><br><span data-ttu-id="aaa28-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-232">
        - Settings</span></span><br><span data-ttu-id="aaa28-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-233">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-234">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-235">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-237">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-237">Office on iPad</span></span><br><span data-ttu-id="aaa28-238">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="aaa28-239">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-239">- TaskPane</span></span><br><span data-ttu-id="aaa28-240">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-240">
        - Content</span></span></td>
    <td><span data-ttu-id="aaa28-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aaa28-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aaa28-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aaa28-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aaa28-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aaa28-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aaa28-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="aaa28-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="aaa28-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="aaa28-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="aaa28-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="aaa28-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-253">- BindingEvents</span></span><br><span data-ttu-id="aaa28-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-254">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-255">
        - File</span></span><br><span data-ttu-id="aaa28-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-256">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-258">
        - Selection</span></span><br><span data-ttu-id="aaa28-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-259">
        - Settings</span></span><br><span data-ttu-id="aaa28-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-260">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-261">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-262">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-264">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-264">Office on Mac</span></span><br><span data-ttu-id="aaa28-265">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="aaa28-266">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-266">- TaskPane</span></span><br><span data-ttu-id="aaa28-267">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-267">
        - Content</span></span><br><span data-ttu-id="aaa28-268">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="aaa28-268">
        - Custom Functions</span></span><br><span data-ttu-id="aaa28-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="aaa28-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aaa28-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aaa28-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aaa28-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aaa28-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aaa28-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aaa28-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="aaa28-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="aaa28-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="aaa28-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="aaa28-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="aaa28-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="aaa28-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-283">- BindingEvents</span></span><br><span data-ttu-id="aaa28-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-284">
        - CompressedFile</span></span><br><span data-ttu-id="aaa28-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-285">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-286">
        - File</span></span><br><span data-ttu-id="aaa28-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-287">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-289">
        - PdfFile</span></span><br><span data-ttu-id="aaa28-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-290">
        - Selection</span></span><br><span data-ttu-id="aaa28-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-291">
        - Settings</span></span><br><span data-ttu-id="aaa28-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-292">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-293">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-294">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-296">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-296">Office 2019 on Mac</span></span><br><span data-ttu-id="aaa28-297">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="aaa28-298">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-298">- TaskPane</span></span><br><span data-ttu-id="aaa28-299">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-299">
        - Content</span></span><br><span data-ttu-id="aaa28-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="aaa28-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aaa28-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aaa28-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aaa28-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aaa28-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aaa28-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aaa28-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="aaa28-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="aaa28-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="aaa28-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-311">- BindingEvents</span></span><br><span data-ttu-id="aaa28-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-312">
        - CompressedFile</span></span><br><span data-ttu-id="aaa28-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-313">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-314">
        - File</span></span><br><span data-ttu-id="aaa28-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-315">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-317">
        - PdfFile</span></span><br><span data-ttu-id="aaa28-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-318">
        - Selection</span></span><br><span data-ttu-id="aaa28-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-319">
        - Settings</span></span><br><span data-ttu-id="aaa28-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-320">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-321">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-322">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-324">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-324">Office 2016 on Mac</span></span><br><span data-ttu-id="aaa28-325">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="aaa28-326">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-326">- TaskPane</span></span><br><span data-ttu-id="aaa28-327">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-327">
        - Content</span></span></td>
    <td><span data-ttu-id="aaa28-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aaa28-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="aaa28-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="aaa28-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="aaa28-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-331">- BindingEvents</span></span><br><span data-ttu-id="aaa28-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-332">
        - CompressedFile</span></span><br><span data-ttu-id="aaa28-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-333">
        - DocumentEvents</span></span><br><span data-ttu-id="aaa28-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-334">
        - File</span></span><br><span data-ttu-id="aaa28-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-335">
        - MatrixBindings</span></span><br><span data-ttu-id="aaa28-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-337">
        - PdfFile</span></span><br><span data-ttu-id="aaa28-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-338">
        - Selection</span></span><br><span data-ttu-id="aaa28-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-339">
        - Settings</span></span><br><span data-ttu-id="aaa28-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-340">
        - TableBindings</span></span><br><span data-ttu-id="aaa28-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-341">
        - TableCoercion</span></span><br><span data-ttu-id="aaa28-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-342">
        - TextBindings</span></span><br><span data-ttu-id="aaa28-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="aaa28-344">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="aaa28-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="aaa28-345">自定义函数</span><span class="sxs-lookup"><span data-stu-id="aaa28-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="aaa28-346">平台</span><span class="sxs-lookup"><span data-stu-id="aaa28-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="aaa28-347">扩展点</span><span class="sxs-lookup"><span data-stu-id="aaa28-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="aaa28-348">API 要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="aaa28-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="aaa28-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-350">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="aaa28-350">Office on the web</span></span></td>
    <td><span data-ttu-id="aaa28-351">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="aaa28-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="aaa28-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-353">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-353">Office on Windows</span></span><br><span data-ttu-id="aaa28-354">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="aaa28-355">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="aaa28-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="aaa28-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="aaa28-357">Office for Mac</span></span><br><span data-ttu-id="aaa28-358">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="aaa28-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="aaa28-359">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="aaa28-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="aaa28-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="aaa28-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="aaa28-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aaa28-362">平台</span><span class="sxs-lookup"><span data-stu-id="aaa28-362">Platform</span></span></th>
    <th><span data-ttu-id="aaa28-363">扩展点</span><span class="sxs-lookup"><span data-stu-id="aaa28-363">Extension points</span></span></th>
    <th><span data-ttu-id="aaa28-364">API 要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="aaa28-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="aaa28-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-366">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="aaa28-366">Office on the web</span></span><br><span data-ttu-id="aaa28-367">（新式）</span><span class="sxs-lookup"><span data-stu-id="aaa28-367">(modern)</span></span></td>
    <td> <span data-ttu-id="aaa28-368">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-368">- Mail Read</span></span><br><span data-ttu-id="aaa28-369">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-369">
      - Mail Compose</span></span><br><span data-ttu-id="aaa28-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aaa28-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="aaa28-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="aaa28-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="aaa28-379">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-380">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="aaa28-380">Office on the web</span></span><br><span data-ttu-id="aaa28-381">（经典）</span><span class="sxs-lookup"><span data-stu-id="aaa28-381">(classic)</span></span></td>
    <td> <span data-ttu-id="aaa28-382">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-382">- Mail Read</span></span><br><span data-ttu-id="aaa28-383">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-383">
      - Mail Compose</span></span><br><span data-ttu-id="aaa28-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aaa28-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="aaa28-391">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-392">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-392">Office on Windows</span></span><br><span data-ttu-id="aaa28-393">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-394">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-394">- Mail Read</span></span><br><span data-ttu-id="aaa28-395">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-395">
      - Mail Compose</span></span><br><span data-ttu-id="aaa28-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="aaa28-397">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="aaa28-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="aaa28-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aaa28-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="aaa28-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="aaa28-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="aaa28-406">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-407">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-407">Office 2019 on Windows</span></span><br><span data-ttu-id="aaa28-408">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-409">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-409">- Mail Read</span></span><br><span data-ttu-id="aaa28-410">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-410">
      - Mail Compose</span></span><br><span data-ttu-id="aaa28-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="aaa28-412">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="aaa28-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="aaa28-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aaa28-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="aaa28-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="aaa28-420">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-421">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-421">Office 2016 on Windows</span></span><br><span data-ttu-id="aaa28-422">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-423">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-423">- Mail Read</span></span><br><span data-ttu-id="aaa28-424">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-424">
      - Mail Compose</span></span><br><span data-ttu-id="aaa28-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="aaa28-426">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="aaa28-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="aaa28-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="aaa28-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="aaa28-431">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-432">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="aaa28-432">Office 2013 on Windows</span></span><br><span data-ttu-id="aaa28-433">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-434">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-434">- Mail Read</span></span><br><span data-ttu-id="aaa28-435">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="aaa28-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="aaa28-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="aaa28-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="aaa28-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="aaa28-440">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-441">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-441">Office on iOS</span></span><br><span data-ttu-id="aaa28-442">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-443">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-443">- Mail Read</span></span><br><span data-ttu-id="aaa28-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="aaa28-450">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-451">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-451">Office on Mac</span></span><br><span data-ttu-id="aaa28-452">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-453">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-453">- Mail Read</span></span><br><span data-ttu-id="aaa28-454">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-454">
      - Mail Compose</span></span><br><span data-ttu-id="aaa28-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aaa28-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="aaa28-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="aaa28-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="aaa28-464">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-465">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-465">Office 2019 on Mac</span></span><br><span data-ttu-id="aaa28-466">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-467">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-467">- Mail Read</span></span><br><span data-ttu-id="aaa28-468">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-468">
      - Mail Compose</span></span><br><span data-ttu-id="aaa28-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aaa28-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="aaa28-476">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-477">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-477">Office 2016 on Mac</span></span><br><span data-ttu-id="aaa28-478">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-479">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-479">- Mail Read</span></span><br><span data-ttu-id="aaa28-480">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="aaa28-480">
      - Mail Compose</span></span><br><span data-ttu-id="aaa28-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aaa28-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="aaa28-488">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-489">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-489">Office on Android</span></span><br><span data-ttu-id="aaa28-490">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-491">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="aaa28-491">- Mail Read</span></span><br><span data-ttu-id="aaa28-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aaa28-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aaa28-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aaa28-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aaa28-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="aaa28-498">不可用</span><span class="sxs-lookup"><span data-stu-id="aaa28-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="aaa28-499">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="aaa28-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="aaa28-500">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="aaa28-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="aaa28-501">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="aaa28-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="aaa28-502">Word</span><span class="sxs-lookup"><span data-stu-id="aaa28-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aaa28-503">平台</span><span class="sxs-lookup"><span data-stu-id="aaa28-503">Platform</span></span></th>
    <th><span data-ttu-id="aaa28-504">扩展点</span><span class="sxs-lookup"><span data-stu-id="aaa28-504">Extension points</span></span></th>
    <th><span data-ttu-id="aaa28-505">API 要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="aaa28-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="aaa28-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-507">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="aaa28-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="aaa28-508">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-508">- TaskPane</span></span><br><span data-ttu-id="aaa28-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="aaa28-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="aaa28-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="aaa28-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="aaa28-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="aaa28-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-516">- BindingEvents</span></span><br><span data-ttu-id="aaa28-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-518">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-519">
         - File</span></span><br><span data-ttu-id="aaa28-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-521">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-524">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-525">
         - Selection</span></span><br><span data-ttu-id="aaa28-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-526">
         - Settings</span></span><br><span data-ttu-id="aaa28-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-527">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-528">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-529">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-530">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-532">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-532">Office on Windows</span></span><br><span data-ttu-id="aaa28-533">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-534">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-534">- TaskPane</span></span><br><span data-ttu-id="aaa28-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="aaa28-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="aaa28-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="aaa28-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="aaa28-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="aaa28-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-542">- BindingEvents</span></span><br><span data-ttu-id="aaa28-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-543">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-545">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-546">
         - File</span></span><br><span data-ttu-id="aaa28-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-548">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-551">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-552">
         - Selection</span></span><br><span data-ttu-id="aaa28-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-553">
         - Settings</span></span><br><span data-ttu-id="aaa28-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-554">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-555">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-556">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-557">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-559">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-559">Office 2019 on Windows</span></span><br><span data-ttu-id="aaa28-560">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-561">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-561">- TaskPane</span></span><br><span data-ttu-id="aaa28-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="aaa28-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="aaa28-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="aaa28-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-568">- BindingEvents</span></span><br><span data-ttu-id="aaa28-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-569">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-571">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-572">
         - File</span></span><br><span data-ttu-id="aaa28-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-574">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-577">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-578">
         - Selection</span></span><br><span data-ttu-id="aaa28-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-579">
         - Settings</span></span><br><span data-ttu-id="aaa28-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-580">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-581">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-582">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-583">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-585">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-585">Office 2016 on Windows</span></span><br><span data-ttu-id="aaa28-586">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-587">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="aaa28-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="aaa28-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="aaa28-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-591">- BindingEvents</span></span><br><span data-ttu-id="aaa28-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-592">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-594">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-595">
         - File</span></span><br><span data-ttu-id="aaa28-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-597">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-600">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-601">
         - Selection</span></span><br><span data-ttu-id="aaa28-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-602">
         - Settings</span></span><br><span data-ttu-id="aaa28-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-603">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-604">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-605">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-606">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-608">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="aaa28-608">Office 2013 on Windows</span></span><br><span data-ttu-id="aaa28-609">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-610">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="aaa28-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="aaa28-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-613">- BindingEvents</span></span><br><span data-ttu-id="aaa28-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-614">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-616">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-617">
         - File</span></span><br><span data-ttu-id="aaa28-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-619">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-622">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-623">
         - Selection</span></span><br><span data-ttu-id="aaa28-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-624">
         - Settings</span></span><br><span data-ttu-id="aaa28-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-625">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-626">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-627">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-628">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-630">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-630">Office on iPad</span></span><br><span data-ttu-id="aaa28-631">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-632">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="aaa28-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="aaa28-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="aaa28-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="aaa28-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-638">- BindingEvents</span></span><br><span data-ttu-id="aaa28-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-639">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-641">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-642">
         - File</span></span><br><span data-ttu-id="aaa28-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-644">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-647">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-648">
         - Selection</span></span><br><span data-ttu-id="aaa28-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-649">
         - Settings</span></span><br><span data-ttu-id="aaa28-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-650">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-651">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-652">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-653">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-655">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-655">Office on Mac</span></span><br><span data-ttu-id="aaa28-656">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-657">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-657">- TaskPane</span></span><br><span data-ttu-id="aaa28-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="aaa28-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="aaa28-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="aaa28-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="aaa28-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="aaa28-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-665">- BindingEvents</span></span><br><span data-ttu-id="aaa28-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-666">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-668">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-669">
         - File</span></span><br><span data-ttu-id="aaa28-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-671">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-674">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-675">
         - Selection</span></span><br><span data-ttu-id="aaa28-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-676">
         - Settings</span></span><br><span data-ttu-id="aaa28-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-677">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-678">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-679">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-680">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-682">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-682">Office 2019 on Mac</span></span><br><span data-ttu-id="aaa28-683">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-684">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-684">- TaskPane</span></span><br><span data-ttu-id="aaa28-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="aaa28-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="aaa28-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="aaa28-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="aaa28-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-691">- BindingEvents</span></span><br><span data-ttu-id="aaa28-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-692">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-694">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-695">
         - File</span></span><br><span data-ttu-id="aaa28-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-697">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-700">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-701">
         - Selection</span></span><br><span data-ttu-id="aaa28-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-702">
         - Settings</span></span><br><span data-ttu-id="aaa28-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-703">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-704">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-705">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-706">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-708">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-708">Office 2016 on Mac</span></span><br><span data-ttu-id="aaa28-709">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-710">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="aaa28-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="aaa28-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="aaa28-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-714">- BindingEvents</span></span><br><span data-ttu-id="aaa28-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-715">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aaa28-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="aaa28-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-717">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-718">
         - File</span></span><br><span data-ttu-id="aaa28-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-720">
         - MatrixBindings</span></span><br><span data-ttu-id="aaa28-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="aaa28-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="aaa28-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-723">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-724">
         - Selection</span></span><br><span data-ttu-id="aaa28-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-725">
         - Settings</span></span><br><span data-ttu-id="aaa28-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-726">
         - TableBindings</span></span><br><span data-ttu-id="aaa28-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-727">
         - TableCoercion</span></span><br><span data-ttu-id="aaa28-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aaa28-728">
         - TextBindings</span></span><br><span data-ttu-id="aaa28-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-729">
         - TextCoercion</span></span><br><span data-ttu-id="aaa28-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="aaa28-731">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="aaa28-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="aaa28-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="aaa28-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aaa28-733">平台</span><span class="sxs-lookup"><span data-stu-id="aaa28-733">Platform</span></span></th>
    <th><span data-ttu-id="aaa28-734">扩展点</span><span class="sxs-lookup"><span data-stu-id="aaa28-734">Extension points</span></span></th>
    <th><span data-ttu-id="aaa28-735">API 要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="aaa28-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="aaa28-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-737">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="aaa28-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="aaa28-738">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-738">- Content</span></span><br><span data-ttu-id="aaa28-739">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-739">
         - TaskPane</span></span><br><span data-ttu-id="aaa28-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="aaa28-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="aaa28-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="aaa28-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-745">- ActiveView</span></span><br><span data-ttu-id="aaa28-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-746">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-747">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-748">
         - File</span></span><br><span data-ttu-id="aaa28-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-749">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-750">
         - Selection</span></span><br><span data-ttu-id="aaa28-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-751">
         - Settings</span></span><br><span data-ttu-id="aaa28-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-753">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-753">Office on Windows</span></span><br><span data-ttu-id="aaa28-754">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-755">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-755">- Content</span></span><br><span data-ttu-id="aaa28-756">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-756">
         - TaskPane</span></span><br><span data-ttu-id="aaa28-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="aaa28-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="aaa28-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="aaa28-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-762">- ActiveView</span></span><br><span data-ttu-id="aaa28-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-763">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-764">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-765">
         - File</span></span><br><span data-ttu-id="aaa28-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-766">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-767">
         - Selection</span></span><br><span data-ttu-id="aaa28-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-768">
         - Settings</span></span><br><span data-ttu-id="aaa28-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-770">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-770">Office 2019 on Windows</span></span><br><span data-ttu-id="aaa28-771">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-772">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-772">- Content</span></span><br><span data-ttu-id="aaa28-773">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-773">
         - TaskPane</span></span><br><span data-ttu-id="aaa28-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-777">- ActiveView</span></span><br><span data-ttu-id="aaa28-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-778">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-779">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-780">
         - File</span></span><br><span data-ttu-id="aaa28-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-781">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-782">
         - Selection</span></span><br><span data-ttu-id="aaa28-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-783">
         - Settings</span></span><br><span data-ttu-id="aaa28-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-785">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-785">Office 2016 on Windows</span></span><br><span data-ttu-id="aaa28-786">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-787">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-787">- Content</span></span><br><span data-ttu-id="aaa28-788">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="aaa28-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="aaa28-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-791">- ActiveView</span></span><br><span data-ttu-id="aaa28-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-792">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-793">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-794">
         - File</span></span><br><span data-ttu-id="aaa28-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-795">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-796">
         - Selection</span></span><br><span data-ttu-id="aaa28-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-797">
         - Settings</span></span><br><span data-ttu-id="aaa28-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-799">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="aaa28-799">Office 2013 on Windows</span></span><br><span data-ttu-id="aaa28-800">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-801">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-801">- Content</span></span><br><span data-ttu-id="aaa28-802">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="aaa28-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="aaa28-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="aaa28-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-805">- ActiveView</span></span><br><span data-ttu-id="aaa28-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-806">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-807">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-808">
         - File</span></span><br><span data-ttu-id="aaa28-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-809">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-810">
         - Selection</span></span><br><span data-ttu-id="aaa28-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-811">
         - Settings</span></span><br><span data-ttu-id="aaa28-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-813">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-813">Office on iPad</span></span><br><span data-ttu-id="aaa28-814">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-815">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-815">- Content</span></span><br><span data-ttu-id="aaa28-816">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="aaa28-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-820">- ActiveView</span></span><br><span data-ttu-id="aaa28-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-821">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-822">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-823">
         - File</span></span><br><span data-ttu-id="aaa28-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-824">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-825">
         - Selection</span></span><br><span data-ttu-id="aaa28-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-826">
         - Settings</span></span><br><span data-ttu-id="aaa28-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-828">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="aaa28-828">Office on Mac</span></span><br><span data-ttu-id="aaa28-829">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="aaa28-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="aaa28-830">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-830">- Content</span></span><br><span data-ttu-id="aaa28-831">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-831">
         - TaskPane</span></span><br><span data-ttu-id="aaa28-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="aaa28-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="aaa28-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="aaa28-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-837">- ActiveView</span></span><br><span data-ttu-id="aaa28-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-838">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-839">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-840">
         - File</span></span><br><span data-ttu-id="aaa28-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-841">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-842">
         - Selection</span></span><br><span data-ttu-id="aaa28-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-843">
         - Settings</span></span><br><span data-ttu-id="aaa28-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-845">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-845">Office 2019 on Mac</span></span><br><span data-ttu-id="aaa28-846">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-847">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-847">- Content</span></span><br><span data-ttu-id="aaa28-848">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-848">
         - TaskPane</span></span><br><span data-ttu-id="aaa28-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-852">- ActiveView</span></span><br><span data-ttu-id="aaa28-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-853">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-854">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-855">
         - File</span></span><br><span data-ttu-id="aaa28-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-856">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-857">
         - Selection</span></span><br><span data-ttu-id="aaa28-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-858">
         - Settings</span></span><br><span data-ttu-id="aaa28-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-860">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-860">Office 2016 on Mac</span></span><br><span data-ttu-id="aaa28-861">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-862">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-862">- Content</span></span><br><span data-ttu-id="aaa28-863">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="aaa28-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="aaa28-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aaa28-866">- ActiveView</span></span><br><span data-ttu-id="aaa28-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-867">
         - CompressedFile</span></span><br><span data-ttu-id="aaa28-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-868">
         - DocumentEvents</span></span><br><span data-ttu-id="aaa28-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="aaa28-869">
         - File</span></span><br><span data-ttu-id="aaa28-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="aaa28-870">
         - PdfFile</span></span><br><span data-ttu-id="aaa28-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-871">
         - Selection</span></span><br><span data-ttu-id="aaa28-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-872">
         - Settings</span></span><br><span data-ttu-id="aaa28-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="aaa28-874">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="aaa28-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="aaa28-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="aaa28-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aaa28-876">平台</span><span class="sxs-lookup"><span data-stu-id="aaa28-876">Platform</span></span></th>
    <th><span data-ttu-id="aaa28-877">扩展点</span><span class="sxs-lookup"><span data-stu-id="aaa28-877">Extension points</span></span></th>
    <th><span data-ttu-id="aaa28-878">API 要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="aaa28-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="aaa28-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-880">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="aaa28-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="aaa28-881">- 内容</span><span class="sxs-lookup"><span data-stu-id="aaa28-881">- Content</span></span><br><span data-ttu-id="aaa28-882">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-882">
         - TaskPane</span></span><br><span data-ttu-id="aaa28-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aaa28-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="aaa28-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="aaa28-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aaa28-887">- DocumentEvents</span></span><br><span data-ttu-id="aaa28-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="aaa28-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="aaa28-889">
         - Settings</span></span><br><span data-ttu-id="aaa28-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="aaa28-891">项目</span><span class="sxs-lookup"><span data-stu-id="aaa28-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aaa28-892">平台</span><span class="sxs-lookup"><span data-stu-id="aaa28-892">Platform</span></span></th>
    <th><span data-ttu-id="aaa28-893">扩展点</span><span class="sxs-lookup"><span data-stu-id="aaa28-893">Extension points</span></span></th>
    <th><span data-ttu-id="aaa28-894">API 要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="aaa28-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="aaa28-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-896">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="aaa28-896">Office 2019 on Windows</span></span><br><span data-ttu-id="aaa28-897">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-898">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-900">- Selection</span></span><br><span data-ttu-id="aaa28-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-902">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="aaa28-902">Office 2016 on Windows</span></span><br><span data-ttu-id="aaa28-903">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-904">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-906">- Selection</span></span><br><span data-ttu-id="aaa28-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aaa28-908">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="aaa28-908">Office 2013 on Windows</span></span><br><span data-ttu-id="aaa28-909">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="aaa28-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="aaa28-910">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aaa28-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="aaa28-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aaa28-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aaa28-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="aaa28-912">- Selection</span></span><br><span data-ttu-id="aaa28-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aaa28-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="aaa28-914">另请参阅</span><span class="sxs-lookup"><span data-stu-id="aaa28-914">See also</span></span>

- [<span data-ttu-id="aaa28-915">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="aaa28-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="aaa28-916">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="aaa28-917">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="aaa28-918">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="aaa28-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="aaa28-919">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="aaa28-919">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="aaa28-920">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="aaa28-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="aaa28-921">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="aaa28-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="aaa28-922">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="aaa28-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="aaa28-923">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="aaa28-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="aaa28-924">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="aaa28-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="aaa28-925">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="aaa28-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
