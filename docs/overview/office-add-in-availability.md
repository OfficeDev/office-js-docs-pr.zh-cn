---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 8c3c187d8f9b70f40a35e3773a2267dc76decbd0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611980"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c4cdf-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="c4cdf-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c4cdf-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="c4cdf-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="c4cdf-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="c4cdf-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="c4cdf-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="c4cdf-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="c4cdf-108">Excel</span><span class="sxs-lookup"><span data-stu-id="c4cdf-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c4cdf-109">平台</span><span class="sxs-lookup"><span data-stu-id="c4cdf-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c4cdf-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="c4cdf-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c4cdf-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c4cdf-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c4cdf-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="c4cdf-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-114">- TaskPane</span></span><br><span data-ttu-id="c4cdf-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-115">
        - Content</span></span><br><span data-ttu-id="c4cdf-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c4cdf-116">
        - Custom Functions</span></span><br><span data-ttu-id="c4cdf-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="c4cdf-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c4cdf-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c4cdf-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c4cdf-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c4cdf-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c4cdf-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c4cdf-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c4cdf-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c4cdf-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c4cdf-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="c4cdf-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c4cdf-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-131">
        - BindingEvents</span></span><br><span data-ttu-id="c4cdf-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-132">
        - CompressedFile</span></span><br><span data-ttu-id="c4cdf-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-133">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-134">
        - File</span></span><br><span data-ttu-id="c4cdf-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-135">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-137">
        - Selection</span></span><br><span data-ttu-id="c4cdf-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-138">
        - Settings</span></span><br><span data-ttu-id="c4cdf-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-139">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-140">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-141">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-143">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-143">Office on Windows</span></span><br><span data-ttu-id="c4cdf-144">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-145">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-145">- TaskPane</span></span><br><span data-ttu-id="c4cdf-146">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-146">
        - Content</span></span><br><span data-ttu-id="c4cdf-147">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c4cdf-147">
        - Custom Functions</span></span><br><span data-ttu-id="c4cdf-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="c4cdf-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c4cdf-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c4cdf-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c4cdf-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c4cdf-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c4cdf-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c4cdf-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c4cdf-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c4cdf-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c4cdf-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c4cdf-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c4cdf-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-163">
        - BindingEvents</span></span><br><span data-ttu-id="c4cdf-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-164">
        - CompressedFile</span></span><br><span data-ttu-id="c4cdf-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-165">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-166">
        - File</span></span><br><span data-ttu-id="c4cdf-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-167">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-169">
        - Selection</span></span><br><span data-ttu-id="c4cdf-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-170">
        - Settings</span></span><br><span data-ttu-id="c4cdf-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-171">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-172">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-173">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-175">Office 2019 on Windows</span></span><br><span data-ttu-id="c4cdf-176">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c4cdf-177">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-177">- TaskPane</span></span><br><span data-ttu-id="c4cdf-178">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-178">
        - Content</span></span><br><span data-ttu-id="c4cdf-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c4cdf-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c4cdf-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c4cdf-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c4cdf-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c4cdf-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c4cdf-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c4cdf-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-190">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-191">
        - CompressedFile</span></span><br><span data-ttu-id="c4cdf-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-192">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-193">
        - File</span></span><br><span data-ttu-id="c4cdf-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-194">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-196">
        - Selection</span></span><br><span data-ttu-id="c4cdf-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-197">
        - Settings</span></span><br><span data-ttu-id="c4cdf-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-198">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-199">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-200">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-202">Office 2016 on Windows</span></span><br><span data-ttu-id="c4cdf-203">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c4cdf-204">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-204">- TaskPane</span></span><br><span data-ttu-id="c4cdf-205">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-205">
        - Content</span></span></td>
    <td><span data-ttu-id="c4cdf-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c4cdf-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c4cdf-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-209">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-210">
        - CompressedFile</span></span><br><span data-ttu-id="c4cdf-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-211">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-212">
        - File</span></span><br><span data-ttu-id="c4cdf-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-213">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-215">
        - Selection</span></span><br><span data-ttu-id="c4cdf-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-216">
        - Settings</span></span><br><span data-ttu-id="c4cdf-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-217">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-218">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-219">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c4cdf-221">Office 2013 on Windows</span></span><br><span data-ttu-id="c4cdf-222">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c4cdf-223">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-223">
        - TaskPane</span></span><br><span data-ttu-id="c4cdf-224">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c4cdf-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c4cdf-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c4cdf-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-227">
        - BindingEvents</span></span><br><span data-ttu-id="c4cdf-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-228">
        - CompressedFile</span></span><br><span data-ttu-id="c4cdf-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-229">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-230">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-230">
        - File</span></span><br><span data-ttu-id="c4cdf-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-231">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-233">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-233">
        - Selection</span></span><br><span data-ttu-id="c4cdf-234">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-234">
        - Settings</span></span><br><span data-ttu-id="c4cdf-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-235">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-236">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-237">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-239">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-239">Office on iPad</span></span><br><span data-ttu-id="c4cdf-240">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c4cdf-241">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-241">- TaskPane</span></span><br><span data-ttu-id="c4cdf-242">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-242">
        - Content</span></span></td>
    <td><span data-ttu-id="c4cdf-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c4cdf-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c4cdf-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c4cdf-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c4cdf-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c4cdf-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c4cdf-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c4cdf-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c4cdf-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c4cdf-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-256">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-257">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-258">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-258">
        - File</span></span><br><span data-ttu-id="c4cdf-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-259">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-261">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-261">
        - Selection</span></span><br><span data-ttu-id="c4cdf-262">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-262">
        - Settings</span></span><br><span data-ttu-id="c4cdf-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-263">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-264">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-265">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-267">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-267">Office on Mac</span></span><br><span data-ttu-id="c4cdf-268">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c4cdf-269">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-269">- TaskPane</span></span><br><span data-ttu-id="c4cdf-270">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-270">
        - Content</span></span><br><span data-ttu-id="c4cdf-271">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c4cdf-271">
        - Custom Functions</span></span><br><span data-ttu-id="c4cdf-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c4cdf-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c4cdf-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c4cdf-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c4cdf-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c4cdf-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c4cdf-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c4cdf-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c4cdf-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c4cdf-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c4cdf-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c4cdf-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-287">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-288">
        - CompressedFile</span></span><br><span data-ttu-id="c4cdf-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-289">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-290">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-290">
        - File</span></span><br><span data-ttu-id="c4cdf-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-291">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-293">
        - PdfFile</span></span><br><span data-ttu-id="c4cdf-294">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-294">
        - Selection</span></span><br><span data-ttu-id="c4cdf-295">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-295">
        - Settings</span></span><br><span data-ttu-id="c4cdf-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-296">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-297">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-298">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-300">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-300">Office 2019 on Mac</span></span><br><span data-ttu-id="c4cdf-301">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c4cdf-302">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-302">- TaskPane</span></span><br><span data-ttu-id="c4cdf-303">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-303">
        - Content</span></span><br><span data-ttu-id="c4cdf-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c4cdf-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c4cdf-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c4cdf-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c4cdf-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c4cdf-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c4cdf-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c4cdf-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-315">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-316">
        - CompressedFile</span></span><br><span data-ttu-id="c4cdf-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-317">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-318">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-318">
        - File</span></span><br><span data-ttu-id="c4cdf-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-319">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-321">
        - PdfFile</span></span><br><span data-ttu-id="c4cdf-322">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-322">
        - Selection</span></span><br><span data-ttu-id="c4cdf-323">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-323">
        - Settings</span></span><br><span data-ttu-id="c4cdf-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-324">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-325">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-326">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-328">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-328">Office 2016 on Mac</span></span><br><span data-ttu-id="c4cdf-329">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c4cdf-330">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-330">- TaskPane</span></span><br><span data-ttu-id="c4cdf-331">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-331">
        - Content</span></span></td>
    <td><span data-ttu-id="c4cdf-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c4cdf-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c4cdf-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-335">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-336">
        - CompressedFile</span></span><br><span data-ttu-id="c4cdf-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-337">
        - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-338">
        - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-338">
        - File</span></span><br><span data-ttu-id="c4cdf-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-339">
        - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-341">
        - PdfFile</span></span><br><span data-ttu-id="c4cdf-342">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-342">
        - Selection</span></span><br><span data-ttu-id="c4cdf-343">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-343">
        - Settings</span></span><br><span data-ttu-id="c4cdf-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-344">
        - TableBindings</span></span><br><span data-ttu-id="c4cdf-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-345">
        - TableCoercion</span></span><br><span data-ttu-id="c4cdf-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-346">
        - TextBindings</span></span><br><span data-ttu-id="c4cdf-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c4cdf-348">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="c4cdf-349">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c4cdf-350">平台</span><span class="sxs-lookup"><span data-stu-id="c4cdf-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c4cdf-351">扩展点</span><span class="sxs-lookup"><span data-stu-id="c4cdf-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c4cdf-352">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c4cdf-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-354">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c4cdf-354">Office on the web</span></span></td>
    <td><span data-ttu-id="c4cdf-355">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c4cdf-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c4cdf-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-357">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-357">Office on Windows</span></span><br><span data-ttu-id="c4cdf-358">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c4cdf-359">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c4cdf-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c4cdf-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-361">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="c4cdf-361">Office for Mac</span></span><br><span data-ttu-id="c4cdf-362">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="c4cdf-363">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c4cdf-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c4cdf-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="c4cdf-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="c4cdf-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c4cdf-366">平台</span><span class="sxs-lookup"><span data-stu-id="c4cdf-366">Platform</span></span></th>
    <th><span data-ttu-id="c4cdf-367">扩展点</span><span class="sxs-lookup"><span data-stu-id="c4cdf-367">Extension points</span></span></th>
    <th><span data-ttu-id="c4cdf-368">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="c4cdf-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-370">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c4cdf-370">Office on the web</span></span><br><span data-ttu-id="c4cdf-371">（新式）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-371">(modern)</span></span></td>
    <td> <span data-ttu-id="c4cdf-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c4cdf-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c4cdf-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c4cdf-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c4cdf-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c4cdf-385">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-386">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c4cdf-386">Office on the web</span></span><br><span data-ttu-id="c4cdf-387">（经典）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-387">(classic)</span></span></td>
    <td> <span data-ttu-id="c4cdf-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c4cdf-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c4cdf-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c4cdf-399">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-400">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-400">Office on Windows</span></span><br><span data-ttu-id="c4cdf-401">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c4cdf-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c4cdf-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c4cdf-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c4cdf-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c4cdf-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c4cdf-416">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-417">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-417">Office 2019 on Windows</span></span><br><span data-ttu-id="c4cdf-418">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c4cdf-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c4cdf-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c4cdf-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c4cdf-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c4cdf-432">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-433">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-433">Office 2016 on Windows</span></span><br><span data-ttu-id="c4cdf-434">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c4cdf-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c4cdf-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c4cdf-445">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-446">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c4cdf-446">Office 2013 on Windows</span></span><br><span data-ttu-id="c4cdf-447">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="c4cdf-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="c4cdf-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c4cdf-456">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-457">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-457">Office on iOS</span></span><br><span data-ttu-id="c4cdf-458">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c4cdf-466">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-467">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-467">Office on Mac</span></span><br><span data-ttu-id="c4cdf-468">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c4cdf-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c4cdf-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c4cdf-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c4cdf-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c4cdf-482">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-483">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-483">Office 2019 on Mac</span></span><br><span data-ttu-id="c4cdf-484">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c4cdf-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c4cdf-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c4cdf-496">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-497">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-497">Office 2016 on Mac</span></span><br><span data-ttu-id="c4cdf-498">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c4cdf-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c4cdf-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c4cdf-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c4cdf-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c4cdf-510">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-511">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-511">Office on Android</span></span><br><span data-ttu-id="c4cdf-512">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c4cdf-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">约会组织者（撰写）：联机会议</a> （预览）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="c4cdf-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c4cdf-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c4cdf-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c4cdf-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c4cdf-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c4cdf-521">不可用</span><span class="sxs-lookup"><span data-stu-id="c4cdf-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="c4cdf-522">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c4cdf-523">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="c4cdf-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="c4cdf-524">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="c4cdf-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="c4cdf-525">Word</span><span class="sxs-lookup"><span data-stu-id="c4cdf-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c4cdf-526">平台</span><span class="sxs-lookup"><span data-stu-id="c4cdf-526">Platform</span></span></th>
    <th><span data-ttu-id="c4cdf-527">扩展点</span><span class="sxs-lookup"><span data-stu-id="c4cdf-527">Extension points</span></span></th>
    <th><span data-ttu-id="c4cdf-528">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="c4cdf-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-530">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c4cdf-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="c4cdf-531">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-531">- TaskPane</span></span><br><span data-ttu-id="c4cdf-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c4cdf-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-539">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-541">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-542">
         - File</span></span><br><span data-ttu-id="c4cdf-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-544">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-547">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-548">
         - Selection</span></span><br><span data-ttu-id="c4cdf-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-549">
         - Settings</span></span><br><span data-ttu-id="c4cdf-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-550">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-551">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-552">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-553">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-555">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-555">Office on Windows</span></span><br><span data-ttu-id="c4cdf-556">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-557">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-557">- TaskPane</span></span><br><span data-ttu-id="c4cdf-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c4cdf-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-565">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-566">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-568">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-569">
         - File</span></span><br><span data-ttu-id="c4cdf-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-571">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-574">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-575">
         - Selection</span></span><br><span data-ttu-id="c4cdf-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-576">
         - Settings</span></span><br><span data-ttu-id="c4cdf-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-577">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-578">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-579">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-580">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-582">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-582">Office 2019 on Windows</span></span><br><span data-ttu-id="c4cdf-583">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-584">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-584">- TaskPane</span></span><br><span data-ttu-id="c4cdf-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-591">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-592">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-594">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-595">
         - File</span></span><br><span data-ttu-id="c4cdf-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-597">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-600">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-601">
         - Selection</span></span><br><span data-ttu-id="c4cdf-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-602">
         - Settings</span></span><br><span data-ttu-id="c4cdf-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-603">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-604">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-605">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-606">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-608">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-608">Office 2016 on Windows</span></span><br><span data-ttu-id="c4cdf-609">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-610">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c4cdf-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-614">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-615">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-617">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-618">
         - File</span></span><br><span data-ttu-id="c4cdf-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-620">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-623">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-624">
         - Selection</span></span><br><span data-ttu-id="c4cdf-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-625">
         - Settings</span></span><br><span data-ttu-id="c4cdf-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-626">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-627">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-628">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-629">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-631">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c4cdf-631">Office 2013 on Windows</span></span><br><span data-ttu-id="c4cdf-632">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-633">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c4cdf-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-636">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-637">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-639">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-640">
         - File</span></span><br><span data-ttu-id="c4cdf-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-642">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-645">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-646">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-646">
         - Selection</span></span><br><span data-ttu-id="c4cdf-647">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-647">
         - Settings</span></span><br><span data-ttu-id="c4cdf-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-648">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-649">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-650">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-651">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-653">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-653">Office on iPad</span></span><br><span data-ttu-id="c4cdf-654">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-655">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c4cdf-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-661">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-662">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-664">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-665">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-665">
         - File</span></span><br><span data-ttu-id="c4cdf-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-667">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-670">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-671">
         - Selection</span></span><br><span data-ttu-id="c4cdf-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-672">
         - Settings</span></span><br><span data-ttu-id="c4cdf-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-673">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-674">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-675">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-676">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-678">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-678">Office on Mac</span></span><br><span data-ttu-id="c4cdf-679">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-680">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-680">- TaskPane</span></span><br><span data-ttu-id="c4cdf-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c4cdf-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="c4cdf-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-688">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-689">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-691">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-692">
         - File</span></span><br><span data-ttu-id="c4cdf-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-694">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-697">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-698">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-698">
         - Selection</span></span><br><span data-ttu-id="c4cdf-699">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-699">
         - Settings</span></span><br><span data-ttu-id="c4cdf-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-700">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-701">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-702">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-703">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-705">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-705">Office 2019 on Mac</span></span><br><span data-ttu-id="c4cdf-706">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-707">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-707">- TaskPane</span></span><br><span data-ttu-id="c4cdf-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c4cdf-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c4cdf-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c4cdf-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-714">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-715">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-717">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-718">
         - File</span></span><br><span data-ttu-id="c4cdf-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-720">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-723">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-724">
         - Selection</span></span><br><span data-ttu-id="c4cdf-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-725">
         - Settings</span></span><br><span data-ttu-id="c4cdf-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-726">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-727">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-728">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-729">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-731">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-731">Office 2016 on Mac</span></span><br><span data-ttu-id="c4cdf-732">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-733">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c4cdf-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-737">- BindingEvents</span></span><br><span data-ttu-id="c4cdf-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-738">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c4cdf-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="c4cdf-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-740">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-741">
         - File</span></span><br><span data-ttu-id="c4cdf-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-743">
         - MatrixBindings</span></span><br><span data-ttu-id="c4cdf-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="c4cdf-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c4cdf-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-746">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-747">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-747">
         - Selection</span></span><br><span data-ttu-id="c4cdf-748">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-748">
         - Settings</span></span><br><span data-ttu-id="c4cdf-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-749">
         - TableBindings</span></span><br><span data-ttu-id="c4cdf-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-750">
         - TableCoercion</span></span><br><span data-ttu-id="c4cdf-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-751">
         - TextBindings</span></span><br><span data-ttu-id="c4cdf-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-752">
         - TextCoercion</span></span><br><span data-ttu-id="c4cdf-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="c4cdf-754">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c4cdf-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c4cdf-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c4cdf-756">平台</span><span class="sxs-lookup"><span data-stu-id="c4cdf-756">Platform</span></span></th>
    <th><span data-ttu-id="c4cdf-757">扩展点</span><span class="sxs-lookup"><span data-stu-id="c4cdf-757">Extension points</span></span></th>
    <th><span data-ttu-id="c4cdf-758">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="c4cdf-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-760">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c4cdf-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="c4cdf-761">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-761">- Content</span></span><br><span data-ttu-id="c4cdf-762">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-762">
         - TaskPane</span></span><br><span data-ttu-id="c4cdf-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c4cdf-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-768">- ActiveView</span></span><br><span data-ttu-id="c4cdf-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-769">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-770">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-771">
         - File</span></span><br><span data-ttu-id="c4cdf-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-772">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-773">
         - Selection</span></span><br><span data-ttu-id="c4cdf-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-774">
         - Settings</span></span><br><span data-ttu-id="c4cdf-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-776">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-776">Office on Windows</span></span><br><span data-ttu-id="c4cdf-777">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-778">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-778">- Content</span></span><br><span data-ttu-id="c4cdf-779">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-779">
         - TaskPane</span></span><br><span data-ttu-id="c4cdf-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c4cdf-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-785">- ActiveView</span></span><br><span data-ttu-id="c4cdf-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-786">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-787">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-788">
         - File</span></span><br><span data-ttu-id="c4cdf-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-789">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-790">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-790">
         - Selection</span></span><br><span data-ttu-id="c4cdf-791">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-791">
         - Settings</span></span><br><span data-ttu-id="c4cdf-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-793">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-793">Office 2019 on Windows</span></span><br><span data-ttu-id="c4cdf-794">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-795">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-795">- Content</span></span><br><span data-ttu-id="c4cdf-796">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-796">
         - TaskPane</span></span><br><span data-ttu-id="c4cdf-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-800">- ActiveView</span></span><br><span data-ttu-id="c4cdf-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-801">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-802">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-803">
         - File</span></span><br><span data-ttu-id="c4cdf-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-804">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-805">
         - Selection</span></span><br><span data-ttu-id="c4cdf-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-806">
         - Settings</span></span><br><span data-ttu-id="c4cdf-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-808">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-808">Office 2016 on Windows</span></span><br><span data-ttu-id="c4cdf-809">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-810">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-810">- Content</span></span><br><span data-ttu-id="c4cdf-811">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c4cdf-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-814">- ActiveView</span></span><br><span data-ttu-id="c4cdf-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-815">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-816">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-817">
         - File</span></span><br><span data-ttu-id="c4cdf-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-818">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-819">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-819">
         - Selection</span></span><br><span data-ttu-id="c4cdf-820">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-820">
         - Settings</span></span><br><span data-ttu-id="c4cdf-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-822">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c4cdf-822">Office 2013 on Windows</span></span><br><span data-ttu-id="c4cdf-823">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-824">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-824">- Content</span></span><br><span data-ttu-id="c4cdf-825">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="c4cdf-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c4cdf-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-828">- ActiveView</span></span><br><span data-ttu-id="c4cdf-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-829">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-830">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-831">
         - File</span></span><br><span data-ttu-id="c4cdf-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-832">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-833">
         - Selection</span></span><br><span data-ttu-id="c4cdf-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-834">
         - Settings</span></span><br><span data-ttu-id="c4cdf-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-836">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-836">Office on iPad</span></span><br><span data-ttu-id="c4cdf-837">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-838">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-838">- Content</span></span><br><span data-ttu-id="c4cdf-839">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-843">- ActiveView</span></span><br><span data-ttu-id="c4cdf-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-844">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-845">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-846">
         - File</span></span><br><span data-ttu-id="c4cdf-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-847">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-848">
         - Selection</span></span><br><span data-ttu-id="c4cdf-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-849">
         - Settings</span></span><br><span data-ttu-id="c4cdf-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-851">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="c4cdf-851">Office on Mac</span></span><br><span data-ttu-id="c4cdf-852">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c4cdf-853">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-853">- Content</span></span><br><span data-ttu-id="c4cdf-854">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-854">
         - TaskPane</span></span><br><span data-ttu-id="c4cdf-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c4cdf-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-860">- ActiveView</span></span><br><span data-ttu-id="c4cdf-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-861">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-862">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-863">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-863">
         - File</span></span><br><span data-ttu-id="c4cdf-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-864">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-865">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-865">
         - Selection</span></span><br><span data-ttu-id="c4cdf-866">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-866">
         - Settings</span></span><br><span data-ttu-id="c4cdf-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-868">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-868">Office 2019 on Mac</span></span><br><span data-ttu-id="c4cdf-869">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-870">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-870">- Content</span></span><br><span data-ttu-id="c4cdf-871">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-871">
         - TaskPane</span></span><br><span data-ttu-id="c4cdf-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-875">- ActiveView</span></span><br><span data-ttu-id="c4cdf-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-876">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-877">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-878">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-878">
         - File</span></span><br><span data-ttu-id="c4cdf-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-879">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-880">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-880">
         - Selection</span></span><br><span data-ttu-id="c4cdf-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-881">
         - Settings</span></span><br><span data-ttu-id="c4cdf-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-883">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-883">Office 2016 on Mac</span></span><br><span data-ttu-id="c4cdf-884">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-885">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-885">- Content</span></span><br><span data-ttu-id="c4cdf-886">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c4cdf-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c4cdf-889">- ActiveView</span></span><br><span data-ttu-id="c4cdf-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-890">
         - CompressedFile</span></span><br><span data-ttu-id="c4cdf-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-891">
         - DocumentEvents</span></span><br><span data-ttu-id="c4cdf-892">
         - File</span><span class="sxs-lookup"><span data-stu-id="c4cdf-892">
         - File</span></span><br><span data-ttu-id="c4cdf-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c4cdf-893">
         - PdfFile</span></span><br><span data-ttu-id="c4cdf-894">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-894">
         - Selection</span></span><br><span data-ttu-id="c4cdf-895">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-895">
         - Settings</span></span><br><span data-ttu-id="c4cdf-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c4cdf-897">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="c4cdf-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="c4cdf-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="c4cdf-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c4cdf-899">平台</span><span class="sxs-lookup"><span data-stu-id="c4cdf-899">Platform</span></span></th>
    <th><span data-ttu-id="c4cdf-900">扩展点</span><span class="sxs-lookup"><span data-stu-id="c4cdf-900">Extension points</span></span></th>
    <th><span data-ttu-id="c4cdf-901">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="c4cdf-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-903">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c4cdf-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="c4cdf-904">- 内容</span><span class="sxs-lookup"><span data-stu-id="c4cdf-904">- Content</span></span><br><span data-ttu-id="c4cdf-905">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-905">
         - TaskPane</span></span><br><span data-ttu-id="c4cdf-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c4cdf-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c4cdf-910">- DocumentEvents</span></span><br><span data-ttu-id="c4cdf-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="c4cdf-912">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c4cdf-912">
         - Settings</span></span><br><span data-ttu-id="c4cdf-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="c4cdf-914">项目</span><span class="sxs-lookup"><span data-stu-id="c4cdf-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c4cdf-915">平台</span><span class="sxs-lookup"><span data-stu-id="c4cdf-915">Platform</span></span></th>
    <th><span data-ttu-id="c4cdf-916">扩展点</span><span class="sxs-lookup"><span data-stu-id="c4cdf-916">Extension points</span></span></th>
    <th><span data-ttu-id="c4cdf-917">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="c4cdf-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-919">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c4cdf-919">Office 2019 on Windows</span></span><br><span data-ttu-id="c4cdf-920">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-921">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-923">- Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-923">- Selection</span></span><br><span data-ttu-id="c4cdf-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-925">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c4cdf-925">Office 2016 on Windows</span></span><br><span data-ttu-id="c4cdf-926">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-927">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-929">- Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-929">- Selection</span></span><br><span data-ttu-id="c4cdf-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c4cdf-931">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c4cdf-931">Office 2013 on Windows</span></span><br><span data-ttu-id="c4cdf-932">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c4cdf-933">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c4cdf-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c4cdf-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c4cdf-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c4cdf-935">- Selection</span><span class="sxs-lookup"><span data-stu-id="c4cdf-935">- Selection</span></span><br><span data-ttu-id="c4cdf-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c4cdf-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c4cdf-937">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c4cdf-937">See also</span></span>

- [<span data-ttu-id="c4cdf-938">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="c4cdf-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c4cdf-939">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="c4cdf-940">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="c4cdf-941">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="c4cdf-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="c4cdf-942">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="c4cdf-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="c4cdf-943">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="c4cdf-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="c4cdf-944">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="c4cdf-945">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="c4cdf-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="c4cdf-946">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c4cdf-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="c4cdf-947">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c4cdf-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="c4cdf-948">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="c4cdf-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="c4cdf-949">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="c4cdf-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)