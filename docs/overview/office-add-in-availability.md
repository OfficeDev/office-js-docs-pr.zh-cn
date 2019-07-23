---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 07/18/2019
localization_priority: Priority
ms.openlocfilehash: 510f2419d5d364a536f8c96f2057505161f03993
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804644"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="653f9-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="653f9-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="653f9-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="653f9-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="653f9-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="653f9-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="653f9-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="653f9-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="653f9-108">Excel</span><span class="sxs-lookup"><span data-stu-id="653f9-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="653f9-109">平台</span><span class="sxs-lookup"><span data-stu-id="653f9-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="653f9-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="653f9-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="653f9-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="653f9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="653f9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="653f9-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="653f9-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-114">- TaskPane</span></span><br><span data-ttu-id="653f9-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-115">
        - Content</span></span><br><span data-ttu-id="653f9-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="653f9-116">
        - Custom Functions</span></span><br><span data-ttu-id="653f9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="653f9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="653f9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="653f9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="653f9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="653f9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="653f9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="653f9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="653f9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="653f9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="653f9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="653f9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="653f9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="653f9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="653f9-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-130">
        - BindingEvents</span></span><br><span data-ttu-id="653f9-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-131">
        - CompressedFile</span></span><br><span data-ttu-id="653f9-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-132">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-133">
        - File</span></span><br><span data-ttu-id="653f9-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-134">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-136">
        - Selection</span></span><br><span data-ttu-id="653f9-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-137">
        - Settings</span></span><br><span data-ttu-id="653f9-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-138">
        - TableBindings</span></span><br><span data-ttu-id="653f9-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-139">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-140">
        - TextBindings</span></span><br><span data-ttu-id="653f9-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-142">Office on Windows</span></span><br><span data-ttu-id="653f9-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-144">- TaskPane</span></span><br><span data-ttu-id="653f9-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-145">
        - Content</span></span><br><span data-ttu-id="653f9-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="653f9-146">
        - Custom Functions</span></span><br><span data-ttu-id="653f9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="653f9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="653f9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="653f9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="653f9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="653f9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="653f9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="653f9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="653f9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="653f9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="653f9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="653f9-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="653f9-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="653f9-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="653f9-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-160">
        - BindingEvents</span></span><br><span data-ttu-id="653f9-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-161">
        - CompressedFile</span></span><br><span data-ttu-id="653f9-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-162">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-163">
        - File</span></span><br><span data-ttu-id="653f9-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-164">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-166">
        - Selection</span></span><br><span data-ttu-id="653f9-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-167">
        - Settings</span></span><br><span data-ttu-id="653f9-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-168">
        - TableBindings</span></span><br><span data-ttu-id="653f9-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-169">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-170">
        - TextBindings</span></span><br><span data-ttu-id="653f9-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-172">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-172">Office 2019 on Windows</span></span><br><span data-ttu-id="653f9-173">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="653f9-174">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-174">- TaskPane</span></span><br><span data-ttu-id="653f9-175">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-175">
        - Content</span></span><br><span data-ttu-id="653f9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="653f9-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="653f9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="653f9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="653f9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="653f9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="653f9-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="653f9-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="653f9-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="653f9-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="653f9-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="653f9-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-187">- BindingEvents</span></span><br><span data-ttu-id="653f9-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-188">
        - CompressedFile</span></span><br><span data-ttu-id="653f9-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-189">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-190">
        - File</span></span><br><span data-ttu-id="653f9-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-191">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-193">
        - Selection</span></span><br><span data-ttu-id="653f9-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-194">
        - Settings</span></span><br><span data-ttu-id="653f9-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-195">
        - TableBindings</span></span><br><span data-ttu-id="653f9-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-196">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-197">
        - TextBindings</span></span><br><span data-ttu-id="653f9-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-199">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-199">Office 2016 on Windows</span></span><br><span data-ttu-id="653f9-200">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="653f9-201">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-201">- TaskPane</span></span><br><span data-ttu-id="653f9-202">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-202">
        - Content</span></span></td>
    <td><span data-ttu-id="653f9-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="653f9-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="653f9-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="653f9-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="653f9-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-206">- BindingEvents</span></span><br><span data-ttu-id="653f9-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-207">
        - CompressedFile</span></span><br><span data-ttu-id="653f9-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-208">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-209">
        - File</span></span><br><span data-ttu-id="653f9-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-210">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-212">
        - Selection</span></span><br><span data-ttu-id="653f9-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-213">
        - Settings</span></span><br><span data-ttu-id="653f9-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-214">
        - TableBindings</span></span><br><span data-ttu-id="653f9-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-215">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-216">
        - TextBindings</span></span><br><span data-ttu-id="653f9-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-218">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="653f9-218">Office 2013 on Windows</span></span><br><span data-ttu-id="653f9-219">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="653f9-220">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-220">
        - TaskPane</span></span><br><span data-ttu-id="653f9-221">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="653f9-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="653f9-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="653f9-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="653f9-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-224">
        - BindingEvents</span></span><br><span data-ttu-id="653f9-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-225">
        - CompressedFile</span></span><br><span data-ttu-id="653f9-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-226">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-227">
        - File</span></span><br><span data-ttu-id="653f9-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-228">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-230">
        - Selection</span></span><br><span data-ttu-id="653f9-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-231">
        - Settings</span></span><br><span data-ttu-id="653f9-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-232">
        - TableBindings</span></span><br><span data-ttu-id="653f9-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-233">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-234">
        - TextBindings</span></span><br><span data-ttu-id="653f9-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-236">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="653f9-237">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="653f9-238">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-238">- TaskPane</span></span><br><span data-ttu-id="653f9-239">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-239">
        - Content</span></span><br><span data-ttu-id="653f9-240">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="653f9-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="653f9-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="653f9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="653f9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="653f9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="653f9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="653f9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="653f9-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="653f9-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="653f9-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="653f9-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="653f9-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="653f9-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="653f9-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-252">- BindingEvents</span></span><br><span data-ttu-id="653f9-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-253">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-254">
        - File</span></span><br><span data-ttu-id="653f9-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-255">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-257">
        - Selection</span></span><br><span data-ttu-id="653f9-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-258">
        - Settings</span></span><br><span data-ttu-id="653f9-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-259">
        - TableBindings</span></span><br><span data-ttu-id="653f9-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-260">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-261">
        - TextBindings</span></span><br><span data-ttu-id="653f9-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-263">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-263">Office apps on Mac</span></span><br><span data-ttu-id="653f9-264">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="653f9-265">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-265">- TaskPane</span></span><br><span data-ttu-id="653f9-266">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-266">
        - Content</span></span><br><span data-ttu-id="653f9-267">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="653f9-267">
        - Custom Functions</span></span><br><span data-ttu-id="653f9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="653f9-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="653f9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="653f9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="653f9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="653f9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="653f9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="653f9-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="653f9-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="653f9-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="653f9-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="653f9-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="653f9-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="653f9-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-281">- BindingEvents</span></span><br><span data-ttu-id="653f9-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-282">
        - CompressedFile</span></span><br><span data-ttu-id="653f9-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-283">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-284">
        - File</span></span><br><span data-ttu-id="653f9-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-285">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-287">
        - PdfFile</span></span><br><span data-ttu-id="653f9-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-288">
        - Selection</span></span><br><span data-ttu-id="653f9-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-289">
        - Settings</span></span><br><span data-ttu-id="653f9-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-290">
        - TableBindings</span></span><br><span data-ttu-id="653f9-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-291">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-292">
        - TextBindings</span></span><br><span data-ttu-id="653f9-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-294">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-294">Office 2019 for Mac</span></span><br><span data-ttu-id="653f9-295">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="653f9-296">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-296">- TaskPane</span></span><br><span data-ttu-id="653f9-297">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-297">
        - Content</span></span><br><span data-ttu-id="653f9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="653f9-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="653f9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="653f9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="653f9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="653f9-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="653f9-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="653f9-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="653f9-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="653f9-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="653f9-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="653f9-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-309">- BindingEvents</span></span><br><span data-ttu-id="653f9-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-310">
        - CompressedFile</span></span><br><span data-ttu-id="653f9-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-311">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-312">
        - File</span></span><br><span data-ttu-id="653f9-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-313">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-315">
        - PdfFile</span></span><br><span data-ttu-id="653f9-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-316">
        - Selection</span></span><br><span data-ttu-id="653f9-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-317">
        - Settings</span></span><br><span data-ttu-id="653f9-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-318">
        - TableBindings</span></span><br><span data-ttu-id="653f9-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-319">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-320">
        - TextBindings</span></span><br><span data-ttu-id="653f9-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-322">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="653f9-323">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="653f9-324">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-324">- TaskPane</span></span><br><span data-ttu-id="653f9-325">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-325">
        - Content</span></span></td>
    <td><span data-ttu-id="653f9-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="653f9-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="653f9-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="653f9-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="653f9-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-329">- BindingEvents</span></span><br><span data-ttu-id="653f9-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-330">
        - CompressedFile</span></span><br><span data-ttu-id="653f9-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-331">
        - DocumentEvents</span></span><br><span data-ttu-id="653f9-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="653f9-332">
        - File</span></span><br><span data-ttu-id="653f9-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-333">
        - MatrixBindings</span></span><br><span data-ttu-id="653f9-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="653f9-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-335">
        - PdfFile</span></span><br><span data-ttu-id="653f9-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-336">
        - Selection</span></span><br><span data-ttu-id="653f9-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-337">
        - Settings</span></span><br><span data-ttu-id="653f9-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-338">
        - TableBindings</span></span><br><span data-ttu-id="653f9-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-339">
        - TableCoercion</span></span><br><span data-ttu-id="653f9-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-340">
        - TextBindings</span></span><br><span data-ttu-id="653f9-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="653f9-342">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="653f9-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="653f9-343">自定义函数</span><span class="sxs-lookup"><span data-stu-id="653f9-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="653f9-344">平台</span><span class="sxs-lookup"><span data-stu-id="653f9-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="653f9-345">扩展点</span><span class="sxs-lookup"><span data-stu-id="653f9-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="653f9-346">API 要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="653f9-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="653f9-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-348">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="653f9-348">Office on the web</span></span></td>
    <td><span data-ttu-id="653f9-349">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="653f9-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="653f9-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-351">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-351">Office on Windows</span></span><br><span data-ttu-id="653f9-352">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="653f9-353">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="653f9-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="653f9-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-355">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="653f9-355">Office for Mac</span></span><br><span data-ttu-id="653f9-356">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="653f9-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="653f9-357">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="653f9-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="653f9-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="653f9-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="653f9-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="653f9-360">平台</span><span class="sxs-lookup"><span data-stu-id="653f9-360">Platform</span></span></th>
    <th><span data-ttu-id="653f9-361">扩展点</span><span class="sxs-lookup"><span data-stu-id="653f9-361">Extension points</span></span></th>
    <th><span data-ttu-id="653f9-362">API 要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="653f9-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="653f9-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-364">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="653f9-364">Office on the web</span></span><br><span data-ttu-id="653f9-365">（新式）</span><span class="sxs-lookup"><span data-stu-id="653f9-365">Modern</span></span></td>
    <td> <span data-ttu-id="653f9-366">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-366">- Mail Read</span></span><br><span data-ttu-id="653f9-367">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-367">
      - Mail Compose</span></span><br><span data-ttu-id="653f9-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="653f9-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="653f9-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="653f9-376">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-377">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="653f9-377">Office on the web</span></span><br><span data-ttu-id="653f9-378">（经典）</span><span class="sxs-lookup"><span data-stu-id="653f9-378">Classic.</span></span></td>
    <td> <span data-ttu-id="653f9-379">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-379">- Mail Read</span></span><br><span data-ttu-id="653f9-380">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-380">
      - Mail Compose</span></span><br><span data-ttu-id="653f9-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="653f9-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="653f9-388">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-389">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-389">Office on Windows</span></span><br><span data-ttu-id="653f9-390">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-391">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-391">- Mail Read</span></span><br><span data-ttu-id="653f9-392">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-392">
      - Mail Compose</span></span><br><span data-ttu-id="653f9-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="653f9-394">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="653f9-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="653f9-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="653f9-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="653f9-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="653f9-402">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-403">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-403">Office 2019 on Windows</span></span><br><span data-ttu-id="653f9-404">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-405">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-405">- Mail Read</span></span><br><span data-ttu-id="653f9-406">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-406">
      - Mail Compose</span></span><br><span data-ttu-id="653f9-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="653f9-408">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="653f9-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="653f9-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="653f9-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="653f9-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="653f9-416">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-417">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-417">Office 2016 on Windows</span></span><br><span data-ttu-id="653f9-418">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-419">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-419">- Mail Read</span></span><br><span data-ttu-id="653f9-420">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-420">
      - Mail Compose</span></span><br><span data-ttu-id="653f9-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="653f9-422">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="653f9-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="653f9-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="653f9-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="653f9-427">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-428">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="653f9-428">Office 2013 on Windows</span></span><br><span data-ttu-id="653f9-429">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-430">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-430">- Mail Read</span></span><br><span data-ttu-id="653f9-431">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="653f9-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="653f9-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="653f9-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="653f9-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="653f9-436">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-437">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-437">Office apps on iOS</span></span><br><span data-ttu-id="653f9-438">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-439">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-439">- Mail Read</span></span><br><span data-ttu-id="653f9-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="653f9-446">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-447">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-447">Office apps on Mac</span></span><br><span data-ttu-id="653f9-448">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-449">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-449">- Mail Read</span></span><br><span data-ttu-id="653f9-450">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-450">
      - Mail Compose</span></span><br><span data-ttu-id="653f9-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="653f9-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="653f9-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="653f9-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="653f9-459">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-460">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-460">Office 2019 for Mac</span></span><br><span data-ttu-id="653f9-461">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-462">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-462">- Mail Read</span></span><br><span data-ttu-id="653f9-463">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-463">
      - Mail Compose</span></span><br><span data-ttu-id="653f9-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="653f9-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="653f9-471">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-472">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="653f9-473">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-474">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-474">- Mail Read</span></span><br><span data-ttu-id="653f9-475">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="653f9-475">
      - Mail Compose</span></span><br><span data-ttu-id="653f9-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="653f9-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="653f9-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="653f9-483">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-484">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-484">Office apps on Android</span></span><br><span data-ttu-id="653f9-485">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-486">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="653f9-486">- Mail Read</span></span><br><span data-ttu-id="653f9-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="653f9-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="653f9-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="653f9-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="653f9-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="653f9-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="653f9-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="653f9-493">不可用</span><span class="sxs-lookup"><span data-stu-id="653f9-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="653f9-494">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="653f9-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="653f9-495">Word</span><span class="sxs-lookup"><span data-stu-id="653f9-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="653f9-496">平台</span><span class="sxs-lookup"><span data-stu-id="653f9-496">Platform</span></span></th>
    <th><span data-ttu-id="653f9-497">扩展点</span><span class="sxs-lookup"><span data-stu-id="653f9-497">Extension points</span></span></th>
    <th><span data-ttu-id="653f9-498">API 要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="653f9-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="653f9-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-500">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="653f9-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="653f9-501">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-501">- TaskPane</span></span><br><span data-ttu-id="653f9-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="653f9-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="653f9-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="653f9-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="653f9-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-509">- BindingEvents</span></span><br><span data-ttu-id="653f9-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-511">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-512">
         - File</span></span><br><span data-ttu-id="653f9-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-514">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-517">
         - PdfFile</span></span><br><span data-ttu-id="653f9-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-518">
         - Selection</span></span><br><span data-ttu-id="653f9-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-519">
         - Settings</span></span><br><span data-ttu-id="653f9-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-520">
         - TableBindings</span></span><br><span data-ttu-id="653f9-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-521">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-522">
         - TextBindings</span></span><br><span data-ttu-id="653f9-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-523">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-525">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-525">Office on Windows</span></span><br><span data-ttu-id="653f9-526">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-527">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-527">- TaskPane</span></span><br><span data-ttu-id="653f9-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="653f9-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="653f9-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="653f9-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="653f9-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-535">- BindingEvents</span></span><br><span data-ttu-id="653f9-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-536">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-538">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-539">
         - File</span></span><br><span data-ttu-id="653f9-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-541">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-544">
         - PdfFile</span></span><br><span data-ttu-id="653f9-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-545">
         - Selection</span></span><br><span data-ttu-id="653f9-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-546">
         - Settings</span></span><br><span data-ttu-id="653f9-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-547">
         - TableBindings</span></span><br><span data-ttu-id="653f9-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-548">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-549">
         - TextBindings</span></span><br><span data-ttu-id="653f9-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-550">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-552">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-552">Office 2019 on Windows</span></span><br><span data-ttu-id="653f9-553">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-554">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-554">- TaskPane</span></span><br><span data-ttu-id="653f9-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="653f9-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="653f9-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="653f9-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-561">- BindingEvents</span></span><br><span data-ttu-id="653f9-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-562">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-564">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-565">
         - File</span></span><br><span data-ttu-id="653f9-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-567">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-570">
         - PdfFile</span></span><br><span data-ttu-id="653f9-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-571">
         - Selection</span></span><br><span data-ttu-id="653f9-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-572">
         - Settings</span></span><br><span data-ttu-id="653f9-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-573">
         - TableBindings</span></span><br><span data-ttu-id="653f9-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-574">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-575">
         - TextBindings</span></span><br><span data-ttu-id="653f9-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-576">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-578">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-578">Office 2016 on Windows</span></span><br><span data-ttu-id="653f9-579">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-580">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="653f9-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="653f9-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="653f9-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-584">- BindingEvents</span></span><br><span data-ttu-id="653f9-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-585">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-587">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-588">
         - File</span></span><br><span data-ttu-id="653f9-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-590">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-593">
         - PdfFile</span></span><br><span data-ttu-id="653f9-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-594">
         - Selection</span></span><br><span data-ttu-id="653f9-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-595">
         - Settings</span></span><br><span data-ttu-id="653f9-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-596">
         - TableBindings</span></span><br><span data-ttu-id="653f9-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-597">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-598">
         - TextBindings</span></span><br><span data-ttu-id="653f9-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-599">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-601">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="653f9-601">Office 2013 on Windows</span></span><br><span data-ttu-id="653f9-602">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-603">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="653f9-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="653f9-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-606">- BindingEvents</span></span><br><span data-ttu-id="653f9-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-607">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-609">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-610">
         - File</span></span><br><span data-ttu-id="653f9-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-612">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-615">
         - PdfFile</span></span><br><span data-ttu-id="653f9-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-616">
         - Selection</span></span><br><span data-ttu-id="653f9-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-617">
         - Settings</span></span><br><span data-ttu-id="653f9-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-618">
         - TableBindings</span></span><br><span data-ttu-id="653f9-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-619">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-620">
         - TextBindings</span></span><br><span data-ttu-id="653f9-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-621">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-623">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="653f9-624">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-625">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="653f9-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="653f9-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="653f9-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="653f9-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-631">- BindingEvents</span></span><br><span data-ttu-id="653f9-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-632">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-634">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-635">
         - File</span></span><br><span data-ttu-id="653f9-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-637">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-640">
         - PdfFile</span></span><br><span data-ttu-id="653f9-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-641">
         - Selection</span></span><br><span data-ttu-id="653f9-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-642">
         - Settings</span></span><br><span data-ttu-id="653f9-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-643">
         - TableBindings</span></span><br><span data-ttu-id="653f9-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-644">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-645">
         - TextBindings</span></span><br><span data-ttu-id="653f9-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-646">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-648">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-648">Office apps on Mac</span></span><br><span data-ttu-id="653f9-649">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-650">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-650">- TaskPane</span></span><br><span data-ttu-id="653f9-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="653f9-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="653f9-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="653f9-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="653f9-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-658">- BindingEvents</span></span><br><span data-ttu-id="653f9-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-659">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-661">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-662">
         - File</span></span><br><span data-ttu-id="653f9-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-664">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-667">
         - PdfFile</span></span><br><span data-ttu-id="653f9-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-668">
         - Selection</span></span><br><span data-ttu-id="653f9-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-669">
         - Settings</span></span><br><span data-ttu-id="653f9-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-670">
         - TableBindings</span></span><br><span data-ttu-id="653f9-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-671">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-672">
         - TextBindings</span></span><br><span data-ttu-id="653f9-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-673">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-675">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-675">Office 2019 for Mac</span></span><br><span data-ttu-id="653f9-676">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-677">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-677">- TaskPane</span></span><br><span data-ttu-id="653f9-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="653f9-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="653f9-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="653f9-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="653f9-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="653f9-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-684">- BindingEvents</span></span><br><span data-ttu-id="653f9-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-685">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-687">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-688">
         - File</span></span><br><span data-ttu-id="653f9-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-690">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-693">
         - PdfFile</span></span><br><span data-ttu-id="653f9-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-694">
         - Selection</span></span><br><span data-ttu-id="653f9-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-695">
         - Settings</span></span><br><span data-ttu-id="653f9-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-696">
         - TableBindings</span></span><br><span data-ttu-id="653f9-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-697">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-698">
         - TextBindings</span></span><br><span data-ttu-id="653f9-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-699">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-701">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="653f9-702">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-703">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="653f9-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="653f9-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="653f9-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-707">- BindingEvents</span></span><br><span data-ttu-id="653f9-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-708">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="653f9-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="653f9-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-710">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-711">
         - File</span></span><br><span data-ttu-id="653f9-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-713">
         - MatrixBindings</span></span><br><span data-ttu-id="653f9-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="653f9-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="653f9-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-716">
         - PdfFile</span></span><br><span data-ttu-id="653f9-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-717">
         - Selection</span></span><br><span data-ttu-id="653f9-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-718">
         - Settings</span></span><br><span data-ttu-id="653f9-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-719">
         - TableBindings</span></span><br><span data-ttu-id="653f9-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-720">
         - TableCoercion</span></span><br><span data-ttu-id="653f9-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="653f9-721">
         - TextBindings</span></span><br><span data-ttu-id="653f9-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-722">
         - TextCoercion</span></span><br><span data-ttu-id="653f9-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="653f9-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="653f9-724">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="653f9-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="653f9-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="653f9-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="653f9-726">平台</span><span class="sxs-lookup"><span data-stu-id="653f9-726">Platform</span></span></th>
    <th><span data-ttu-id="653f9-727">扩展点</span><span class="sxs-lookup"><span data-stu-id="653f9-727">Extension points</span></span></th>
    <th><span data-ttu-id="653f9-728">API 要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="653f9-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="653f9-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-730">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="653f9-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="653f9-731">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-731">- Content</span></span><br><span data-ttu-id="653f9-732">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-732">
         - TaskPane</span></span><br><span data-ttu-id="653f9-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="653f9-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-737">- ActiveView</span></span><br><span data-ttu-id="653f9-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-738">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-739">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-740">
         - File</span></span><br><span data-ttu-id="653f9-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-741">
         - PdfFile</span></span><br><span data-ttu-id="653f9-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-742">
         - Selection</span></span><br><span data-ttu-id="653f9-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-743">
         - Settings</span></span><br><span data-ttu-id="653f9-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-745">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-745">Office on Windows</span></span><br><span data-ttu-id="653f9-746">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-747">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-747">- Content</span></span><br><span data-ttu-id="653f9-748">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-748">
         - TaskPane</span></span><br><span data-ttu-id="653f9-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="653f9-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-753">- ActiveView</span></span><br><span data-ttu-id="653f9-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-754">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-755">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-756">
         - File</span></span><br><span data-ttu-id="653f9-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-757">
         - PdfFile</span></span><br><span data-ttu-id="653f9-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-758">
         - Selection</span></span><br><span data-ttu-id="653f9-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-759">
         - Settings</span></span><br><span data-ttu-id="653f9-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-761">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-761">Office 2019 on Windows</span></span><br><span data-ttu-id="653f9-762">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-763">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-763">- Content</span></span><br><span data-ttu-id="653f9-764">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-764">
         - TaskPane</span></span><br><span data-ttu-id="653f9-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-768">- ActiveView</span></span><br><span data-ttu-id="653f9-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-769">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-770">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-771">
         - File</span></span><br><span data-ttu-id="653f9-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-772">
         - PdfFile</span></span><br><span data-ttu-id="653f9-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-773">
         - Selection</span></span><br><span data-ttu-id="653f9-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-774">
         - Settings</span></span><br><span data-ttu-id="653f9-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-776">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-776">Office 2016 on Windows</span></span><br><span data-ttu-id="653f9-777">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-778">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-778">- Content</span></span><br><span data-ttu-id="653f9-779">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="653f9-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="653f9-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-782">- ActiveView</span></span><br><span data-ttu-id="653f9-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-783">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-784">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-785">
         - File</span></span><br><span data-ttu-id="653f9-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-786">
         - PdfFile</span></span><br><span data-ttu-id="653f9-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-787">
         - Selection</span></span><br><span data-ttu-id="653f9-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-788">
         - Settings</span></span><br><span data-ttu-id="653f9-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-790">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="653f9-790">Office 2013 on Windows</span></span><br><span data-ttu-id="653f9-791">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-792">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-792">- Content</span></span><br><span data-ttu-id="653f9-793">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="653f9-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="653f9-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="653f9-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-796">- ActiveView</span></span><br><span data-ttu-id="653f9-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-797">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-798">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-799">
         - File</span></span><br><span data-ttu-id="653f9-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-800">
         - PdfFile</span></span><br><span data-ttu-id="653f9-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-801">
         - Selection</span></span><br><span data-ttu-id="653f9-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-802">
         - Settings</span></span><br><span data-ttu-id="653f9-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-804">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="653f9-805">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-806">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-806">- Content</span></span><br><span data-ttu-id="653f9-807">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-810">- ActiveView</span></span><br><span data-ttu-id="653f9-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-811">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-812">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-813">
         - File</span></span><br><span data-ttu-id="653f9-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-814">
         - PdfFile</span></span><br><span data-ttu-id="653f9-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-815">
         - Selection</span></span><br><span data-ttu-id="653f9-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-816">
         - Settings</span></span><br><span data-ttu-id="653f9-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-818">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="653f9-818">Office apps on Mac</span></span><br><span data-ttu-id="653f9-819">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="653f9-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="653f9-820">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-820">- Content</span></span><br><span data-ttu-id="653f9-821">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-821">
         - TaskPane</span></span><br><span data-ttu-id="653f9-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="653f9-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="653f9-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="653f9-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-826">- ActiveView</span></span><br><span data-ttu-id="653f9-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-827">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-828">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-829">
         - File</span></span><br><span data-ttu-id="653f9-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-830">
         - PdfFile</span></span><br><span data-ttu-id="653f9-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-831">
         - Selection</span></span><br><span data-ttu-id="653f9-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-832">
         - Settings</span></span><br><span data-ttu-id="653f9-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-834">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-834">Office 2019 for Mac</span></span><br><span data-ttu-id="653f9-835">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-836">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-836">- Content</span></span><br><span data-ttu-id="653f9-837">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-837">
         - TaskPane</span></span><br><span data-ttu-id="653f9-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-841">- ActiveView</span></span><br><span data-ttu-id="653f9-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-842">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-843">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-844">
         - File</span></span><br><span data-ttu-id="653f9-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-845">
         - PdfFile</span></span><br><span data-ttu-id="653f9-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-846">
         - Selection</span></span><br><span data-ttu-id="653f9-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-847">
         - Settings</span></span><br><span data-ttu-id="653f9-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-849">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="653f9-850">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-851">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-851">- Content</span></span><br><span data-ttu-id="653f9-852">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="653f9-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="653f9-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="653f9-855">- ActiveView</span></span><br><span data-ttu-id="653f9-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="653f9-856">
         - CompressedFile</span></span><br><span data-ttu-id="653f9-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-857">
         - DocumentEvents</span></span><br><span data-ttu-id="653f9-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="653f9-858">
         - File</span></span><br><span data-ttu-id="653f9-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="653f9-859">
         - PdfFile</span></span><br><span data-ttu-id="653f9-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-860">
         - Selection</span></span><br><span data-ttu-id="653f9-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-861">
         - Settings</span></span><br><span data-ttu-id="653f9-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="653f9-863">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="653f9-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="653f9-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="653f9-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="653f9-865">平台</span><span class="sxs-lookup"><span data-stu-id="653f9-865">Platform</span></span></th>
    <th><span data-ttu-id="653f9-866">扩展点</span><span class="sxs-lookup"><span data-stu-id="653f9-866">Extension points</span></span></th>
    <th><span data-ttu-id="653f9-867">API 要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="653f9-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="653f9-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-869">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="653f9-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="653f9-870">- 内容</span><span class="sxs-lookup"><span data-stu-id="653f9-870">- Content</span></span><br><span data-ttu-id="653f9-871">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-871">
         - TaskPane</span></span><br><span data-ttu-id="653f9-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="653f9-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="653f9-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="653f9-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="653f9-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="653f9-876">- DocumentEvents</span></span><br><span data-ttu-id="653f9-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="653f9-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="653f9-878">
         - Settings</span></span><br><span data-ttu-id="653f9-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="653f9-880">项目</span><span class="sxs-lookup"><span data-stu-id="653f9-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="653f9-881">平台</span><span class="sxs-lookup"><span data-stu-id="653f9-881">Platform</span></span></th>
    <th><span data-ttu-id="653f9-882">扩展点</span><span class="sxs-lookup"><span data-stu-id="653f9-882">Extension points</span></span></th>
    <th><span data-ttu-id="653f9-883">API 要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="653f9-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="653f9-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-885">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="653f9-885">Office 2019 on Windows</span></span><br><span data-ttu-id="653f9-886">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-887">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-889">- Selection</span></span><br><span data-ttu-id="653f9-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-891">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="653f9-891">Office 2016 on Windows</span></span><br><span data-ttu-id="653f9-892">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-893">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-895">- Selection</span></span><br><span data-ttu-id="653f9-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="653f9-897">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="653f9-897">Office 2013 on Windows</span></span><br><span data-ttu-id="653f9-898">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="653f9-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="653f9-899">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="653f9-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="653f9-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="653f9-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="653f9-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="653f9-901">- Selection</span></span><br><span data-ttu-id="653f9-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="653f9-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="653f9-903">另请参阅</span><span class="sxs-lookup"><span data-stu-id="653f9-903">See also</span></span>

- [<span data-ttu-id="653f9-904">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="653f9-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="653f9-905">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="653f9-906">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="653f9-907">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="653f9-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="653f9-908">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="653f9-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="653f9-909">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="653f9-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="653f9-910">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="653f9-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="653f9-911">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="653f9-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="653f9-912">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="653f9-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="653f9-913">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="653f9-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="653f9-914">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="653f9-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
