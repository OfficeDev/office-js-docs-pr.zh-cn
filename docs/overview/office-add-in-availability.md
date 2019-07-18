---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: 2bfeb7cc5c6e8846f1d882abf3a0149302e53914
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771833"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="2d727-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="2d727-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="2d727-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="2d727-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="2d727-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="2d727-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="2d727-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="2d727-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="2d727-108">Excel</span><span class="sxs-lookup"><span data-stu-id="2d727-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="2d727-109">平台</span><span class="sxs-lookup"><span data-stu-id="2d727-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="2d727-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="2d727-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="2d727-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="2d727-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2d727-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2d727-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="2d727-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-114">- TaskPane</span></span><br><span data-ttu-id="2d727-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-115">
        - Content</span></span><br><span data-ttu-id="2d727-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2d727-116">
        - Custom Functions</span></span><br><span data-ttu-id="2d727-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="2d727-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="2d727-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2d727-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2d727-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2d727-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2d727-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2d727-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2d727-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2d727-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2d727-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2d727-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="2d727-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="2d727-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="2d727-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-130">
        - BindingEvents</span></span><br><span data-ttu-id="2d727-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-131">
        - CompressedFile</span></span><br><span data-ttu-id="2d727-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-132">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-133">
        - File</span></span><br><span data-ttu-id="2d727-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-134">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-136">
        - Selection</span></span><br><span data-ttu-id="2d727-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-137">
        - Settings</span></span><br><span data-ttu-id="2d727-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-138">
        - TableBindings</span></span><br><span data-ttu-id="2d727-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-139">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-140">
        - TextBindings</span></span><br><span data-ttu-id="2d727-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-142">Office on Windows</span></span><br><span data-ttu-id="2d727-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-144">- TaskPane</span></span><br><span data-ttu-id="2d727-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-145">
        - Content</span></span><br><span data-ttu-id="2d727-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2d727-146">
        - Custom Functions</span></span><br><span data-ttu-id="2d727-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="2d727-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="2d727-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2d727-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2d727-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2d727-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2d727-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2d727-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2d727-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2d727-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2d727-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2d727-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="2d727-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="2d727-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="2d727-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-160">
        - BindingEvents</span></span><br><span data-ttu-id="2d727-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-161">
        - CompressedFile</span></span><br><span data-ttu-id="2d727-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-162">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-163">
        - File</span></span><br><span data-ttu-id="2d727-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-164">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-166">
        - Selection</span></span><br><span data-ttu-id="2d727-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-167">
        - Settings</span></span><br><span data-ttu-id="2d727-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-168">
        - TableBindings</span></span><br><span data-ttu-id="2d727-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-169">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-170">
        - TextBindings</span></span><br><span data-ttu-id="2d727-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-172">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-172">Office 2019 on Windows</span></span><br><span data-ttu-id="2d727-173">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2d727-174">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-174">- TaskPane</span></span><br><span data-ttu-id="2d727-175">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-175">
        - Content</span></span><br><span data-ttu-id="2d727-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2d727-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2d727-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2d727-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2d727-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2d727-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2d727-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2d727-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2d727-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2d727-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2d727-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2d727-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-187">- BindingEvents</span></span><br><span data-ttu-id="2d727-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-188">
        - CompressedFile</span></span><br><span data-ttu-id="2d727-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-189">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-190">
        - File</span></span><br><span data-ttu-id="2d727-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-191">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-193">
        - Selection</span></span><br><span data-ttu-id="2d727-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-194">
        - Settings</span></span><br><span data-ttu-id="2d727-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-195">
        - TableBindings</span></span><br><span data-ttu-id="2d727-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-196">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-197">
        - TextBindings</span></span><br><span data-ttu-id="2d727-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-199">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-199">Office 2016 on Windows</span></span><br><span data-ttu-id="2d727-200">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2d727-201">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-201">- TaskPane</span></span><br><span data-ttu-id="2d727-202">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-202">
        - Content</span></span></td>
    <td><span data-ttu-id="2d727-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2d727-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="2d727-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="2d727-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2d727-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-206">- BindingEvents</span></span><br><span data-ttu-id="2d727-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-207">
        - CompressedFile</span></span><br><span data-ttu-id="2d727-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-208">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-209">
        - File</span></span><br><span data-ttu-id="2d727-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-210">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-212">
        - Selection</span></span><br><span data-ttu-id="2d727-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-213">
        - Settings</span></span><br><span data-ttu-id="2d727-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-214">
        - TableBindings</span></span><br><span data-ttu-id="2d727-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-215">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-216">
        - TextBindings</span></span><br><span data-ttu-id="2d727-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-218">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2d727-218">Office 2013 on Windows</span></span><br><span data-ttu-id="2d727-219">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2d727-220">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-220">
        - TaskPane</span></span><br><span data-ttu-id="2d727-221">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="2d727-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2d727-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2d727-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2d727-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-224">
        - BindingEvents</span></span><br><span data-ttu-id="2d727-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-225">
        - CompressedFile</span></span><br><span data-ttu-id="2d727-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-226">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-227">
        - File</span></span><br><span data-ttu-id="2d727-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-228">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-230">
        - Selection</span></span><br><span data-ttu-id="2d727-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-231">
        - Settings</span></span><br><span data-ttu-id="2d727-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-232">
        - TableBindings</span></span><br><span data-ttu-id="2d727-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-233">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-234">
        - TextBindings</span></span><br><span data-ttu-id="2d727-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-236">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="2d727-237">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="2d727-238">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-238">- TaskPane</span></span><br><span data-ttu-id="2d727-239">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-239">
        - Content</span></span><br><span data-ttu-id="2d727-240">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2d727-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="2d727-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2d727-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2d727-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2d727-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2d727-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2d727-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2d727-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2d727-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2d727-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2d727-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="2d727-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="2d727-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2d727-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-252">- BindingEvents</span></span><br><span data-ttu-id="2d727-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-253">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-254">
        - File</span></span><br><span data-ttu-id="2d727-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-255">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-257">
        - Selection</span></span><br><span data-ttu-id="2d727-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-258">
        - Settings</span></span><br><span data-ttu-id="2d727-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-259">
        - TableBindings</span></span><br><span data-ttu-id="2d727-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-260">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-261">
        - TextBindings</span></span><br><span data-ttu-id="2d727-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-263">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-263">Office apps on Mac</span></span><br><span data-ttu-id="2d727-264">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="2d727-265">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-265">- TaskPane</span></span><br><span data-ttu-id="2d727-266">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-266">
        - Content</span></span><br><span data-ttu-id="2d727-267">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2d727-267">
        - Custom Functions</span></span><br><span data-ttu-id="2d727-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2d727-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2d727-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2d727-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2d727-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2d727-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2d727-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2d727-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2d727-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2d727-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2d727-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="2d727-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="2d727-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="2d727-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-281">- BindingEvents</span></span><br><span data-ttu-id="2d727-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-282">
        - CompressedFile</span></span><br><span data-ttu-id="2d727-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-283">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-284">
        - File</span></span><br><span data-ttu-id="2d727-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-285">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-287">
        - PdfFile</span></span><br><span data-ttu-id="2d727-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-288">
        - Selection</span></span><br><span data-ttu-id="2d727-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-289">
        - Settings</span></span><br><span data-ttu-id="2d727-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-290">
        - TableBindings</span></span><br><span data-ttu-id="2d727-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-291">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-292">
        - TextBindings</span></span><br><span data-ttu-id="2d727-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-294">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-294">Office 2019 for Mac</span></span><br><span data-ttu-id="2d727-295">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2d727-296">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-296">- TaskPane</span></span><br><span data-ttu-id="2d727-297">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-297">
        - Content</span></span><br><span data-ttu-id="2d727-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2d727-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2d727-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2d727-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2d727-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2d727-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2d727-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2d727-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2d727-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2d727-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2d727-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2d727-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-309">- BindingEvents</span></span><br><span data-ttu-id="2d727-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-310">
        - CompressedFile</span></span><br><span data-ttu-id="2d727-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-311">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-312">
        - File</span></span><br><span data-ttu-id="2d727-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-313">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-315">
        - PdfFile</span></span><br><span data-ttu-id="2d727-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-316">
        - Selection</span></span><br><span data-ttu-id="2d727-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-317">
        - Settings</span></span><br><span data-ttu-id="2d727-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-318">
        - TableBindings</span></span><br><span data-ttu-id="2d727-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-319">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-320">
        - TextBindings</span></span><br><span data-ttu-id="2d727-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-322">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="2d727-323">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2d727-324">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-324">- TaskPane</span></span><br><span data-ttu-id="2d727-325">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-325">
        - Content</span></span></td>
    <td><span data-ttu-id="2d727-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2d727-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="2d727-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="2d727-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2d727-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-329">- BindingEvents</span></span><br><span data-ttu-id="2d727-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-330">
        - CompressedFile</span></span><br><span data-ttu-id="2d727-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-331">
        - DocumentEvents</span></span><br><span data-ttu-id="2d727-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="2d727-332">
        - File</span></span><br><span data-ttu-id="2d727-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-333">
        - MatrixBindings</span></span><br><span data-ttu-id="2d727-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="2d727-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-335">
        - PdfFile</span></span><br><span data-ttu-id="2d727-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-336">
        - Selection</span></span><br><span data-ttu-id="2d727-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-337">
        - Settings</span></span><br><span data-ttu-id="2d727-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-338">
        - TableBindings</span></span><br><span data-ttu-id="2d727-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-339">
        - TableCoercion</span></span><br><span data-ttu-id="2d727-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-340">
        - TextBindings</span></span><br><span data-ttu-id="2d727-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="2d727-342">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="2d727-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="2d727-343">自定义函数</span><span class="sxs-lookup"><span data-stu-id="2d727-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="2d727-344">平台</span><span class="sxs-lookup"><span data-stu-id="2d727-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="2d727-345">扩展点</span><span class="sxs-lookup"><span data-stu-id="2d727-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="2d727-346">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="2d727-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2d727-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-348">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2d727-348">Office on the web</span></span></td>
    <td><span data-ttu-id="2d727-349">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2d727-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="2d727-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-351">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-351">Office on Windows</span></span><br><span data-ttu-id="2d727-352">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="2d727-353">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2d727-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="2d727-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-355">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="2d727-355">Office for Mac</span></span><br><span data-ttu-id="2d727-356">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="2d727-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="2d727-357">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2d727-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="2d727-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="2d727-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="2d727-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2d727-360">平台</span><span class="sxs-lookup"><span data-stu-id="2d727-360">Platform</span></span></th>
    <th><span data-ttu-id="2d727-361">扩展点</span><span class="sxs-lookup"><span data-stu-id="2d727-361">Extension points</span></span></th>
    <th><span data-ttu-id="2d727-362">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="2d727-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2d727-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-364">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2d727-364">Office on the web</span></span><br><span data-ttu-id="2d727-365">（新）</span><span class="sxs-lookup"><span data-stu-id="2d727-365">New</span></span></td>
    <td> <span data-ttu-id="2d727-366">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-366">- Mail Read</span></span><br><span data-ttu-id="2d727-367">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-367">
      - Mail Compose</span></span><br><span data-ttu-id="2d727-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2d727-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2d727-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2d727-376">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-377">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2d727-377">Office on the web</span></span><br><span data-ttu-id="2d727-378">（经典）</span><span class="sxs-lookup"><span data-stu-id="2d727-378">Classic.</span></span></td>
    <td> <span data-ttu-id="2d727-379">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-379">- Mail Read</span></span><br><span data-ttu-id="2d727-380">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-380">
      - Mail Compose</span></span><br><span data-ttu-id="2d727-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2d727-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="2d727-388">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-389">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-389">Office on Windows</span></span><br><span data-ttu-id="2d727-390">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-391">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-391">- Mail Read</span></span><br><span data-ttu-id="2d727-392">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-392">
      - Mail Compose</span></span><br><span data-ttu-id="2d727-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="2d727-394">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="2d727-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="2d727-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2d727-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2d727-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2d727-402">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-403">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-403">Office 2019 on Windows</span></span><br><span data-ttu-id="2d727-404">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-405">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-405">- Mail Read</span></span><br><span data-ttu-id="2d727-406">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-406">
      - Mail Compose</span></span><br><span data-ttu-id="2d727-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="2d727-408">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="2d727-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="2d727-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2d727-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2d727-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2d727-416">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-417">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-417">Office 2016 on Windows</span></span><br><span data-ttu-id="2d727-418">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-419">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-419">- Mail Read</span></span><br><span data-ttu-id="2d727-420">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-420">
      - Mail Compose</span></span><br><span data-ttu-id="2d727-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="2d727-422">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="2d727-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="2d727-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="2d727-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="2d727-427">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-428">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2d727-428">Office 2013 on Windows</span></span><br><span data-ttu-id="2d727-429">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-430">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-430">- Mail Read</span></span><br><span data-ttu-id="2d727-431">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="2d727-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="2d727-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="2d727-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="2d727-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="2d727-436">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-437">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-437">Office apps on iOS</span></span><br><span data-ttu-id="2d727-438">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-439">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-439">- Mail Read</span></span><br><span data-ttu-id="2d727-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="2d727-446">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-447">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-447">Office apps on Mac</span></span><br><span data-ttu-id="2d727-448">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-449">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-449">- Mail Read</span></span><br><span data-ttu-id="2d727-450">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-450">
      - Mail Compose</span></span><br><span data-ttu-id="2d727-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2d727-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2d727-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2d727-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2d727-459">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-460">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-460">Office 2019 for Mac</span></span><br><span data-ttu-id="2d727-461">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-462">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-462">- Mail Read</span></span><br><span data-ttu-id="2d727-463">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-463">
      - Mail Compose</span></span><br><span data-ttu-id="2d727-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2d727-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="2d727-471">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-472">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="2d727-473">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-474">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-474">- Mail Read</span></span><br><span data-ttu-id="2d727-475">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2d727-475">
      - Mail Compose</span></span><br><span data-ttu-id="2d727-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2d727-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2d727-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="2d727-483">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-484">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-484">Office apps on Android</span></span><br><span data-ttu-id="2d727-485">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-486">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2d727-486">- Mail Read</span></span><br><span data-ttu-id="2d727-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2d727-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2d727-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2d727-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2d727-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2d727-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2d727-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="2d727-493">不可用</span><span class="sxs-lookup"><span data-stu-id="2d727-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="2d727-494">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="2d727-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="2d727-495">Word</span><span class="sxs-lookup"><span data-stu-id="2d727-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2d727-496">平台</span><span class="sxs-lookup"><span data-stu-id="2d727-496">Platform</span></span></th>
    <th><span data-ttu-id="2d727-497">扩展点</span><span class="sxs-lookup"><span data-stu-id="2d727-497">Extension points</span></span></th>
    <th><span data-ttu-id="2d727-498">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="2d727-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2d727-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-500">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2d727-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="2d727-501">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-501">- TaskPane</span></span><br><span data-ttu-id="2d727-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2d727-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2d727-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2d727-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2d727-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-509">- BindingEvents</span></span><br><span data-ttu-id="2d727-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-511">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-512">
         - File</span></span><br><span data-ttu-id="2d727-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-514">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-517">
         - PdfFile</span></span><br><span data-ttu-id="2d727-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-518">
         - Selection</span></span><br><span data-ttu-id="2d727-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-519">
         - Settings</span></span><br><span data-ttu-id="2d727-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-520">
         - TableBindings</span></span><br><span data-ttu-id="2d727-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-521">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-522">
         - TextBindings</span></span><br><span data-ttu-id="2d727-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-523">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-525">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-525">Office on Windows</span></span><br><span data-ttu-id="2d727-526">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-527">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-527">- TaskPane</span></span><br><span data-ttu-id="2d727-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2d727-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2d727-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2d727-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2d727-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-535">- BindingEvents</span></span><br><span data-ttu-id="2d727-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-536">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-538">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-539">
         - File</span></span><br><span data-ttu-id="2d727-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-541">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-544">
         - PdfFile</span></span><br><span data-ttu-id="2d727-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-545">
         - Selection</span></span><br><span data-ttu-id="2d727-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-546">
         - Settings</span></span><br><span data-ttu-id="2d727-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-547">
         - TableBindings</span></span><br><span data-ttu-id="2d727-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-548">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-549">
         - TextBindings</span></span><br><span data-ttu-id="2d727-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-550">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-552">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-552">Office 2019 on Windows</span></span><br><span data-ttu-id="2d727-553">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-554">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-554">- TaskPane</span></span><br><span data-ttu-id="2d727-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2d727-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2d727-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2d727-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-561">- BindingEvents</span></span><br><span data-ttu-id="2d727-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-562">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-564">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-565">
         - File</span></span><br><span data-ttu-id="2d727-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-567">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-570">
         - PdfFile</span></span><br><span data-ttu-id="2d727-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-571">
         - Selection</span></span><br><span data-ttu-id="2d727-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-572">
         - Settings</span></span><br><span data-ttu-id="2d727-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-573">
         - TableBindings</span></span><br><span data-ttu-id="2d727-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-574">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-575">
         - TextBindings</span></span><br><span data-ttu-id="2d727-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-576">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-578">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-578">Office 2016 on Windows</span></span><br><span data-ttu-id="2d727-579">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-580">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2d727-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="2d727-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="2d727-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-584">- BindingEvents</span></span><br><span data-ttu-id="2d727-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-585">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-587">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-588">
         - File</span></span><br><span data-ttu-id="2d727-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-590">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-593">
         - PdfFile</span></span><br><span data-ttu-id="2d727-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-594">
         - Selection</span></span><br><span data-ttu-id="2d727-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-595">
         - Settings</span></span><br><span data-ttu-id="2d727-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-596">
         - TableBindings</span></span><br><span data-ttu-id="2d727-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-597">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-598">
         - TextBindings</span></span><br><span data-ttu-id="2d727-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-599">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-601">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2d727-601">Office 2013 on Windows</span></span><br><span data-ttu-id="2d727-602">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-603">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2d727-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2d727-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-606">- BindingEvents</span></span><br><span data-ttu-id="2d727-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-607">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-609">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-610">
         - File</span></span><br><span data-ttu-id="2d727-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-612">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-615">
         - PdfFile</span></span><br><span data-ttu-id="2d727-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-616">
         - Selection</span></span><br><span data-ttu-id="2d727-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-617">
         - Settings</span></span><br><span data-ttu-id="2d727-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-618">
         - TableBindings</span></span><br><span data-ttu-id="2d727-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-619">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-620">
         - TextBindings</span></span><br><span data-ttu-id="2d727-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-621">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-623">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="2d727-624">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-625">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2d727-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2d727-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2d727-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="2d727-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-631">- BindingEvents</span></span><br><span data-ttu-id="2d727-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-632">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-634">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-635">
         - File</span></span><br><span data-ttu-id="2d727-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-637">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-640">
         - PdfFile</span></span><br><span data-ttu-id="2d727-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-641">
         - Selection</span></span><br><span data-ttu-id="2d727-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-642">
         - Settings</span></span><br><span data-ttu-id="2d727-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-643">
         - TableBindings</span></span><br><span data-ttu-id="2d727-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-644">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-645">
         - TextBindings</span></span><br><span data-ttu-id="2d727-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-646">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-648">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-648">Office apps on Mac</span></span><br><span data-ttu-id="2d727-649">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-650">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-650">- TaskPane</span></span><br><span data-ttu-id="2d727-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2d727-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2d727-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2d727-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="2d727-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-658">- BindingEvents</span></span><br><span data-ttu-id="2d727-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-659">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-661">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-662">
         - File</span></span><br><span data-ttu-id="2d727-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-664">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-667">
         - PdfFile</span></span><br><span data-ttu-id="2d727-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-668">
         - Selection</span></span><br><span data-ttu-id="2d727-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-669">
         - Settings</span></span><br><span data-ttu-id="2d727-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-670">
         - TableBindings</span></span><br><span data-ttu-id="2d727-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-671">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-672">
         - TextBindings</span></span><br><span data-ttu-id="2d727-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-673">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-675">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-675">Office 2019 for Mac</span></span><br><span data-ttu-id="2d727-676">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-677">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-677">- TaskPane</span></span><br><span data-ttu-id="2d727-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2d727-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2d727-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2d727-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2d727-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="2d727-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-684">- BindingEvents</span></span><br><span data-ttu-id="2d727-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-685">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-687">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-688">
         - File</span></span><br><span data-ttu-id="2d727-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-690">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-693">
         - PdfFile</span></span><br><span data-ttu-id="2d727-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-694">
         - Selection</span></span><br><span data-ttu-id="2d727-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-695">
         - Settings</span></span><br><span data-ttu-id="2d727-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-696">
         - TableBindings</span></span><br><span data-ttu-id="2d727-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-697">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-698">
         - TextBindings</span></span><br><span data-ttu-id="2d727-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-699">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-701">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="2d727-702">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-703">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2d727-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="2d727-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="2d727-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-707">- BindingEvents</span></span><br><span data-ttu-id="2d727-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-708">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2d727-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="2d727-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-710">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-711">
         - File</span></span><br><span data-ttu-id="2d727-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-713">
         - MatrixBindings</span></span><br><span data-ttu-id="2d727-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="2d727-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2d727-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-716">
         - PdfFile</span></span><br><span data-ttu-id="2d727-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-717">
         - Selection</span></span><br><span data-ttu-id="2d727-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-718">
         - Settings</span></span><br><span data-ttu-id="2d727-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-719">
         - TableBindings</span></span><br><span data-ttu-id="2d727-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-720">
         - TableCoercion</span></span><br><span data-ttu-id="2d727-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2d727-721">
         - TextBindings</span></span><br><span data-ttu-id="2d727-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-722">
         - TextCoercion</span></span><br><span data-ttu-id="2d727-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2d727-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="2d727-724">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="2d727-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="2d727-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="2d727-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2d727-726">平台</span><span class="sxs-lookup"><span data-stu-id="2d727-726">Platform</span></span></th>
    <th><span data-ttu-id="2d727-727">扩展点</span><span class="sxs-lookup"><span data-stu-id="2d727-727">Extension points</span></span></th>
    <th><span data-ttu-id="2d727-728">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="2d727-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2d727-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-730">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2d727-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="2d727-731">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-731">- Content</span></span><br><span data-ttu-id="2d727-732">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-732">
         - TaskPane</span></span><br><span data-ttu-id="2d727-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2d727-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-737">- ActiveView</span></span><br><span data-ttu-id="2d727-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-738">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-739">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-740">
         - File</span></span><br><span data-ttu-id="2d727-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-741">
         - PdfFile</span></span><br><span data-ttu-id="2d727-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-742">
         - Selection</span></span><br><span data-ttu-id="2d727-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-743">
         - Settings</span></span><br><span data-ttu-id="2d727-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-745">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-745">Office on Windows</span></span><br><span data-ttu-id="2d727-746">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-747">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-747">- Content</span></span><br><span data-ttu-id="2d727-748">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-748">
         - TaskPane</span></span><br><span data-ttu-id="2d727-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2d727-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-753">- ActiveView</span></span><br><span data-ttu-id="2d727-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-754">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-755">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-756">
         - File</span></span><br><span data-ttu-id="2d727-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-757">
         - PdfFile</span></span><br><span data-ttu-id="2d727-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-758">
         - Selection</span></span><br><span data-ttu-id="2d727-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-759">
         - Settings</span></span><br><span data-ttu-id="2d727-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-761">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-761">Office 2019 on Windows</span></span><br><span data-ttu-id="2d727-762">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-763">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-763">- Content</span></span><br><span data-ttu-id="2d727-764">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-764">
         - TaskPane</span></span><br><span data-ttu-id="2d727-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-768">- ActiveView</span></span><br><span data-ttu-id="2d727-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-769">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-770">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-771">
         - File</span></span><br><span data-ttu-id="2d727-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-772">
         - PdfFile</span></span><br><span data-ttu-id="2d727-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-773">
         - Selection</span></span><br><span data-ttu-id="2d727-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-774">
         - Settings</span></span><br><span data-ttu-id="2d727-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-776">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-776">Office 2016 on Windows</span></span><br><span data-ttu-id="2d727-777">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-778">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-778">- Content</span></span><br><span data-ttu-id="2d727-779">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2d727-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2d727-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-782">- ActiveView</span></span><br><span data-ttu-id="2d727-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-783">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-784">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-785">
         - File</span></span><br><span data-ttu-id="2d727-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-786">
         - PdfFile</span></span><br><span data-ttu-id="2d727-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-787">
         - Selection</span></span><br><span data-ttu-id="2d727-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-788">
         - Settings</span></span><br><span data-ttu-id="2d727-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-790">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2d727-790">Office 2013 on Windows</span></span><br><span data-ttu-id="2d727-791">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-792">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-792">- Content</span></span><br><span data-ttu-id="2d727-793">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="2d727-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2d727-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2d727-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-796">- ActiveView</span></span><br><span data-ttu-id="2d727-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-797">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-798">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-799">
         - File</span></span><br><span data-ttu-id="2d727-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-800">
         - PdfFile</span></span><br><span data-ttu-id="2d727-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-801">
         - Selection</span></span><br><span data-ttu-id="2d727-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-802">
         - Settings</span></span><br><span data-ttu-id="2d727-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-804">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="2d727-805">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-806">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-806">- Content</span></span><br><span data-ttu-id="2d727-807">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-810">- ActiveView</span></span><br><span data-ttu-id="2d727-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-811">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-812">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-813">
         - File</span></span><br><span data-ttu-id="2d727-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-814">
         - PdfFile</span></span><br><span data-ttu-id="2d727-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-815">
         - Selection</span></span><br><span data-ttu-id="2d727-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-816">
         - Settings</span></span><br><span data-ttu-id="2d727-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-818">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2d727-818">Office apps on Mac</span></span><br><span data-ttu-id="2d727-819">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2d727-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2d727-820">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-820">- Content</span></span><br><span data-ttu-id="2d727-821">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-821">
         - TaskPane</span></span><br><span data-ttu-id="2d727-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2d727-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2d727-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2d727-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-826">- ActiveView</span></span><br><span data-ttu-id="2d727-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-827">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-828">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-829">
         - File</span></span><br><span data-ttu-id="2d727-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-830">
         - PdfFile</span></span><br><span data-ttu-id="2d727-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-831">
         - Selection</span></span><br><span data-ttu-id="2d727-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-832">
         - Settings</span></span><br><span data-ttu-id="2d727-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-834">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-834">Office 2019 for Mac</span></span><br><span data-ttu-id="2d727-835">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-836">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-836">- Content</span></span><br><span data-ttu-id="2d727-837">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-837">
         - TaskPane</span></span><br><span data-ttu-id="2d727-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-841">- ActiveView</span></span><br><span data-ttu-id="2d727-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-842">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-843">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-844">
         - File</span></span><br><span data-ttu-id="2d727-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-845">
         - PdfFile</span></span><br><span data-ttu-id="2d727-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-846">
         - Selection</span></span><br><span data-ttu-id="2d727-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-847">
         - Settings</span></span><br><span data-ttu-id="2d727-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-849">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="2d727-850">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-851">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-851">- Content</span></span><br><span data-ttu-id="2d727-852">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2d727-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2d727-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2d727-855">- ActiveView</span></span><br><span data-ttu-id="2d727-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2d727-856">
         - CompressedFile</span></span><br><span data-ttu-id="2d727-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-857">
         - DocumentEvents</span></span><br><span data-ttu-id="2d727-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="2d727-858">
         - File</span></span><br><span data-ttu-id="2d727-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2d727-859">
         - PdfFile</span></span><br><span data-ttu-id="2d727-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-860">
         - Selection</span></span><br><span data-ttu-id="2d727-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-861">
         - Settings</span></span><br><span data-ttu-id="2d727-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="2d727-863">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="2d727-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="2d727-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="2d727-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2d727-865">平台</span><span class="sxs-lookup"><span data-stu-id="2d727-865">Platform</span></span></th>
    <th><span data-ttu-id="2d727-866">扩展点</span><span class="sxs-lookup"><span data-stu-id="2d727-866">Extension points</span></span></th>
    <th><span data-ttu-id="2d727-867">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="2d727-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2d727-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-869">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2d727-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="2d727-870">- 内容</span><span class="sxs-lookup"><span data-stu-id="2d727-870">- Content</span></span><br><span data-ttu-id="2d727-871">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-871">
         - TaskPane</span></span><br><span data-ttu-id="2d727-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2d727-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2d727-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="2d727-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2d727-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2d727-876">- DocumentEvents</span></span><br><span data-ttu-id="2d727-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="2d727-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2d727-878">
         - Settings</span></span><br><span data-ttu-id="2d727-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="2d727-880">项目</span><span class="sxs-lookup"><span data-stu-id="2d727-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2d727-881">平台</span><span class="sxs-lookup"><span data-stu-id="2d727-881">Platform</span></span></th>
    <th><span data-ttu-id="2d727-882">扩展点</span><span class="sxs-lookup"><span data-stu-id="2d727-882">Extension points</span></span></th>
    <th><span data-ttu-id="2d727-883">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="2d727-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2d727-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-885">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2d727-885">Office 2019 on Windows</span></span><br><span data-ttu-id="2d727-886">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-887">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-889">- Selection</span></span><br><span data-ttu-id="2d727-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-891">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2d727-891">Office 2016 on Windows</span></span><br><span data-ttu-id="2d727-892">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-893">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-895">- Selection</span></span><br><span data-ttu-id="2d727-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2d727-897">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2d727-897">Office 2013 on Windows</span></span><br><span data-ttu-id="2d727-898">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2d727-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2d727-899">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2d727-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2d727-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2d727-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2d727-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="2d727-901">- Selection</span></span><br><span data-ttu-id="2d727-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2d727-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="2d727-903">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2d727-903">See also</span></span>

- [<span data-ttu-id="2d727-904">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="2d727-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="2d727-905">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="2d727-906">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="2d727-907">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="2d727-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="2d727-908">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="2d727-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="2d727-909">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="2d727-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="2d727-910">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="2d727-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="2d727-911">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="2d727-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="2d727-912">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="2d727-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="2d727-913">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="2d727-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="2d727-914">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="2d727-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
