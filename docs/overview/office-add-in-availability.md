---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: 956ee6b8a9e990a3d6d942ee4a65a1e9275ea025
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851367"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="59dc7-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="59dc7-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="59dc7-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="59dc7-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="59dc7-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="59dc7-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="59dc7-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="59dc7-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="59dc7-108">Excel</span><span class="sxs-lookup"><span data-stu-id="59dc7-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="59dc7-109">平台</span><span class="sxs-lookup"><span data-stu-id="59dc7-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="59dc7-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="59dc7-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="59dc7-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="59dc7-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="59dc7-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="59dc7-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="59dc7-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-114">- TaskPane</span></span><br><span data-ttu-id="59dc7-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-115">
        - Content</span></span><br><span data-ttu-id="59dc7-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="59dc7-116">
        - Custom Functions</span></span><br><span data-ttu-id="59dc7-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="59dc7-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="59dc7-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="59dc7-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="59dc7-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="59dc7-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="59dc7-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="59dc7-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="59dc7-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="59dc7-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="59dc7-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="59dc7-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="59dc7-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="59dc7-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="59dc7-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-130">
        - BindingEvents</span></span><br><span data-ttu-id="59dc7-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-131">
        - CompressedFile</span></span><br><span data-ttu-id="59dc7-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-132">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-133">
        - File</span></span><br><span data-ttu-id="59dc7-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-134">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-136">
        - Selection</span></span><br><span data-ttu-id="59dc7-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-137">
        - Settings</span></span><br><span data-ttu-id="59dc7-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-138">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-139">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-140">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-142">Office on Windows</span></span><br><span data-ttu-id="59dc7-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-144">- TaskPane</span></span><br><span data-ttu-id="59dc7-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-145">
        - Content</span></span><br><span data-ttu-id="59dc7-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="59dc7-146">
        - Custom Functions</span></span><br><span data-ttu-id="59dc7-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="59dc7-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="59dc7-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="59dc7-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="59dc7-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="59dc7-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="59dc7-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="59dc7-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="59dc7-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="59dc7-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="59dc7-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="59dc7-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="59dc7-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="59dc7-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="59dc7-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-161">
        - BindingEvents</span></span><br><span data-ttu-id="59dc7-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-162">
        - CompressedFile</span></span><br><span data-ttu-id="59dc7-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-163">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-164">
        - File</span></span><br><span data-ttu-id="59dc7-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-165">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-167">
        - Selection</span></span><br><span data-ttu-id="59dc7-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-168">
        - Settings</span></span><br><span data-ttu-id="59dc7-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-169">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-170">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-171">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-173">Office 2019 on Windows</span></span><br><span data-ttu-id="59dc7-174">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="59dc7-175">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-175">- TaskPane</span></span><br><span data-ttu-id="59dc7-176">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-176">
        - Content</span></span><br><span data-ttu-id="59dc7-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="59dc7-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="59dc7-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="59dc7-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="59dc7-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="59dc7-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="59dc7-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="59dc7-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="59dc7-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="59dc7-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="59dc7-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-188">- BindingEvents</span></span><br><span data-ttu-id="59dc7-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-189">
        - CompressedFile</span></span><br><span data-ttu-id="59dc7-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-190">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-191">
        - File</span></span><br><span data-ttu-id="59dc7-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-192">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-194">
        - Selection</span></span><br><span data-ttu-id="59dc7-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-195">
        - Settings</span></span><br><span data-ttu-id="59dc7-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-196">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-197">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-198">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-200">Office 2016 on Windows</span></span><br><span data-ttu-id="59dc7-201">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="59dc7-202">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-202">- TaskPane</span></span><br><span data-ttu-id="59dc7-203">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-203">
        - Content</span></span></td>
    <td><span data-ttu-id="59dc7-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="59dc7-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="59dc7-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="59dc7-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="59dc7-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-207">- BindingEvents</span></span><br><span data-ttu-id="59dc7-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-208">
        - CompressedFile</span></span><br><span data-ttu-id="59dc7-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-209">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-210">
        - File</span></span><br><span data-ttu-id="59dc7-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-211">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-213">
        - Selection</span></span><br><span data-ttu-id="59dc7-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-214">
        - Settings</span></span><br><span data-ttu-id="59dc7-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-215">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-216">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-217">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="59dc7-219">Office 2013 on Windows</span></span><br><span data-ttu-id="59dc7-220">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="59dc7-221">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-221">
        - TaskPane</span></span><br><span data-ttu-id="59dc7-222">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="59dc7-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="59dc7-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="59dc7-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="59dc7-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-225">
        - BindingEvents</span></span><br><span data-ttu-id="59dc7-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-226">
        - CompressedFile</span></span><br><span data-ttu-id="59dc7-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-227">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-228">
        - File</span></span><br><span data-ttu-id="59dc7-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-229">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-231">
        - Selection</span></span><br><span data-ttu-id="59dc7-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-232">
        - Settings</span></span><br><span data-ttu-id="59dc7-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-233">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-234">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-235">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-237">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-237">Office on iPad</span></span><br><span data-ttu-id="59dc7-238">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="59dc7-239">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-239">- TaskPane</span></span><br><span data-ttu-id="59dc7-240">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-240">
        - Content</span></span></td>
    <td><span data-ttu-id="59dc7-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="59dc7-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="59dc7-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="59dc7-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="59dc7-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="59dc7-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="59dc7-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="59dc7-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="59dc7-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="59dc7-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="59dc7-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="59dc7-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-253">- BindingEvents</span></span><br><span data-ttu-id="59dc7-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-254">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-255">
        - File</span></span><br><span data-ttu-id="59dc7-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-256">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-258">
        - Selection</span></span><br><span data-ttu-id="59dc7-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-259">
        - Settings</span></span><br><span data-ttu-id="59dc7-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-260">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-261">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-262">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-264">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-264">Office on Mac</span></span><br><span data-ttu-id="59dc7-265">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="59dc7-266">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-266">- TaskPane</span></span><br><span data-ttu-id="59dc7-267">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-267">
        - Content</span></span><br><span data-ttu-id="59dc7-268">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="59dc7-268">
        - Custom Functions</span></span><br><span data-ttu-id="59dc7-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="59dc7-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="59dc7-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="59dc7-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="59dc7-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="59dc7-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="59dc7-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="59dc7-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="59dc7-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="59dc7-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="59dc7-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="59dc7-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="59dc7-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="59dc7-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-283">- BindingEvents</span></span><br><span data-ttu-id="59dc7-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-284">
        - CompressedFile</span></span><br><span data-ttu-id="59dc7-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-285">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-286">
        - File</span></span><br><span data-ttu-id="59dc7-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-287">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-289">
        - PdfFile</span></span><br><span data-ttu-id="59dc7-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-290">
        - Selection</span></span><br><span data-ttu-id="59dc7-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-291">
        - Settings</span></span><br><span data-ttu-id="59dc7-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-292">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-293">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-294">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-296">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-296">Office 2019 on Mac</span></span><br><span data-ttu-id="59dc7-297">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="59dc7-298">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-298">- TaskPane</span></span><br><span data-ttu-id="59dc7-299">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-299">
        - Content</span></span><br><span data-ttu-id="59dc7-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="59dc7-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="59dc7-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="59dc7-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="59dc7-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="59dc7-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="59dc7-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="59dc7-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="59dc7-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="59dc7-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="59dc7-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-311">- BindingEvents</span></span><br><span data-ttu-id="59dc7-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-312">
        - CompressedFile</span></span><br><span data-ttu-id="59dc7-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-313">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-314">
        - File</span></span><br><span data-ttu-id="59dc7-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-315">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-317">
        - PdfFile</span></span><br><span data-ttu-id="59dc7-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-318">
        - Selection</span></span><br><span data-ttu-id="59dc7-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-319">
        - Settings</span></span><br><span data-ttu-id="59dc7-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-320">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-321">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-322">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-324">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-324">Office 2016 on Mac</span></span><br><span data-ttu-id="59dc7-325">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="59dc7-326">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-326">- TaskPane</span></span><br><span data-ttu-id="59dc7-327">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-327">
        - Content</span></span></td>
    <td><span data-ttu-id="59dc7-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="59dc7-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="59dc7-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="59dc7-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="59dc7-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-331">- BindingEvents</span></span><br><span data-ttu-id="59dc7-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-332">
        - CompressedFile</span></span><br><span data-ttu-id="59dc7-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-333">
        - DocumentEvents</span></span><br><span data-ttu-id="59dc7-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-334">
        - File</span></span><br><span data-ttu-id="59dc7-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-335">
        - MatrixBindings</span></span><br><span data-ttu-id="59dc7-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-337">
        - PdfFile</span></span><br><span data-ttu-id="59dc7-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-338">
        - Selection</span></span><br><span data-ttu-id="59dc7-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-339">
        - Settings</span></span><br><span data-ttu-id="59dc7-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-340">
        - TableBindings</span></span><br><span data-ttu-id="59dc7-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-341">
        - TableCoercion</span></span><br><span data-ttu-id="59dc7-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-342">
        - TextBindings</span></span><br><span data-ttu-id="59dc7-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="59dc7-344">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="59dc7-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="59dc7-345">自定义函数</span><span class="sxs-lookup"><span data-stu-id="59dc7-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="59dc7-346">平台</span><span class="sxs-lookup"><span data-stu-id="59dc7-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="59dc7-347">扩展点</span><span class="sxs-lookup"><span data-stu-id="59dc7-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="59dc7-348">API 要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="59dc7-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="59dc7-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-350">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="59dc7-350">Office on the web</span></span></td>
    <td><span data-ttu-id="59dc7-351">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="59dc7-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="59dc7-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-353">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-353">Office on Windows</span></span><br><span data-ttu-id="59dc7-354">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="59dc7-355">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="59dc7-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="59dc7-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="59dc7-357">Office for Mac</span></span><br><span data-ttu-id="59dc7-358">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="59dc7-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="59dc7-359">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="59dc7-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="59dc7-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="59dc7-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="59dc7-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="59dc7-362">平台</span><span class="sxs-lookup"><span data-stu-id="59dc7-362">Platform</span></span></th>
    <th><span data-ttu-id="59dc7-363">扩展点</span><span class="sxs-lookup"><span data-stu-id="59dc7-363">Extension points</span></span></th>
    <th><span data-ttu-id="59dc7-364">API 要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="59dc7-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="59dc7-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-366">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="59dc7-366">Office on the web</span></span><br><span data-ttu-id="59dc7-367">（新式）</span><span class="sxs-lookup"><span data-stu-id="59dc7-367">(modern)</span></span></td>
    <td> <span data-ttu-id="59dc7-368">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-368">- Mail Read</span></span><br><span data-ttu-id="59dc7-369">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-369">
      - Mail Compose</span></span><br><span data-ttu-id="59dc7-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="59dc7-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="59dc7-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="59dc7-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="59dc7-379">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-380">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="59dc7-380">Office on the web</span></span><br><span data-ttu-id="59dc7-381">（经典）</span><span class="sxs-lookup"><span data-stu-id="59dc7-381">(classic)</span></span></td>
    <td> <span data-ttu-id="59dc7-382">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-382">- Mail Read</span></span><br><span data-ttu-id="59dc7-383">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-383">
      - Mail Compose</span></span><br><span data-ttu-id="59dc7-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="59dc7-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="59dc7-391">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-392">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-392">Office on Windows</span></span><br><span data-ttu-id="59dc7-393">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-394">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-394">- Mail Read</span></span><br><span data-ttu-id="59dc7-395">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-395">
      - Mail Compose</span></span><br><span data-ttu-id="59dc7-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="59dc7-397">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="59dc7-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="59dc7-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="59dc7-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="59dc7-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="59dc7-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="59dc7-406">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-407">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-407">Office 2019 on Windows</span></span><br><span data-ttu-id="59dc7-408">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-409">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-409">- Mail Read</span></span><br><span data-ttu-id="59dc7-410">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-410">
      - Mail Compose</span></span><br><span data-ttu-id="59dc7-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="59dc7-412">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="59dc7-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="59dc7-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="59dc7-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="59dc7-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="59dc7-420">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-421">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-421">Office 2016 on Windows</span></span><br><span data-ttu-id="59dc7-422">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-423">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-423">- Mail Read</span></span><br><span data-ttu-id="59dc7-424">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-424">
      - Mail Compose</span></span><br><span data-ttu-id="59dc7-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="59dc7-426">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="59dc7-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="59dc7-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="59dc7-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="59dc7-431">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-432">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="59dc7-432">Office 2013 on Windows</span></span><br><span data-ttu-id="59dc7-433">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-434">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-434">- Mail Read</span></span><br><span data-ttu-id="59dc7-435">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="59dc7-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="59dc7-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="59dc7-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="59dc7-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="59dc7-440">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-441">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-441">Office on iOS</span></span><br><span data-ttu-id="59dc7-442">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-443">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-443">- Mail Read</span></span><br><span data-ttu-id="59dc7-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="59dc7-450">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-451">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-451">Office on Mac</span></span><br><span data-ttu-id="59dc7-452">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-453">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-453">- Mail Read</span></span><br><span data-ttu-id="59dc7-454">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-454">
      - Mail Compose</span></span><br><span data-ttu-id="59dc7-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="59dc7-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="59dc7-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="59dc7-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="59dc7-464">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-465">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-465">Office 2019 on Mac</span></span><br><span data-ttu-id="59dc7-466">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-467">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-467">- Mail Read</span></span><br><span data-ttu-id="59dc7-468">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-468">
      - Mail Compose</span></span><br><span data-ttu-id="59dc7-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="59dc7-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="59dc7-476">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-477">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-477">Office 2016 on Mac</span></span><br><span data-ttu-id="59dc7-478">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-479">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-479">- Mail Read</span></span><br><span data-ttu-id="59dc7-480">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="59dc7-480">
      - Mail Compose</span></span><br><span data-ttu-id="59dc7-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="59dc7-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="59dc7-488">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-489">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-489">Office on Android</span></span><br><span data-ttu-id="59dc7-490">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-491">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="59dc7-491">- Mail Read</span></span><br><span data-ttu-id="59dc7-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="59dc7-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="59dc7-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="59dc7-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="59dc7-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="59dc7-498">不可用</span><span class="sxs-lookup"><span data-stu-id="59dc7-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="59dc7-499">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="59dc7-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="59dc7-500">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="59dc7-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="59dc7-501">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="59dc7-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="59dc7-502">Word</span><span class="sxs-lookup"><span data-stu-id="59dc7-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="59dc7-503">平台</span><span class="sxs-lookup"><span data-stu-id="59dc7-503">Platform</span></span></th>
    <th><span data-ttu-id="59dc7-504">扩展点</span><span class="sxs-lookup"><span data-stu-id="59dc7-504">Extension points</span></span></th>
    <th><span data-ttu-id="59dc7-505">API 要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="59dc7-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="59dc7-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-507">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="59dc7-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="59dc7-508">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-508">- TaskPane</span></span><br><span data-ttu-id="59dc7-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="59dc7-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="59dc7-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="59dc7-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="59dc7-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="59dc7-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-516">- BindingEvents</span></span><br><span data-ttu-id="59dc7-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-518">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-519">
         - File</span></span><br><span data-ttu-id="59dc7-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-521">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-524">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-525">
         - Selection</span></span><br><span data-ttu-id="59dc7-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-526">
         - Settings</span></span><br><span data-ttu-id="59dc7-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-527">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-528">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-529">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-530">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-532">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-532">Office on Windows</span></span><br><span data-ttu-id="59dc7-533">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-534">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-534">- TaskPane</span></span><br><span data-ttu-id="59dc7-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="59dc7-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="59dc7-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="59dc7-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="59dc7-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="59dc7-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-542">- BindingEvents</span></span><br><span data-ttu-id="59dc7-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-543">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-545">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-546">
         - File</span></span><br><span data-ttu-id="59dc7-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-548">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-551">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-552">
         - Selection</span></span><br><span data-ttu-id="59dc7-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-553">
         - Settings</span></span><br><span data-ttu-id="59dc7-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-554">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-555">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-556">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-557">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-559">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-559">Office 2019 on Windows</span></span><br><span data-ttu-id="59dc7-560">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-561">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-561">- TaskPane</span></span><br><span data-ttu-id="59dc7-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="59dc7-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="59dc7-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="59dc7-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-568">- BindingEvents</span></span><br><span data-ttu-id="59dc7-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-569">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-571">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-572">
         - File</span></span><br><span data-ttu-id="59dc7-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-574">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-577">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-578">
         - Selection</span></span><br><span data-ttu-id="59dc7-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-579">
         - Settings</span></span><br><span data-ttu-id="59dc7-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-580">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-581">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-582">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-583">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-585">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-585">Office 2016 on Windows</span></span><br><span data-ttu-id="59dc7-586">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-587">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="59dc7-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="59dc7-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="59dc7-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-591">- BindingEvents</span></span><br><span data-ttu-id="59dc7-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-592">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-594">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-595">
         - File</span></span><br><span data-ttu-id="59dc7-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-597">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-600">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-601">
         - Selection</span></span><br><span data-ttu-id="59dc7-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-602">
         - Settings</span></span><br><span data-ttu-id="59dc7-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-603">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-604">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-605">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-606">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-608">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="59dc7-608">Office 2013 on Windows</span></span><br><span data-ttu-id="59dc7-609">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-610">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="59dc7-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="59dc7-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-613">- BindingEvents</span></span><br><span data-ttu-id="59dc7-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-614">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-616">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-617">
         - File</span></span><br><span data-ttu-id="59dc7-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-619">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-622">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-623">
         - Selection</span></span><br><span data-ttu-id="59dc7-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-624">
         - Settings</span></span><br><span data-ttu-id="59dc7-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-625">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-626">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-627">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-628">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-630">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-630">Office on iPad</span></span><br><span data-ttu-id="59dc7-631">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-632">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="59dc7-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="59dc7-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="59dc7-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="59dc7-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-638">- BindingEvents</span></span><br><span data-ttu-id="59dc7-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-639">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-641">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-642">
         - File</span></span><br><span data-ttu-id="59dc7-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-644">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-647">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-648">
         - Selection</span></span><br><span data-ttu-id="59dc7-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-649">
         - Settings</span></span><br><span data-ttu-id="59dc7-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-650">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-651">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-652">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-653">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-655">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-655">Office on Mac</span></span><br><span data-ttu-id="59dc7-656">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-657">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-657">- TaskPane</span></span><br><span data-ttu-id="59dc7-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="59dc7-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="59dc7-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="59dc7-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="59dc7-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="59dc7-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-665">- BindingEvents</span></span><br><span data-ttu-id="59dc7-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-666">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-668">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-669">
         - File</span></span><br><span data-ttu-id="59dc7-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-671">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-674">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-675">
         - Selection</span></span><br><span data-ttu-id="59dc7-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-676">
         - Settings</span></span><br><span data-ttu-id="59dc7-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-677">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-678">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-679">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-680">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-682">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-682">Office 2019 on Mac</span></span><br><span data-ttu-id="59dc7-683">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-684">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-684">- TaskPane</span></span><br><span data-ttu-id="59dc7-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="59dc7-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="59dc7-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="59dc7-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="59dc7-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-691">- BindingEvents</span></span><br><span data-ttu-id="59dc7-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-692">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-694">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-695">
         - File</span></span><br><span data-ttu-id="59dc7-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-697">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-700">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-701">
         - Selection</span></span><br><span data-ttu-id="59dc7-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-702">
         - Settings</span></span><br><span data-ttu-id="59dc7-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-703">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-704">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-705">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-706">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-708">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-708">Office 2016 on Mac</span></span><br><span data-ttu-id="59dc7-709">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-710">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="59dc7-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="59dc7-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="59dc7-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-714">- BindingEvents</span></span><br><span data-ttu-id="59dc7-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-715">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="59dc7-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="59dc7-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-717">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-718">
         - File</span></span><br><span data-ttu-id="59dc7-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-720">
         - MatrixBindings</span></span><br><span data-ttu-id="59dc7-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="59dc7-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="59dc7-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-723">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-724">
         - Selection</span></span><br><span data-ttu-id="59dc7-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-725">
         - Settings</span></span><br><span data-ttu-id="59dc7-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-726">
         - TableBindings</span></span><br><span data-ttu-id="59dc7-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-727">
         - TableCoercion</span></span><br><span data-ttu-id="59dc7-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="59dc7-728">
         - TextBindings</span></span><br><span data-ttu-id="59dc7-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-729">
         - TextCoercion</span></span><br><span data-ttu-id="59dc7-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="59dc7-731">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="59dc7-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="59dc7-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="59dc7-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="59dc7-733">平台</span><span class="sxs-lookup"><span data-stu-id="59dc7-733">Platform</span></span></th>
    <th><span data-ttu-id="59dc7-734">扩展点</span><span class="sxs-lookup"><span data-stu-id="59dc7-734">Extension points</span></span></th>
    <th><span data-ttu-id="59dc7-735">API 要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="59dc7-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="59dc7-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-737">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="59dc7-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="59dc7-738">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-738">- Content</span></span><br><span data-ttu-id="59dc7-739">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-739">
         - TaskPane</span></span><br><span data-ttu-id="59dc7-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="59dc7-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="59dc7-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="59dc7-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-745">- ActiveView</span></span><br><span data-ttu-id="59dc7-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-746">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-747">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-748">
         - File</span></span><br><span data-ttu-id="59dc7-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-749">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-750">
         - Selection</span></span><br><span data-ttu-id="59dc7-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-751">
         - Settings</span></span><br><span data-ttu-id="59dc7-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-753">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-753">Office on Windows</span></span><br><span data-ttu-id="59dc7-754">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-755">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-755">- Content</span></span><br><span data-ttu-id="59dc7-756">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-756">
         - TaskPane</span></span><br><span data-ttu-id="59dc7-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="59dc7-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="59dc7-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="59dc7-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-762">- ActiveView</span></span><br><span data-ttu-id="59dc7-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-763">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-764">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-765">
         - File</span></span><br><span data-ttu-id="59dc7-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-766">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-767">
         - Selection</span></span><br><span data-ttu-id="59dc7-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-768">
         - Settings</span></span><br><span data-ttu-id="59dc7-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-770">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-770">Office 2019 on Windows</span></span><br><span data-ttu-id="59dc7-771">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-772">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-772">- Content</span></span><br><span data-ttu-id="59dc7-773">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-773">
         - TaskPane</span></span><br><span data-ttu-id="59dc7-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-777">- ActiveView</span></span><br><span data-ttu-id="59dc7-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-778">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-779">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-780">
         - File</span></span><br><span data-ttu-id="59dc7-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-781">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-782">
         - Selection</span></span><br><span data-ttu-id="59dc7-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-783">
         - Settings</span></span><br><span data-ttu-id="59dc7-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-785">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-785">Office 2016 on Windows</span></span><br><span data-ttu-id="59dc7-786">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-787">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-787">- Content</span></span><br><span data-ttu-id="59dc7-788">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="59dc7-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="59dc7-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-791">- ActiveView</span></span><br><span data-ttu-id="59dc7-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-792">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-793">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-794">
         - File</span></span><br><span data-ttu-id="59dc7-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-795">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-796">
         - Selection</span></span><br><span data-ttu-id="59dc7-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-797">
         - Settings</span></span><br><span data-ttu-id="59dc7-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-799">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="59dc7-799">Office 2013 on Windows</span></span><br><span data-ttu-id="59dc7-800">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-801">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-801">- Content</span></span><br><span data-ttu-id="59dc7-802">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="59dc7-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="59dc7-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="59dc7-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-805">- ActiveView</span></span><br><span data-ttu-id="59dc7-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-806">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-807">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-808">
         - File</span></span><br><span data-ttu-id="59dc7-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-809">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-810">
         - Selection</span></span><br><span data-ttu-id="59dc7-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-811">
         - Settings</span></span><br><span data-ttu-id="59dc7-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-813">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-813">Office on iPad</span></span><br><span data-ttu-id="59dc7-814">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-815">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-815">- Content</span></span><br><span data-ttu-id="59dc7-816">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="59dc7-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-820">- ActiveView</span></span><br><span data-ttu-id="59dc7-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-821">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-822">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-823">
         - File</span></span><br><span data-ttu-id="59dc7-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-824">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-825">
         - Selection</span></span><br><span data-ttu-id="59dc7-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-826">
         - Settings</span></span><br><span data-ttu-id="59dc7-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-828">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="59dc7-828">Office on Mac</span></span><br><span data-ttu-id="59dc7-829">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="59dc7-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="59dc7-830">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-830">- Content</span></span><br><span data-ttu-id="59dc7-831">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-831">
         - TaskPane</span></span><br><span data-ttu-id="59dc7-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="59dc7-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="59dc7-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="59dc7-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-837">- ActiveView</span></span><br><span data-ttu-id="59dc7-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-838">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-839">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-840">
         - File</span></span><br><span data-ttu-id="59dc7-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-841">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-842">
         - Selection</span></span><br><span data-ttu-id="59dc7-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-843">
         - Settings</span></span><br><span data-ttu-id="59dc7-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-845">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-845">Office 2019 on Mac</span></span><br><span data-ttu-id="59dc7-846">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-847">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-847">- Content</span></span><br><span data-ttu-id="59dc7-848">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-848">
         - TaskPane</span></span><br><span data-ttu-id="59dc7-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-852">- ActiveView</span></span><br><span data-ttu-id="59dc7-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-853">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-854">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-855">
         - File</span></span><br><span data-ttu-id="59dc7-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-856">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-857">
         - Selection</span></span><br><span data-ttu-id="59dc7-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-858">
         - Settings</span></span><br><span data-ttu-id="59dc7-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-860">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-860">Office 2016 on Mac</span></span><br><span data-ttu-id="59dc7-861">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-862">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-862">- Content</span></span><br><span data-ttu-id="59dc7-863">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="59dc7-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="59dc7-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="59dc7-866">- ActiveView</span></span><br><span data-ttu-id="59dc7-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-867">
         - CompressedFile</span></span><br><span data-ttu-id="59dc7-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-868">
         - DocumentEvents</span></span><br><span data-ttu-id="59dc7-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="59dc7-869">
         - File</span></span><br><span data-ttu-id="59dc7-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="59dc7-870">
         - PdfFile</span></span><br><span data-ttu-id="59dc7-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-871">
         - Selection</span></span><br><span data-ttu-id="59dc7-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-872">
         - Settings</span></span><br><span data-ttu-id="59dc7-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="59dc7-874">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="59dc7-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="59dc7-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="59dc7-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="59dc7-876">平台</span><span class="sxs-lookup"><span data-stu-id="59dc7-876">Platform</span></span></th>
    <th><span data-ttu-id="59dc7-877">扩展点</span><span class="sxs-lookup"><span data-stu-id="59dc7-877">Extension points</span></span></th>
    <th><span data-ttu-id="59dc7-878">API 要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="59dc7-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="59dc7-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-880">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="59dc7-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="59dc7-881">- 内容</span><span class="sxs-lookup"><span data-stu-id="59dc7-881">- Content</span></span><br><span data-ttu-id="59dc7-882">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-882">
         - TaskPane</span></span><br><span data-ttu-id="59dc7-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="59dc7-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="59dc7-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="59dc7-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="59dc7-887">- DocumentEvents</span></span><br><span data-ttu-id="59dc7-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="59dc7-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="59dc7-889">
         - Settings</span></span><br><span data-ttu-id="59dc7-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="59dc7-891">项目</span><span class="sxs-lookup"><span data-stu-id="59dc7-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="59dc7-892">平台</span><span class="sxs-lookup"><span data-stu-id="59dc7-892">Platform</span></span></th>
    <th><span data-ttu-id="59dc7-893">扩展点</span><span class="sxs-lookup"><span data-stu-id="59dc7-893">Extension points</span></span></th>
    <th><span data-ttu-id="59dc7-894">API 要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="59dc7-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="59dc7-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-896">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="59dc7-896">Office 2019 on Windows</span></span><br><span data-ttu-id="59dc7-897">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-898">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-900">- Selection</span></span><br><span data-ttu-id="59dc7-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-902">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="59dc7-902">Office 2016 on Windows</span></span><br><span data-ttu-id="59dc7-903">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-904">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-906">- Selection</span></span><br><span data-ttu-id="59dc7-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="59dc7-908">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="59dc7-908">Office 2013 on Windows</span></span><br><span data-ttu-id="59dc7-909">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="59dc7-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="59dc7-910">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="59dc7-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="59dc7-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="59dc7-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="59dc7-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="59dc7-912">- Selection</span></span><br><span data-ttu-id="59dc7-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="59dc7-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="59dc7-914">另请参阅</span><span class="sxs-lookup"><span data-stu-id="59dc7-914">See also</span></span>

- [<span data-ttu-id="59dc7-915">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="59dc7-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="59dc7-916">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="59dc7-917">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="59dc7-918">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="59dc7-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="59dc7-919">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="59dc7-919">API reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="59dc7-920">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="59dc7-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="59dc7-921">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="59dc7-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="59dc7-922">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="59dc7-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="59dc7-923">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="59dc7-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="59dc7-924">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="59dc7-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="59dc7-925">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="59dc7-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="59dc7-926">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="59dc7-926">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)