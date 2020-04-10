---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 04/07/2020
localization_priority: Priority
ms.openlocfilehash: 823fd53e71c71f4a845f9a7b5c6177ad3f14745f
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185615"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c72a7-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="c72a7-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c72a7-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="c72a7-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="c72a7-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="c72a7-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="c72a7-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="c72a7-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="c72a7-108">Excel</span><span class="sxs-lookup"><span data-stu-id="c72a7-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c72a7-109">平台</span><span class="sxs-lookup"><span data-stu-id="c72a7-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c72a7-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="c72a7-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c72a7-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c72a7-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c72a7-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c72a7-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="c72a7-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-114">- TaskPane</span></span><br><span data-ttu-id="c72a7-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-115">
        - Content</span></span><br><span data-ttu-id="c72a7-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c72a7-116">
        - Custom Functions</span></span><br><span data-ttu-id="c72a7-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="c72a7-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c72a7-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c72a7-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c72a7-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c72a7-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c72a7-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c72a7-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c72a7-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c72a7-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c72a7-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c72a7-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c72a7-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="c72a7-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c72a7-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-130">
        - BindingEvents</span></span><br><span data-ttu-id="c72a7-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-131">
        - CompressedFile</span></span><br><span data-ttu-id="c72a7-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-132">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-133">
        - File</span></span><br><span data-ttu-id="c72a7-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-134">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-136">
        - Selection</span></span><br><span data-ttu-id="c72a7-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-137">
        - Settings</span></span><br><span data-ttu-id="c72a7-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-138">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-139">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-140">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-142">Office on Windows</span></span><br><span data-ttu-id="c72a7-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-144">- TaskPane</span></span><br><span data-ttu-id="c72a7-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-145">
        - Content</span></span><br><span data-ttu-id="c72a7-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c72a7-146">
        - Custom Functions</span></span><br><span data-ttu-id="c72a7-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="c72a7-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c72a7-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c72a7-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c72a7-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c72a7-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c72a7-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c72a7-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c72a7-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c72a7-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c72a7-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c72a7-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c72a7-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c72a7-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c72a7-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-161">
        - BindingEvents</span></span><br><span data-ttu-id="c72a7-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-162">
        - CompressedFile</span></span><br><span data-ttu-id="c72a7-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-163">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-164">
        - File</span></span><br><span data-ttu-id="c72a7-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-165">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-167">
        - Selection</span></span><br><span data-ttu-id="c72a7-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-168">
        - Settings</span></span><br><span data-ttu-id="c72a7-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-169">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-170">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-171">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-173">Office 2019 on Windows</span></span><br><span data-ttu-id="c72a7-174">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c72a7-175">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-175">- TaskPane</span></span><br><span data-ttu-id="c72a7-176">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-176">
        - Content</span></span><br><span data-ttu-id="c72a7-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c72a7-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c72a7-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c72a7-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c72a7-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c72a7-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c72a7-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c72a7-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c72a7-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c72a7-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c72a7-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-188">- BindingEvents</span></span><br><span data-ttu-id="c72a7-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-189">
        - CompressedFile</span></span><br><span data-ttu-id="c72a7-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-190">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-191">
        - File</span></span><br><span data-ttu-id="c72a7-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-192">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-194">
        - Selection</span></span><br><span data-ttu-id="c72a7-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-195">
        - Settings</span></span><br><span data-ttu-id="c72a7-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-196">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-197">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-198">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-200">Office 2016 on Windows</span></span><br><span data-ttu-id="c72a7-201">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c72a7-202">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-202">- TaskPane</span></span><br><span data-ttu-id="c72a7-203">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-203">
        - Content</span></span></td>
    <td><span data-ttu-id="c72a7-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c72a7-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c72a7-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c72a7-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c72a7-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-207">- BindingEvents</span></span><br><span data-ttu-id="c72a7-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-208">
        - CompressedFile</span></span><br><span data-ttu-id="c72a7-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-209">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-210">
        - File</span></span><br><span data-ttu-id="c72a7-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-211">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-213">
        - Selection</span></span><br><span data-ttu-id="c72a7-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-214">
        - Settings</span></span><br><span data-ttu-id="c72a7-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-215">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-216">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-217">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c72a7-219">Office 2013 on Windows</span></span><br><span data-ttu-id="c72a7-220">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c72a7-221">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-221">
        - TaskPane</span></span><br><span data-ttu-id="c72a7-222">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c72a7-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c72a7-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c72a7-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c72a7-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-225">
        - BindingEvents</span></span><br><span data-ttu-id="c72a7-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-226">
        - CompressedFile</span></span><br><span data-ttu-id="c72a7-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-227">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-228">
        - File</span></span><br><span data-ttu-id="c72a7-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-229">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-231">
        - Selection</span></span><br><span data-ttu-id="c72a7-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-232">
        - Settings</span></span><br><span data-ttu-id="c72a7-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-233">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-234">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-235">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-237">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-237">Office on iPad</span></span><br><span data-ttu-id="c72a7-238">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c72a7-239">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-239">- TaskPane</span></span><br><span data-ttu-id="c72a7-240">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-240">
        - Content</span></span></td>
    <td><span data-ttu-id="c72a7-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c72a7-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c72a7-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c72a7-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c72a7-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c72a7-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c72a7-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c72a7-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c72a7-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c72a7-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c72a7-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c72a7-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-253">- BindingEvents</span></span><br><span data-ttu-id="c72a7-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-254">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-255">
        - File</span></span><br><span data-ttu-id="c72a7-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-256">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-258">
        - Selection</span></span><br><span data-ttu-id="c72a7-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-259">
        - Settings</span></span><br><span data-ttu-id="c72a7-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-260">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-261">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-262">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-264">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-264">Office on Mac</span></span><br><span data-ttu-id="c72a7-265">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c72a7-266">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-266">- TaskPane</span></span><br><span data-ttu-id="c72a7-267">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-267">
        - Content</span></span><br><span data-ttu-id="c72a7-268">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c72a7-268">
        - Custom Functions</span></span><br><span data-ttu-id="c72a7-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c72a7-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c72a7-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c72a7-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c72a7-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c72a7-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c72a7-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c72a7-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c72a7-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c72a7-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c72a7-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c72a7-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c72a7-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c72a7-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-283">- BindingEvents</span></span><br><span data-ttu-id="c72a7-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-284">
        - CompressedFile</span></span><br><span data-ttu-id="c72a7-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-285">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-286">
        - File</span></span><br><span data-ttu-id="c72a7-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-287">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-289">
        - PdfFile</span></span><br><span data-ttu-id="c72a7-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-290">
        - Selection</span></span><br><span data-ttu-id="c72a7-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-291">
        - Settings</span></span><br><span data-ttu-id="c72a7-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-292">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-293">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-294">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-296">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-296">Office 2019 on Mac</span></span><br><span data-ttu-id="c72a7-297">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c72a7-298">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-298">- TaskPane</span></span><br><span data-ttu-id="c72a7-299">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-299">
        - Content</span></span><br><span data-ttu-id="c72a7-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c72a7-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c72a7-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c72a7-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c72a7-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c72a7-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c72a7-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c72a7-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c72a7-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c72a7-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c72a7-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-311">- BindingEvents</span></span><br><span data-ttu-id="c72a7-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-312">
        - CompressedFile</span></span><br><span data-ttu-id="c72a7-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-313">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-314">
        - File</span></span><br><span data-ttu-id="c72a7-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-315">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-317">
        - PdfFile</span></span><br><span data-ttu-id="c72a7-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-318">
        - Selection</span></span><br><span data-ttu-id="c72a7-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-319">
        - Settings</span></span><br><span data-ttu-id="c72a7-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-320">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-321">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-322">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-324">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-324">Office 2016 on Mac</span></span><br><span data-ttu-id="c72a7-325">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c72a7-326">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-326">- TaskPane</span></span><br><span data-ttu-id="c72a7-327">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-327">
        - Content</span></span></td>
    <td><span data-ttu-id="c72a7-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c72a7-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c72a7-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c72a7-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c72a7-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-331">- BindingEvents</span></span><br><span data-ttu-id="c72a7-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-332">
        - CompressedFile</span></span><br><span data-ttu-id="c72a7-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-333">
        - DocumentEvents</span></span><br><span data-ttu-id="c72a7-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-334">
        - File</span></span><br><span data-ttu-id="c72a7-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-335">
        - MatrixBindings</span></span><br><span data-ttu-id="c72a7-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-337">
        - PdfFile</span></span><br><span data-ttu-id="c72a7-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-338">
        - Selection</span></span><br><span data-ttu-id="c72a7-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-339">
        - Settings</span></span><br><span data-ttu-id="c72a7-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-340">
        - TableBindings</span></span><br><span data-ttu-id="c72a7-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-341">
        - TableCoercion</span></span><br><span data-ttu-id="c72a7-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-342">
        - TextBindings</span></span><br><span data-ttu-id="c72a7-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c72a7-344">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="c72a7-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="c72a7-345">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="c72a7-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c72a7-346">平台</span><span class="sxs-lookup"><span data-stu-id="c72a7-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c72a7-347">扩展点</span><span class="sxs-lookup"><span data-stu-id="c72a7-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c72a7-348">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c72a7-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c72a7-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-350">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c72a7-350">Office on the web</span></span></td>
    <td><span data-ttu-id="c72a7-351">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c72a7-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c72a7-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-353">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-353">Office on Windows</span></span><br><span data-ttu-id="c72a7-354">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c72a7-355">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c72a7-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c72a7-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="c72a7-357">Office for Mac</span></span><br><span data-ttu-id="c72a7-358">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="c72a7-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="c72a7-359">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="c72a7-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c72a7-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="c72a7-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="c72a7-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c72a7-362">平台</span><span class="sxs-lookup"><span data-stu-id="c72a7-362">Platform</span></span></th>
    <th><span data-ttu-id="c72a7-363">扩展点</span><span class="sxs-lookup"><span data-stu-id="c72a7-363">Extension points</span></span></th>
    <th><span data-ttu-id="c72a7-364">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="c72a7-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c72a7-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-366">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c72a7-366">Office on the web</span></span><br><span data-ttu-id="c72a7-367">（新式）</span><span class="sxs-lookup"><span data-stu-id="c72a7-367">(modern)</span></span></td>
    <td> <span data-ttu-id="c72a7-368">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-368">- Message Read</span></span><br><span data-ttu-id="c72a7-369">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-369">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-370">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-370">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-371">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-371">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="c72a7-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c72a7-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c72a7-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c72a7-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c72a7-381">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-382">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c72a7-382">Office on the web</span></span><br><span data-ttu-id="c72a7-383">（经典）</span><span class="sxs-lookup"><span data-stu-id="c72a7-383">(classic)</span></span></td>
    <td> <span data-ttu-id="c72a7-384">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-384">- Message Read</span></span><br><span data-ttu-id="c72a7-385">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-385">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-386">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-386">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-387">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-387">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="c72a7-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c72a7-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c72a7-395">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-396">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-396">Office on Windows</span></span><br><span data-ttu-id="c72a7-397">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-398">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-398">- Message Read</span></span><br><span data-ttu-id="c72a7-399">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-399">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-400">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-400">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-401">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-401">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="c72a7-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c72a7-403">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="c72a7-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c72a7-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c72a7-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c72a7-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c72a7-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c72a7-412">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-413">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-413">Office 2019 on Windows</span></span><br><span data-ttu-id="c72a7-414">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-415">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-415">- Message Read</span></span><br><span data-ttu-id="c72a7-416">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-416">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-417">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-417">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-418">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-418">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="c72a7-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c72a7-420">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="c72a7-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c72a7-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c72a7-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c72a7-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c72a7-428">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-429">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-429">Office 2016 on Windows</span></span><br><span data-ttu-id="c72a7-430">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-431">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-431">- Message Read</span></span><br><span data-ttu-id="c72a7-432">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-432">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-433">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-433">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-434">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-434">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="c72a7-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c72a7-436">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="c72a7-436">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c72a7-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c72a7-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c72a7-441">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-442">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c72a7-442">Office 2013 on Windows</span></span><br><span data-ttu-id="c72a7-443">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-444">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-444">- Message Read</span></span><br><span data-ttu-id="c72a7-445">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-445">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-446">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-446">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-447">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-447">
      - Appointment Organizer (Compose)</span></span><br>
    <td> <span data-ttu-id="c72a7-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="c72a7-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="c72a7-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c72a7-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c72a7-452">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-453">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-453">Office on iOS</span></span><br><span data-ttu-id="c72a7-454">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-455">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-455">- Message Read</span></span><br><span data-ttu-id="c72a7-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c72a7-462">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-463">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-463">Office on Mac</span></span><br><span data-ttu-id="c72a7-464">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-465">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-465">- Message Read</span></span><br><span data-ttu-id="c72a7-466">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-466">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-467">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-467">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-468">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-468">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="c72a7-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c72a7-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c72a7-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c72a7-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c72a7-478">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-479">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-479">Office 2019 on Mac</span></span><br><span data-ttu-id="c72a7-480">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-481">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-481">- Message Read</span></span><br><span data-ttu-id="c72a7-482">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-482">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-483">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-483">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-484">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-484">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="c72a7-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c72a7-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c72a7-492">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-493">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-493">Office 2016 on Mac</span></span><br><span data-ttu-id="c72a7-494">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-495">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-495">- Message Read</span></span><br><span data-ttu-id="c72a7-496">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="c72a7-496">
      - Message Compose</span></span><br><span data-ttu-id="c72a7-497">
      -约会参与者（阅读）</span><span class="sxs-lookup"><span data-stu-id="c72a7-497">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="c72a7-498">
      -约会参与者（撰写）</span><span class="sxs-lookup"><span data-stu-id="c72a7-498">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="c72a7-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c72a7-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c72a7-506">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-507">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-507">Office on Android</span></span><br><span data-ttu-id="c72a7-508">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-509">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="c72a7-509">- Message Read</span></span><br><span data-ttu-id="c72a7-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c72a7-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c72a7-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c72a7-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c72a7-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c72a7-516">不可用</span><span class="sxs-lookup"><span data-stu-id="c72a7-516">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="c72a7-517">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="c72a7-517">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c72a7-518">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="c72a7-518">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="c72a7-519">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="c72a7-519">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="c72a7-520">Word</span><span class="sxs-lookup"><span data-stu-id="c72a7-520">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c72a7-521">平台</span><span class="sxs-lookup"><span data-stu-id="c72a7-521">Platform</span></span></th>
    <th><span data-ttu-id="c72a7-522">扩展点</span><span class="sxs-lookup"><span data-stu-id="c72a7-522">Extension points</span></span></th>
    <th><span data-ttu-id="c72a7-523">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-523">API requirement sets</span></span></th>
    <th><span data-ttu-id="c72a7-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c72a7-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-525">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c72a7-525">Office on the web</span></span></td>
    <td> <span data-ttu-id="c72a7-526">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-526">- TaskPane</span></span><br><span data-ttu-id="c72a7-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c72a7-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c72a7-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c72a7-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c72a7-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c72a7-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-534">- BindingEvents</span></span><br><span data-ttu-id="c72a7-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-536">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-537">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-537">
         - File</span></span><br><span data-ttu-id="c72a7-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-539">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-542">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-543">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-543">
         - Selection</span></span><br><span data-ttu-id="c72a7-544">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-544">
         - Settings</span></span><br><span data-ttu-id="c72a7-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-545">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-546">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-547">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-548">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-549">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-550">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-550">Office on Windows</span></span><br><span data-ttu-id="c72a7-551">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-551">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-552">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-552">- TaskPane</span></span><br><span data-ttu-id="c72a7-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c72a7-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c72a7-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c72a7-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c72a7-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c72a7-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-560">- BindingEvents</span></span><br><span data-ttu-id="c72a7-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-561">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-563">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-564">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-564">
         - File</span></span><br><span data-ttu-id="c72a7-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-566">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-569">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-570">
         - Selection</span></span><br><span data-ttu-id="c72a7-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-571">
         - Settings</span></span><br><span data-ttu-id="c72a7-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-572">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-573">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-574">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-575">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-577">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-577">Office 2019 on Windows</span></span><br><span data-ttu-id="c72a7-578">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-579">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-579">- TaskPane</span></span><br><span data-ttu-id="c72a7-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c72a7-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c72a7-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c72a7-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-586">- BindingEvents</span></span><br><span data-ttu-id="c72a7-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-587">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-589">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-590">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-590">
         - File</span></span><br><span data-ttu-id="c72a7-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-592">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-595">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-596">
         - Selection</span></span><br><span data-ttu-id="c72a7-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-597">
         - Settings</span></span><br><span data-ttu-id="c72a7-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-598">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-599">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-600">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-601">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-603">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-603">Office 2016 on Windows</span></span><br><span data-ttu-id="c72a7-604">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-605">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c72a7-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c72a7-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c72a7-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-609">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-609">- BindingEvents</span></span><br><span data-ttu-id="c72a7-610">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-610">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-611">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-611">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-612">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-613">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-613">
         - File</span></span><br><span data-ttu-id="c72a7-614">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-614">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-615">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-615">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-616">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-616">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-617">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-617">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-618">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-618">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-619">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-619">
         - Selection</span></span><br><span data-ttu-id="c72a7-620">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-620">
         - Settings</span></span><br><span data-ttu-id="c72a7-621">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-621">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-622">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-622">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-623">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-623">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-624">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-625">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-625">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-626">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c72a7-626">Office 2013 on Windows</span></span><br><span data-ttu-id="c72a7-627">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-627">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-628">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-628">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c72a7-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c72a7-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-631">- BindingEvents</span></span><br><span data-ttu-id="c72a7-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-632">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-634">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-635">
         - File</span></span><br><span data-ttu-id="c72a7-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-637">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-640">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-641">
         - Selection</span></span><br><span data-ttu-id="c72a7-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-642">
         - Settings</span></span><br><span data-ttu-id="c72a7-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-643">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-644">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-645">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-646">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-647">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-648">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-648">Office on iPad</span></span><br><span data-ttu-id="c72a7-649">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-650">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-650">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c72a7-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c72a7-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c72a7-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c72a7-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-656">- BindingEvents</span></span><br><span data-ttu-id="c72a7-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-657">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-659">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-660">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-660">
         - File</span></span><br><span data-ttu-id="c72a7-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-662">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-665">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-666">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-666">
         - Selection</span></span><br><span data-ttu-id="c72a7-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-667">
         - Settings</span></span><br><span data-ttu-id="c72a7-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-668">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-669">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-670">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-671">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-673">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-673">Office on Mac</span></span><br><span data-ttu-id="c72a7-674">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-674">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-675">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-675">- TaskPane</span></span><br><span data-ttu-id="c72a7-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c72a7-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c72a7-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c72a7-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c72a7-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="c72a7-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-683">- BindingEvents</span></span><br><span data-ttu-id="c72a7-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-684">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-686">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-687">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-687">
         - File</span></span><br><span data-ttu-id="c72a7-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-689">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-692">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-693">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-693">
         - Selection</span></span><br><span data-ttu-id="c72a7-694">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-694">
         - Settings</span></span><br><span data-ttu-id="c72a7-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-695">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-696">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-697">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-698">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-700">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-700">Office 2019 on Mac</span></span><br><span data-ttu-id="c72a7-701">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-702">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-702">- TaskPane</span></span><br><span data-ttu-id="c72a7-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c72a7-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c72a7-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c72a7-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c72a7-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-709">- BindingEvents</span></span><br><span data-ttu-id="c72a7-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-710">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-712">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-713">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-713">
         - File</span></span><br><span data-ttu-id="c72a7-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-715">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-718">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-719">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-719">
         - Selection</span></span><br><span data-ttu-id="c72a7-720">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-720">
         - Settings</span></span><br><span data-ttu-id="c72a7-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-721">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-722">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-723">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-724">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-725">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-726">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-726">Office 2016 on Mac</span></span><br><span data-ttu-id="c72a7-727">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-727">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-728">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-728">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c72a7-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c72a7-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c72a7-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-732">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-732">- BindingEvents</span></span><br><span data-ttu-id="c72a7-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-733">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-734">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c72a7-734">
         - CustomXmlParts</span></span><br><span data-ttu-id="c72a7-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-735">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-736">
         - File</span></span><br><span data-ttu-id="c72a7-737">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-737">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-738">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-738">
         - MatrixBindings</span></span><br><span data-ttu-id="c72a7-739">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-739">
         - MatrixCoercion</span></span><br><span data-ttu-id="c72a7-740">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-740">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c72a7-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-741">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-742">
         - Selection</span></span><br><span data-ttu-id="c72a7-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-743">
         - Settings</span></span><br><span data-ttu-id="c72a7-744">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-744">
         - TableBindings</span></span><br><span data-ttu-id="c72a7-745">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-745">
         - TableCoercion</span></span><br><span data-ttu-id="c72a7-746">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c72a7-746">
         - TextBindings</span></span><br><span data-ttu-id="c72a7-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-747">
         - TextCoercion</span></span><br><span data-ttu-id="c72a7-748">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-748">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="c72a7-749">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="c72a7-749">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c72a7-750">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c72a7-750">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c72a7-751">平台</span><span class="sxs-lookup"><span data-stu-id="c72a7-751">Platform</span></span></th>
    <th><span data-ttu-id="c72a7-752">扩展点</span><span class="sxs-lookup"><span data-stu-id="c72a7-752">Extension points</span></span></th>
    <th><span data-ttu-id="c72a7-753">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-753">API requirement sets</span></span></th>
    <th><span data-ttu-id="c72a7-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c72a7-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-755">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c72a7-755">Office on the web</span></span></td>
    <td> <span data-ttu-id="c72a7-756">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-756">- Content</span></span><br><span data-ttu-id="c72a7-757">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-757">
         - TaskPane</span></span><br><span data-ttu-id="c72a7-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c72a7-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c72a7-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c72a7-763">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-763">- ActiveView</span></span><br><span data-ttu-id="c72a7-764">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-764">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-765">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-765">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-766">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-766">
         - File</span></span><br><span data-ttu-id="c72a7-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-767">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-768">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-768">
         - Selection</span></span><br><span data-ttu-id="c72a7-769">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-769">
         - Settings</span></span><br><span data-ttu-id="c72a7-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-771">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-771">Office on Windows</span></span><br><span data-ttu-id="c72a7-772">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-772">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-773">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-773">- Content</span></span><br><span data-ttu-id="c72a7-774">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-774">
         - TaskPane</span></span><br><span data-ttu-id="c72a7-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c72a7-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c72a7-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c72a7-780">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-780">- ActiveView</span></span><br><span data-ttu-id="c72a7-781">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-781">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-782">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-782">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-783">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-783">
         - File</span></span><br><span data-ttu-id="c72a7-784">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-784">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-785">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-785">
         - Selection</span></span><br><span data-ttu-id="c72a7-786">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-786">
         - Settings</span></span><br><span data-ttu-id="c72a7-787">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-787">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-788">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-788">Office 2019 on Windows</span></span><br><span data-ttu-id="c72a7-789">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-789">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-790">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-790">- Content</span></span><br><span data-ttu-id="c72a7-791">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-791">
         - TaskPane</span></span><br><span data-ttu-id="c72a7-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-795">- ActiveView</span></span><br><span data-ttu-id="c72a7-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-796">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-797">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-798">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-798">
         - File</span></span><br><span data-ttu-id="c72a7-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-799">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-800">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-800">
         - Selection</span></span><br><span data-ttu-id="c72a7-801">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-801">
         - Settings</span></span><br><span data-ttu-id="c72a7-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-803">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-803">Office 2016 on Windows</span></span><br><span data-ttu-id="c72a7-804">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-804">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-805">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-805">- Content</span></span><br><span data-ttu-id="c72a7-806">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c72a7-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c72a7-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-809">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-809">- ActiveView</span></span><br><span data-ttu-id="c72a7-810">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-810">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-811">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-811">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-812">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-812">
         - File</span></span><br><span data-ttu-id="c72a7-813">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-813">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-814">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-814">
         - Selection</span></span><br><span data-ttu-id="c72a7-815">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-815">
         - Settings</span></span><br><span data-ttu-id="c72a7-816">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-816">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-817">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c72a7-817">Office 2013 on Windows</span></span><br><span data-ttu-id="c72a7-818">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-818">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-819">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-819">- Content</span></span><br><span data-ttu-id="c72a7-820">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-820">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="c72a7-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c72a7-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c72a7-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-823">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-823">- ActiveView</span></span><br><span data-ttu-id="c72a7-824">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-824">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-825">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-825">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-826">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-826">
         - File</span></span><br><span data-ttu-id="c72a7-827">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-827">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-828">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-828">
         - Selection</span></span><br><span data-ttu-id="c72a7-829">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-829">
         - Settings</span></span><br><span data-ttu-id="c72a7-830">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-830">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-831">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-831">Office on iPad</span></span><br><span data-ttu-id="c72a7-832">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-832">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-833">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-833">- Content</span></span><br><span data-ttu-id="c72a7-834">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-834">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c72a7-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-838">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-838">- ActiveView</span></span><br><span data-ttu-id="c72a7-839">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-839">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-840">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-840">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-841">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-841">
         - File</span></span><br><span data-ttu-id="c72a7-842">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-842">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-843">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-843">
         - Selection</span></span><br><span data-ttu-id="c72a7-844">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-844">
         - Settings</span></span><br><span data-ttu-id="c72a7-845">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-845">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-846">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="c72a7-846">Office on Mac</span></span><br><span data-ttu-id="c72a7-847">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="c72a7-847">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c72a7-848">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-848">- Content</span></span><br><span data-ttu-id="c72a7-849">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-849">
         - TaskPane</span></span><br><span data-ttu-id="c72a7-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c72a7-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c72a7-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c72a7-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-855">- ActiveView</span></span><br><span data-ttu-id="c72a7-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-856">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-857">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-858">
         - File</span></span><br><span data-ttu-id="c72a7-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-859">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-860">
         - Selection</span></span><br><span data-ttu-id="c72a7-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-861">
         - Settings</span></span><br><span data-ttu-id="c72a7-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-862">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-863">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-863">Office 2019 on Mac</span></span><br><span data-ttu-id="c72a7-864">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-864">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-865">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-865">- Content</span></span><br><span data-ttu-id="c72a7-866">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-866">
         - TaskPane</span></span><br><span data-ttu-id="c72a7-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-870">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-870">- ActiveView</span></span><br><span data-ttu-id="c72a7-871">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-871">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-872">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-872">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-873">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-873">
         - File</span></span><br><span data-ttu-id="c72a7-874">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-874">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-875">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-875">
         - Selection</span></span><br><span data-ttu-id="c72a7-876">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-876">
         - Settings</span></span><br><span data-ttu-id="c72a7-877">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-877">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-878">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-878">Office 2016 on Mac</span></span><br><span data-ttu-id="c72a7-879">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-879">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-880">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-880">- Content</span></span><br><span data-ttu-id="c72a7-881">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-881">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c72a7-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c72a7-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-884">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c72a7-884">- ActiveView</span></span><br><span data-ttu-id="c72a7-885">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-885">
         - CompressedFile</span></span><br><span data-ttu-id="c72a7-886">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-886">
         - DocumentEvents</span></span><br><span data-ttu-id="c72a7-887">
         - File</span><span class="sxs-lookup"><span data-stu-id="c72a7-887">
         - File</span></span><br><span data-ttu-id="c72a7-888">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c72a7-888">
         - PdfFile</span></span><br><span data-ttu-id="c72a7-889">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-889">
         - Selection</span></span><br><span data-ttu-id="c72a7-890">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-890">
         - Settings</span></span><br><span data-ttu-id="c72a7-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c72a7-892">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="c72a7-892">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="c72a7-893">OneNote</span><span class="sxs-lookup"><span data-stu-id="c72a7-893">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c72a7-894">平台</span><span class="sxs-lookup"><span data-stu-id="c72a7-894">Platform</span></span></th>
    <th><span data-ttu-id="c72a7-895">扩展点</span><span class="sxs-lookup"><span data-stu-id="c72a7-895">Extension points</span></span></th>
    <th><span data-ttu-id="c72a7-896">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-896">API requirement sets</span></span></th>
    <th><span data-ttu-id="c72a7-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c72a7-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-898">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="c72a7-898">Office on the web</span></span></td>
    <td> <span data-ttu-id="c72a7-899">- 内容</span><span class="sxs-lookup"><span data-stu-id="c72a7-899">- Content</span></span><br><span data-ttu-id="c72a7-900">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-900">
         - TaskPane</span></span><br><span data-ttu-id="c72a7-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c72a7-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c72a7-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c72a7-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-905">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c72a7-905">- DocumentEvents</span></span><br><span data-ttu-id="c72a7-906">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-906">
         - HtmlCoercion</span></span><br><span data-ttu-id="c72a7-907">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c72a7-907">
         - Settings</span></span><br><span data-ttu-id="c72a7-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="c72a7-909">项目</span><span class="sxs-lookup"><span data-stu-id="c72a7-909">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c72a7-910">平台</span><span class="sxs-lookup"><span data-stu-id="c72a7-910">Platform</span></span></th>
    <th><span data-ttu-id="c72a7-911">扩展点</span><span class="sxs-lookup"><span data-stu-id="c72a7-911">Extension points</span></span></th>
    <th><span data-ttu-id="c72a7-912">API 要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-912">API requirement sets</span></span></th>
    <th><span data-ttu-id="c72a7-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c72a7-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-914">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c72a7-914">Office 2019 on Windows</span></span><br><span data-ttu-id="c72a7-915">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-915">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-916">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-916">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-918">- Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-918">- Selection</span></span><br><span data-ttu-id="c72a7-919">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-919">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-920">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c72a7-920">Office 2016 on Windows</span></span><br><span data-ttu-id="c72a7-921">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-921">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-922">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-922">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-924">- Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-924">- Selection</span></span><br><span data-ttu-id="c72a7-925">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-925">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c72a7-926">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c72a7-926">Office 2013 on Windows</span></span><br><span data-ttu-id="c72a7-927">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="c72a7-927">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c72a7-928">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c72a7-928">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c72a7-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c72a7-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c72a7-930">- Selection</span><span class="sxs-lookup"><span data-stu-id="c72a7-930">- Selection</span></span><br><span data-ttu-id="c72a7-931">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c72a7-931">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c72a7-932">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c72a7-932">See also</span></span>

- [<span data-ttu-id="c72a7-933">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="c72a7-933">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c72a7-934">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-934">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="c72a7-935">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-935">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="c72a7-936">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="c72a7-936">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="c72a7-937">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="c72a7-937">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="c72a7-938">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="c72a7-938">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="c72a7-939">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="c72a7-939">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="c72a7-940">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="c72a7-940">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="c72a7-941">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c72a7-941">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="c72a7-942">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c72a7-942">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="c72a7-943">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="c72a7-943">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="c72a7-944">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="c72a7-944">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)