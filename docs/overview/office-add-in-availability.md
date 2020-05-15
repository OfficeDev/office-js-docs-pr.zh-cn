---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 36c6bc6b6348ac988049f9a50127f6dd2f94bf37
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217821"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="ffdb1-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="ffdb1-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="ffdb1-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="ffdb1-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="ffdb1-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="ffdb1-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="ffdb1-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="ffdb1-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="ffdb1-108">Excel</span><span class="sxs-lookup"><span data-stu-id="ffdb1-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ffdb1-109">平台</span><span class="sxs-lookup"><span data-stu-id="ffdb1-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ffdb1-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="ffdb1-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ffdb1-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ffdb1-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ffdb1-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="ffdb1-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-114">- TaskPane</span></span><br><span data-ttu-id="ffdb1-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-115">
        - Content</span></span><br><span data-ttu-id="ffdb1-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ffdb1-116">
        - Custom Functions</span></span><br><span data-ttu-id="ffdb1-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="ffdb1-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ffdb1-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ffdb1-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ffdb1-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ffdb1-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ffdb1-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ffdb1-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ffdb1-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ffdb1-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ffdb1-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="ffdb1-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ffdb1-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-131">
        - BindingEvents</span></span><br><span data-ttu-id="ffdb1-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-132">
        - CompressedFile</span></span><br><span data-ttu-id="ffdb1-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-133">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-134">
        - File</span></span><br><span data-ttu-id="ffdb1-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-135">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-137">
        - Selection</span></span><br><span data-ttu-id="ffdb1-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-138">
        - Settings</span></span><br><span data-ttu-id="ffdb1-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-139">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-140">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-141">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-143">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-143">Office on Windows</span></span><br><span data-ttu-id="ffdb1-144">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-145">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-145">- TaskPane</span></span><br><span data-ttu-id="ffdb1-146">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-146">
        - Content</span></span><br><span data-ttu-id="ffdb1-147">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ffdb1-147">
        - Custom Functions</span></span><br><span data-ttu-id="ffdb1-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="ffdb1-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ffdb1-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ffdb1-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ffdb1-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ffdb1-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ffdb1-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ffdb1-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ffdb1-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ffdb1-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ffdb1-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ffdb1-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ffdb1-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-163">
        - BindingEvents</span></span><br><span data-ttu-id="ffdb1-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-164">
        - CompressedFile</span></span><br><span data-ttu-id="ffdb1-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-165">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-166">
        - File</span></span><br><span data-ttu-id="ffdb1-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-167">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-169">
        - Selection</span></span><br><span data-ttu-id="ffdb1-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-170">
        - Settings</span></span><br><span data-ttu-id="ffdb1-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-171">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-172">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-173">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-175">Office 2019 on Windows</span></span><br><span data-ttu-id="ffdb1-176">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ffdb1-177">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-177">- TaskPane</span></span><br><span data-ttu-id="ffdb1-178">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-178">
        - Content</span></span><br><span data-ttu-id="ffdb1-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ffdb1-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ffdb1-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ffdb1-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ffdb1-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ffdb1-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ffdb1-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ffdb1-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-190">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-191">
        - CompressedFile</span></span><br><span data-ttu-id="ffdb1-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-192">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-193">
        - File</span></span><br><span data-ttu-id="ffdb1-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-194">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-196">
        - Selection</span></span><br><span data-ttu-id="ffdb1-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-197">
        - Settings</span></span><br><span data-ttu-id="ffdb1-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-198">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-199">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-200">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-202">Office 2016 on Windows</span></span><br><span data-ttu-id="ffdb1-203">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ffdb1-204">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-204">- TaskPane</span></span><br><span data-ttu-id="ffdb1-205">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-205">
        - Content</span></span></td>
    <td><span data-ttu-id="ffdb1-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ffdb1-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ffdb1-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-209">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-210">
        - CompressedFile</span></span><br><span data-ttu-id="ffdb1-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-211">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-212">
        - File</span></span><br><span data-ttu-id="ffdb1-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-213">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-215">
        - Selection</span></span><br><span data-ttu-id="ffdb1-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-216">
        - Settings</span></span><br><span data-ttu-id="ffdb1-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-217">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-218">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-219">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ffdb1-221">Office 2013 on Windows</span></span><br><span data-ttu-id="ffdb1-222">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ffdb1-223">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-223">
        - TaskPane</span></span><br><span data-ttu-id="ffdb1-224">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="ffdb1-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ffdb1-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ffdb1-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-227">
        - BindingEvents</span></span><br><span data-ttu-id="ffdb1-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-228">
        - CompressedFile</span></span><br><span data-ttu-id="ffdb1-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-229">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-230">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-230">
        - File</span></span><br><span data-ttu-id="ffdb1-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-231">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-233">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-233">
        - Selection</span></span><br><span data-ttu-id="ffdb1-234">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-234">
        - Settings</span></span><br><span data-ttu-id="ffdb1-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-235">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-236">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-237">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-239">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-239">Office on iPad</span></span><br><span data-ttu-id="ffdb1-240">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ffdb1-241">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-241">- TaskPane</span></span><br><span data-ttu-id="ffdb1-242">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-242">
        - Content</span></span></td>
    <td><span data-ttu-id="ffdb1-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ffdb1-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ffdb1-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ffdb1-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ffdb1-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ffdb1-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ffdb1-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ffdb1-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ffdb1-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ffdb1-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-256">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-257">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-258">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-258">
        - File</span></span><br><span data-ttu-id="ffdb1-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-259">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-261">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-261">
        - Selection</span></span><br><span data-ttu-id="ffdb1-262">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-262">
        - Settings</span></span><br><span data-ttu-id="ffdb1-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-263">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-264">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-265">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-267">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-267">Office on Mac</span></span><br><span data-ttu-id="ffdb1-268">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ffdb1-269">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-269">- TaskPane</span></span><br><span data-ttu-id="ffdb1-270">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-270">
        - Content</span></span><br><span data-ttu-id="ffdb1-271">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ffdb1-271">
        - Custom Functions</span></span><br><span data-ttu-id="ffdb1-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ffdb1-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ffdb1-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ffdb1-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ffdb1-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ffdb1-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ffdb1-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ffdb1-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ffdb1-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ffdb1-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ffdb1-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ffdb1-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-287">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-288">
        - CompressedFile</span></span><br><span data-ttu-id="ffdb1-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-289">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-290">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-290">
        - File</span></span><br><span data-ttu-id="ffdb1-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-291">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-293">
        - PdfFile</span></span><br><span data-ttu-id="ffdb1-294">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-294">
        - Selection</span></span><br><span data-ttu-id="ffdb1-295">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-295">
        - Settings</span></span><br><span data-ttu-id="ffdb1-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-296">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-297">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-298">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-300">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-300">Office 2019 on Mac</span></span><br><span data-ttu-id="ffdb1-301">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ffdb1-302">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-302">- TaskPane</span></span><br><span data-ttu-id="ffdb1-303">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-303">
        - Content</span></span><br><span data-ttu-id="ffdb1-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ffdb1-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ffdb1-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ffdb1-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ffdb1-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ffdb1-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ffdb1-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ffdb1-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-315">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-316">
        - CompressedFile</span></span><br><span data-ttu-id="ffdb1-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-317">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-318">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-318">
        - File</span></span><br><span data-ttu-id="ffdb1-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-319">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-321">
        - PdfFile</span></span><br><span data-ttu-id="ffdb1-322">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-322">
        - Selection</span></span><br><span data-ttu-id="ffdb1-323">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-323">
        - Settings</span></span><br><span data-ttu-id="ffdb1-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-324">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-325">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-326">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-328">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-328">Office 2016 on Mac</span></span><br><span data-ttu-id="ffdb1-329">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ffdb1-330">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-330">- TaskPane</span></span><br><span data-ttu-id="ffdb1-331">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-331">
        - Content</span></span></td>
    <td><span data-ttu-id="ffdb1-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ffdb1-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ffdb1-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-335">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-336">
        - CompressedFile</span></span><br><span data-ttu-id="ffdb1-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-337">
        - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-338">
        - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-338">
        - File</span></span><br><span data-ttu-id="ffdb1-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-339">
        - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-341">
        - PdfFile</span></span><br><span data-ttu-id="ffdb1-342">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-342">
        - Selection</span></span><br><span data-ttu-id="ffdb1-343">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-343">
        - Settings</span></span><br><span data-ttu-id="ffdb1-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-344">
        - TableBindings</span></span><br><span data-ttu-id="ffdb1-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-345">
        - TableCoercion</span></span><br><span data-ttu-id="ffdb1-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-346">
        - TextBindings</span></span><br><span data-ttu-id="ffdb1-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ffdb1-348">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="ffdb1-349">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ffdb1-350">平台</span><span class="sxs-lookup"><span data-stu-id="ffdb1-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ffdb1-351">扩展点</span><span class="sxs-lookup"><span data-stu-id="ffdb1-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ffdb1-352">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ffdb1-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-354">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ffdb1-354">Office on the web</span></span></td>
    <td><span data-ttu-id="ffdb1-355">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ffdb1-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ffdb1-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-357">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-357">Office on Windows</span></span><br><span data-ttu-id="ffdb1-358">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ffdb1-359">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ffdb1-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ffdb1-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-361">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="ffdb1-361">Office for Mac</span></span><br><span data-ttu-id="ffdb1-362">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="ffdb1-363">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ffdb1-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ffdb1-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="ffdb1-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="ffdb1-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ffdb1-366">平台</span><span class="sxs-lookup"><span data-stu-id="ffdb1-366">Platform</span></span></th>
    <th><span data-ttu-id="ffdb1-367">扩展点</span><span class="sxs-lookup"><span data-stu-id="ffdb1-367">Extension points</span></span></th>
    <th><span data-ttu-id="ffdb1-368">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="ffdb1-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-370">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ffdb1-370">Office on the web</span></span><br><span data-ttu-id="ffdb1-371">（新式）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-371">(modern)</span></span></td>
    <td> <span data-ttu-id="ffdb1-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ffdb1-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ffdb1-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ffdb1-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ffdb1-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ffdb1-385">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-386">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ffdb1-386">Office on the web</span></span><br><span data-ttu-id="ffdb1-387">（经典）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-387">(classic)</span></span></td>
    <td> <span data-ttu-id="ffdb1-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ffdb1-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ffdb1-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ffdb1-399">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-400">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-400">Office on Windows</span></span><br><span data-ttu-id="ffdb1-401">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ffdb1-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ffdb1-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ffdb1-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ffdb1-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ffdb1-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ffdb1-416">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-417">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-417">Office 2019 on Windows</span></span><br><span data-ttu-id="ffdb1-418">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ffdb1-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ffdb1-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ffdb1-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ffdb1-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ffdb1-432">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-433">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-433">Office 2016 on Windows</span></span><br><span data-ttu-id="ffdb1-434">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ffdb1-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ffdb1-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ffdb1-445">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-446">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ffdb1-446">Office 2013 on Windows</span></span><br><span data-ttu-id="ffdb1-447">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="ffdb1-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="ffdb1-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ffdb1-456">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-457">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-457">Office on iOS</span></span><br><span data-ttu-id="ffdb1-458">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ffdb1-466">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-467">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-467">Office on Mac</span></span><br><span data-ttu-id="ffdb1-468">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ffdb1-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ffdb1-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ffdb1-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ffdb1-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ffdb1-482">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-483">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-483">Office 2019 on Mac</span></span><br><span data-ttu-id="ffdb1-484">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ffdb1-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ffdb1-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ffdb1-496">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-497">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-497">Office 2016 on Mac</span></span><br><span data-ttu-id="ffdb1-498">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ffdb1-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ffdb1-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ffdb1-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ffdb1-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ffdb1-510">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-511">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-511">Office on Android</span></span><br><span data-ttu-id="ffdb1-512">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ffdb1-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">约会组织者（撰写）：联机会议</a> （预览）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="ffdb1-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ffdb1-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ffdb1-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ffdb1-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ffdb1-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ffdb1-521">不可用</span><span class="sxs-lookup"><span data-stu-id="ffdb1-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="ffdb1-522">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ffdb1-523">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="ffdb1-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="ffdb1-524">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="ffdb1-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="ffdb1-525">Word</span><span class="sxs-lookup"><span data-stu-id="ffdb1-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ffdb1-526">平台</span><span class="sxs-lookup"><span data-stu-id="ffdb1-526">Platform</span></span></th>
    <th><span data-ttu-id="ffdb1-527">扩展点</span><span class="sxs-lookup"><span data-stu-id="ffdb1-527">Extension points</span></span></th>
    <th><span data-ttu-id="ffdb1-528">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="ffdb1-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-530">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ffdb1-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="ffdb1-531">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-531">- TaskPane</span></span><br><span data-ttu-id="ffdb1-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ffdb1-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-539">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-541">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-542">
         - File</span></span><br><span data-ttu-id="ffdb1-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-544">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-547">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-548">
         - Selection</span></span><br><span data-ttu-id="ffdb1-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-549">
         - Settings</span></span><br><span data-ttu-id="ffdb1-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-550">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-551">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-552">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-553">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-555">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-555">Office on Windows</span></span><br><span data-ttu-id="ffdb1-556">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-557">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-557">- TaskPane</span></span><br><span data-ttu-id="ffdb1-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ffdb1-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-565">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-566">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-568">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-569">
         - File</span></span><br><span data-ttu-id="ffdb1-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-571">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-574">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-575">
         - Selection</span></span><br><span data-ttu-id="ffdb1-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-576">
         - Settings</span></span><br><span data-ttu-id="ffdb1-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-577">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-578">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-579">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-580">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-582">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-582">Office 2019 on Windows</span></span><br><span data-ttu-id="ffdb1-583">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-584">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-584">- TaskPane</span></span><br><span data-ttu-id="ffdb1-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-591">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-592">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-594">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-595">
         - File</span></span><br><span data-ttu-id="ffdb1-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-597">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-600">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-601">
         - Selection</span></span><br><span data-ttu-id="ffdb1-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-602">
         - Settings</span></span><br><span data-ttu-id="ffdb1-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-603">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-604">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-605">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-606">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-608">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-608">Office 2016 on Windows</span></span><br><span data-ttu-id="ffdb1-609">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-610">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ffdb1-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-614">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-615">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-617">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-618">
         - File</span></span><br><span data-ttu-id="ffdb1-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-620">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-623">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-624">
         - Selection</span></span><br><span data-ttu-id="ffdb1-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-625">
         - Settings</span></span><br><span data-ttu-id="ffdb1-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-626">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-627">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-628">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-629">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-631">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ffdb1-631">Office 2013 on Windows</span></span><br><span data-ttu-id="ffdb1-632">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-633">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ffdb1-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-636">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-637">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-639">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-640">
         - File</span></span><br><span data-ttu-id="ffdb1-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-642">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-645">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-646">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-646">
         - Selection</span></span><br><span data-ttu-id="ffdb1-647">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-647">
         - Settings</span></span><br><span data-ttu-id="ffdb1-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-648">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-649">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-650">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-651">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-653">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-653">Office on iPad</span></span><br><span data-ttu-id="ffdb1-654">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-655">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ffdb1-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-661">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-662">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-664">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-665">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-665">
         - File</span></span><br><span data-ttu-id="ffdb1-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-667">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-670">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-671">
         - Selection</span></span><br><span data-ttu-id="ffdb1-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-672">
         - Settings</span></span><br><span data-ttu-id="ffdb1-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-673">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-674">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-675">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-676">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-678">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-678">Office on Mac</span></span><br><span data-ttu-id="ffdb1-679">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-680">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-680">- TaskPane</span></span><br><span data-ttu-id="ffdb1-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ffdb1-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="ffdb1-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-688">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-689">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-691">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-692">
         - File</span></span><br><span data-ttu-id="ffdb1-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-694">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-697">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-698">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-698">
         - Selection</span></span><br><span data-ttu-id="ffdb1-699">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-699">
         - Settings</span></span><br><span data-ttu-id="ffdb1-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-700">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-701">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-702">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-703">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-705">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-705">Office 2019 on Mac</span></span><br><span data-ttu-id="ffdb1-706">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-707">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-707">- TaskPane</span></span><br><span data-ttu-id="ffdb1-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ffdb1-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ffdb1-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ffdb1-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-714">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-715">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-717">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-718">
         - File</span></span><br><span data-ttu-id="ffdb1-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-720">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-723">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-724">
         - Selection</span></span><br><span data-ttu-id="ffdb1-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-725">
         - Settings</span></span><br><span data-ttu-id="ffdb1-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-726">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-727">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-728">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-729">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-731">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-731">Office 2016 on Mac</span></span><br><span data-ttu-id="ffdb1-732">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-733">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ffdb1-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-737">- BindingEvents</span></span><br><span data-ttu-id="ffdb1-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-738">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ffdb1-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="ffdb1-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-740">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-741">
         - File</span></span><br><span data-ttu-id="ffdb1-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-743">
         - MatrixBindings</span></span><br><span data-ttu-id="ffdb1-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="ffdb1-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ffdb1-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-746">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-747">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-747">
         - Selection</span></span><br><span data-ttu-id="ffdb1-748">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-748">
         - Settings</span></span><br><span data-ttu-id="ffdb1-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-749">
         - TableBindings</span></span><br><span data-ttu-id="ffdb1-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-750">
         - TableCoercion</span></span><br><span data-ttu-id="ffdb1-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-751">
         - TextBindings</span></span><br><span data-ttu-id="ffdb1-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-752">
         - TextCoercion</span></span><br><span data-ttu-id="ffdb1-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="ffdb1-754">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ffdb1-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ffdb1-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ffdb1-756">平台</span><span class="sxs-lookup"><span data-stu-id="ffdb1-756">Platform</span></span></th>
    <th><span data-ttu-id="ffdb1-757">扩展点</span><span class="sxs-lookup"><span data-stu-id="ffdb1-757">Extension points</span></span></th>
    <th><span data-ttu-id="ffdb1-758">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="ffdb1-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-760">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ffdb1-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="ffdb1-761">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-761">- Content</span></span><br><span data-ttu-id="ffdb1-762">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-762">
         - TaskPane</span></span><br><span data-ttu-id="ffdb1-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ffdb1-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-768">- ActiveView</span></span><br><span data-ttu-id="ffdb1-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-769">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-770">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-771">
         - File</span></span><br><span data-ttu-id="ffdb1-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-772">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-773">
         - Selection</span></span><br><span data-ttu-id="ffdb1-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-774">
         - Settings</span></span><br><span data-ttu-id="ffdb1-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-776">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-776">Office on Windows</span></span><br><span data-ttu-id="ffdb1-777">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-778">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-778">- Content</span></span><br><span data-ttu-id="ffdb1-779">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-779">
         - TaskPane</span></span><br><span data-ttu-id="ffdb1-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ffdb1-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-785">- ActiveView</span></span><br><span data-ttu-id="ffdb1-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-786">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-787">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-788">
         - File</span></span><br><span data-ttu-id="ffdb1-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-789">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-790">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-790">
         - Selection</span></span><br><span data-ttu-id="ffdb1-791">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-791">
         - Settings</span></span><br><span data-ttu-id="ffdb1-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-793">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-793">Office 2019 on Windows</span></span><br><span data-ttu-id="ffdb1-794">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-795">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-795">- Content</span></span><br><span data-ttu-id="ffdb1-796">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-796">
         - TaskPane</span></span><br><span data-ttu-id="ffdb1-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-800">- ActiveView</span></span><br><span data-ttu-id="ffdb1-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-801">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-802">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-803">
         - File</span></span><br><span data-ttu-id="ffdb1-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-804">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-805">
         - Selection</span></span><br><span data-ttu-id="ffdb1-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-806">
         - Settings</span></span><br><span data-ttu-id="ffdb1-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-808">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-808">Office 2016 on Windows</span></span><br><span data-ttu-id="ffdb1-809">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-810">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-810">- Content</span></span><br><span data-ttu-id="ffdb1-811">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ffdb1-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-814">- ActiveView</span></span><br><span data-ttu-id="ffdb1-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-815">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-816">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-817">
         - File</span></span><br><span data-ttu-id="ffdb1-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-818">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-819">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-819">
         - Selection</span></span><br><span data-ttu-id="ffdb1-820">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-820">
         - Settings</span></span><br><span data-ttu-id="ffdb1-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-822">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ffdb1-822">Office 2013 on Windows</span></span><br><span data-ttu-id="ffdb1-823">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-824">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-824">- Content</span></span><br><span data-ttu-id="ffdb1-825">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="ffdb1-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ffdb1-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-828">- ActiveView</span></span><br><span data-ttu-id="ffdb1-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-829">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-830">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-831">
         - File</span></span><br><span data-ttu-id="ffdb1-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-832">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-833">
         - Selection</span></span><br><span data-ttu-id="ffdb1-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-834">
         - Settings</span></span><br><span data-ttu-id="ffdb1-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-836">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-836">Office on iPad</span></span><br><span data-ttu-id="ffdb1-837">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-838">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-838">- Content</span></span><br><span data-ttu-id="ffdb1-839">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-843">- ActiveView</span></span><br><span data-ttu-id="ffdb1-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-844">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-845">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-846">
         - File</span></span><br><span data-ttu-id="ffdb1-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-847">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-848">
         - Selection</span></span><br><span data-ttu-id="ffdb1-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-849">
         - Settings</span></span><br><span data-ttu-id="ffdb1-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-851">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ffdb1-851">Office on Mac</span></span><br><span data-ttu-id="ffdb1-852">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ffdb1-853">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-853">- Content</span></span><br><span data-ttu-id="ffdb1-854">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-854">
         - TaskPane</span></span><br><span data-ttu-id="ffdb1-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ffdb1-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-860">- ActiveView</span></span><br><span data-ttu-id="ffdb1-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-861">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-862">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-863">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-863">
         - File</span></span><br><span data-ttu-id="ffdb1-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-864">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-865">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-865">
         - Selection</span></span><br><span data-ttu-id="ffdb1-866">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-866">
         - Settings</span></span><br><span data-ttu-id="ffdb1-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-868">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-868">Office 2019 on Mac</span></span><br><span data-ttu-id="ffdb1-869">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-870">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-870">- Content</span></span><br><span data-ttu-id="ffdb1-871">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-871">
         - TaskPane</span></span><br><span data-ttu-id="ffdb1-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-875">- ActiveView</span></span><br><span data-ttu-id="ffdb1-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-876">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-877">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-878">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-878">
         - File</span></span><br><span data-ttu-id="ffdb1-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-879">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-880">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-880">
         - Selection</span></span><br><span data-ttu-id="ffdb1-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-881">
         - Settings</span></span><br><span data-ttu-id="ffdb1-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-883">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-883">Office 2016 on Mac</span></span><br><span data-ttu-id="ffdb1-884">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-885">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-885">- Content</span></span><br><span data-ttu-id="ffdb1-886">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ffdb1-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ffdb1-889">- ActiveView</span></span><br><span data-ttu-id="ffdb1-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-890">
         - CompressedFile</span></span><br><span data-ttu-id="ffdb1-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-891">
         - DocumentEvents</span></span><br><span data-ttu-id="ffdb1-892">
         - File</span><span class="sxs-lookup"><span data-stu-id="ffdb1-892">
         - File</span></span><br><span data-ttu-id="ffdb1-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ffdb1-893">
         - PdfFile</span></span><br><span data-ttu-id="ffdb1-894">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-894">
         - Selection</span></span><br><span data-ttu-id="ffdb1-895">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-895">
         - Settings</span></span><br><span data-ttu-id="ffdb1-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ffdb1-897">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="ffdb1-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="ffdb1-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="ffdb1-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ffdb1-899">平台</span><span class="sxs-lookup"><span data-stu-id="ffdb1-899">Platform</span></span></th>
    <th><span data-ttu-id="ffdb1-900">扩展点</span><span class="sxs-lookup"><span data-stu-id="ffdb1-900">Extension points</span></span></th>
    <th><span data-ttu-id="ffdb1-901">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="ffdb1-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-903">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ffdb1-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="ffdb1-904">- 内容</span><span class="sxs-lookup"><span data-stu-id="ffdb1-904">- Content</span></span><br><span data-ttu-id="ffdb1-905">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-905">
         - TaskPane</span></span><br><span data-ttu-id="ffdb1-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ffdb1-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ffdb1-910">- DocumentEvents</span></span><br><span data-ttu-id="ffdb1-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="ffdb1-912">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ffdb1-912">
         - Settings</span></span><br><span data-ttu-id="ffdb1-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="ffdb1-914">项目</span><span class="sxs-lookup"><span data-stu-id="ffdb1-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ffdb1-915">平台</span><span class="sxs-lookup"><span data-stu-id="ffdb1-915">Platform</span></span></th>
    <th><span data-ttu-id="ffdb1-916">扩展点</span><span class="sxs-lookup"><span data-stu-id="ffdb1-916">Extension points</span></span></th>
    <th><span data-ttu-id="ffdb1-917">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="ffdb1-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-919">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ffdb1-919">Office 2019 on Windows</span></span><br><span data-ttu-id="ffdb1-920">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-921">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-923">- Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-923">- Selection</span></span><br><span data-ttu-id="ffdb1-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-925">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ffdb1-925">Office 2016 on Windows</span></span><br><span data-ttu-id="ffdb1-926">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-927">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-929">- Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-929">- Selection</span></span><br><span data-ttu-id="ffdb1-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ffdb1-931">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ffdb1-931">Office 2013 on Windows</span></span><br><span data-ttu-id="ffdb1-932">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ffdb1-933">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ffdb1-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ffdb1-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ffdb1-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ffdb1-935">- Selection</span><span class="sxs-lookup"><span data-stu-id="ffdb1-935">- Selection</span></span><br><span data-ttu-id="ffdb1-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ffdb1-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ffdb1-937">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ffdb1-937">See also</span></span>

- [<span data-ttu-id="ffdb1-938">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="ffdb1-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ffdb1-939">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ffdb1-940">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="ffdb1-941">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="ffdb1-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="ffdb1-942">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="ffdb1-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="ffdb1-943">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="ffdb1-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="ffdb1-944">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="ffdb1-945">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="ffdb1-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="ffdb1-946">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ffdb1-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="ffdb1-947">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ffdb1-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="ffdb1-948">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="ffdb1-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="ffdb1-949">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="ffdb1-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)