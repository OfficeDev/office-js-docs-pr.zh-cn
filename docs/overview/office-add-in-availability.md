---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 01/23/2020
localization_priority: Priority
ms.openlocfilehash: b30fe872fd89bb02afac99a7838d43d1fbee5464
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554018"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="ac57b-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="ac57b-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="ac57b-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="ac57b-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="ac57b-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="ac57b-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="ac57b-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="ac57b-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="ac57b-108">Excel</span><span class="sxs-lookup"><span data-stu-id="ac57b-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ac57b-109">平台</span><span class="sxs-lookup"><span data-stu-id="ac57b-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ac57b-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="ac57b-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ac57b-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ac57b-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ac57b-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ac57b-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="ac57b-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-114">- TaskPane</span></span><br><span data-ttu-id="ac57b-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-115">
        - Content</span></span><br><span data-ttu-id="ac57b-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ac57b-116">
        - Custom Functions</span></span><br><span data-ttu-id="ac57b-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac57b-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac57b-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac57b-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac57b-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac57b-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac57b-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac57b-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac57b-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac57b-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac57b-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac57b-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac57b-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="ac57b-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ac57b-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-130">
        - BindingEvents</span></span><br><span data-ttu-id="ac57b-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-131">
        - CompressedFile</span></span><br><span data-ttu-id="ac57b-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-132">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-133">
        - File</span></span><br><span data-ttu-id="ac57b-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-134">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-136">
        - Selection</span></span><br><span data-ttu-id="ac57b-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-137">
        - Settings</span></span><br><span data-ttu-id="ac57b-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-138">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-139">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-140">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-142">Office on Windows</span></span><br><span data-ttu-id="ac57b-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-144">- TaskPane</span></span><br><span data-ttu-id="ac57b-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-145">
        - Content</span></span><br><span data-ttu-id="ac57b-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ac57b-146">
        - Custom Functions</span></span><br><span data-ttu-id="ac57b-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac57b-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac57b-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac57b-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac57b-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac57b-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac57b-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac57b-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac57b-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac57b-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac57b-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac57b-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac57b-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac57b-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ac57b-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-161">
        - BindingEvents</span></span><br><span data-ttu-id="ac57b-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-162">
        - CompressedFile</span></span><br><span data-ttu-id="ac57b-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-163">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-164">
        - File</span></span><br><span data-ttu-id="ac57b-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-165">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-167">
        - Selection</span></span><br><span data-ttu-id="ac57b-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-168">
        - Settings</span></span><br><span data-ttu-id="ac57b-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-169">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-170">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-171">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-173">Office 2019 on Windows</span></span><br><span data-ttu-id="ac57b-174">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac57b-175">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-175">- TaskPane</span></span><br><span data-ttu-id="ac57b-176">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-176">
        - Content</span></span><br><span data-ttu-id="ac57b-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ac57b-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac57b-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac57b-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac57b-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac57b-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac57b-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac57b-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac57b-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac57b-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac57b-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-188">- BindingEvents</span></span><br><span data-ttu-id="ac57b-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-189">
        - CompressedFile</span></span><br><span data-ttu-id="ac57b-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-190">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-191">
        - File</span></span><br><span data-ttu-id="ac57b-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-192">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-194">
        - Selection</span></span><br><span data-ttu-id="ac57b-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-195">
        - Settings</span></span><br><span data-ttu-id="ac57b-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-196">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-197">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-198">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-200">Office 2016 on Windows</span></span><br><span data-ttu-id="ac57b-201">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac57b-202">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-202">- TaskPane</span></span><br><span data-ttu-id="ac57b-203">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-203">
        - Content</span></span></td>
    <td><span data-ttu-id="ac57b-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac57b-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac57b-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac57b-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac57b-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-207">- BindingEvents</span></span><br><span data-ttu-id="ac57b-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-208">
        - CompressedFile</span></span><br><span data-ttu-id="ac57b-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-209">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-210">
        - File</span></span><br><span data-ttu-id="ac57b-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-211">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-213">
        - Selection</span></span><br><span data-ttu-id="ac57b-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-214">
        - Settings</span></span><br><span data-ttu-id="ac57b-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-215">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-216">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-217">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ac57b-219">Office 2013 on Windows</span></span><br><span data-ttu-id="ac57b-220">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac57b-221">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-221">
        - TaskPane</span></span><br><span data-ttu-id="ac57b-222">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="ac57b-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac57b-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac57b-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac57b-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-225">
        - BindingEvents</span></span><br><span data-ttu-id="ac57b-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-226">
        - CompressedFile</span></span><br><span data-ttu-id="ac57b-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-227">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-228">
        - File</span></span><br><span data-ttu-id="ac57b-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-229">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-231">
        - Selection</span></span><br><span data-ttu-id="ac57b-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-232">
        - Settings</span></span><br><span data-ttu-id="ac57b-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-233">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-234">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-235">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-237">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-237">Office on iPad</span></span><br><span data-ttu-id="ac57b-238">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac57b-239">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-239">- TaskPane</span></span><br><span data-ttu-id="ac57b-240">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-240">
        - Content</span></span></td>
    <td><span data-ttu-id="ac57b-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac57b-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac57b-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac57b-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac57b-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac57b-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac57b-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac57b-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac57b-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac57b-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac57b-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac57b-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-253">- BindingEvents</span></span><br><span data-ttu-id="ac57b-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-254">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-255">
        - File</span></span><br><span data-ttu-id="ac57b-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-256">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-258">
        - Selection</span></span><br><span data-ttu-id="ac57b-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-259">
        - Settings</span></span><br><span data-ttu-id="ac57b-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-260">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-261">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-262">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-264">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-264">Office on Mac</span></span><br><span data-ttu-id="ac57b-265">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac57b-266">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-266">- TaskPane</span></span><br><span data-ttu-id="ac57b-267">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-267">
        - Content</span></span><br><span data-ttu-id="ac57b-268">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ac57b-268">
        - Custom Functions</span></span><br><span data-ttu-id="ac57b-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ac57b-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac57b-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac57b-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac57b-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac57b-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac57b-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac57b-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac57b-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac57b-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac57b-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac57b-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac57b-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ac57b-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-283">- BindingEvents</span></span><br><span data-ttu-id="ac57b-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-284">
        - CompressedFile</span></span><br><span data-ttu-id="ac57b-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-285">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-286">
        - File</span></span><br><span data-ttu-id="ac57b-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-287">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-289">
        - PdfFile</span></span><br><span data-ttu-id="ac57b-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-290">
        - Selection</span></span><br><span data-ttu-id="ac57b-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-291">
        - Settings</span></span><br><span data-ttu-id="ac57b-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-292">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-293">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-294">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-296">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-296">Office 2019 on Mac</span></span><br><span data-ttu-id="ac57b-297">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac57b-298">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-298">- TaskPane</span></span><br><span data-ttu-id="ac57b-299">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-299">
        - Content</span></span><br><span data-ttu-id="ac57b-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ac57b-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac57b-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac57b-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac57b-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac57b-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac57b-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac57b-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac57b-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac57b-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac57b-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-311">- BindingEvents</span></span><br><span data-ttu-id="ac57b-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-312">
        - CompressedFile</span></span><br><span data-ttu-id="ac57b-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-313">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-314">
        - File</span></span><br><span data-ttu-id="ac57b-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-315">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-317">
        - PdfFile</span></span><br><span data-ttu-id="ac57b-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-318">
        - Selection</span></span><br><span data-ttu-id="ac57b-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-319">
        - Settings</span></span><br><span data-ttu-id="ac57b-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-320">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-321">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-322">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-324">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-324">Office 2016 on Mac</span></span><br><span data-ttu-id="ac57b-325">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac57b-326">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-326">- TaskPane</span></span><br><span data-ttu-id="ac57b-327">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-327">
        - Content</span></span></td>
    <td><span data-ttu-id="ac57b-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac57b-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac57b-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac57b-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac57b-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-331">- BindingEvents</span></span><br><span data-ttu-id="ac57b-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-332">
        - CompressedFile</span></span><br><span data-ttu-id="ac57b-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-333">
        - DocumentEvents</span></span><br><span data-ttu-id="ac57b-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-334">
        - File</span></span><br><span data-ttu-id="ac57b-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-335">
        - MatrixBindings</span></span><br><span data-ttu-id="ac57b-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-337">
        - PdfFile</span></span><br><span data-ttu-id="ac57b-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-338">
        - Selection</span></span><br><span data-ttu-id="ac57b-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-339">
        - Settings</span></span><br><span data-ttu-id="ac57b-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-340">
        - TableBindings</span></span><br><span data-ttu-id="ac57b-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-341">
        - TableCoercion</span></span><br><span data-ttu-id="ac57b-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-342">
        - TextBindings</span></span><br><span data-ttu-id="ac57b-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ac57b-344">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="ac57b-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="ac57b-345">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="ac57b-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ac57b-346">平台</span><span class="sxs-lookup"><span data-stu-id="ac57b-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ac57b-347">扩展点</span><span class="sxs-lookup"><span data-stu-id="ac57b-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ac57b-348">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ac57b-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ac57b-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-350">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ac57b-350">Office on the web</span></span></td>
    <td><span data-ttu-id="ac57b-351">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ac57b-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ac57b-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-353">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-353">Office on Windows</span></span><br><span data-ttu-id="ac57b-354">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac57b-355">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ac57b-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ac57b-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="ac57b-357">Office for Mac</span></span><br><span data-ttu-id="ac57b-358">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="ac57b-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="ac57b-359">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="ac57b-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ac57b-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="ac57b-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="ac57b-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac57b-362">平台</span><span class="sxs-lookup"><span data-stu-id="ac57b-362">Platform</span></span></th>
    <th><span data-ttu-id="ac57b-363">扩展点</span><span class="sxs-lookup"><span data-stu-id="ac57b-363">Extension points</span></span></th>
    <th><span data-ttu-id="ac57b-364">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac57b-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ac57b-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-366">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ac57b-366">Office on the web</span></span><br><span data-ttu-id="ac57b-367">（新式）</span><span class="sxs-lookup"><span data-stu-id="ac57b-367">(modern)</span></span></td>
    <td> <span data-ttu-id="ac57b-368">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-368">- Mail Read</span></span><br><span data-ttu-id="ac57b-369">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-369">
      - Mail Compose</span></span><br><span data-ttu-id="ac57b-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac57b-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac57b-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac57b-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ac57b-379">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-380">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ac57b-380">Office on the web</span></span><br><span data-ttu-id="ac57b-381">（经典）</span><span class="sxs-lookup"><span data-stu-id="ac57b-381">(classic)</span></span></td>
    <td> <span data-ttu-id="ac57b-382">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-382">- Mail Read</span></span><br><span data-ttu-id="ac57b-383">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-383">
      - Mail Compose</span></span><br><span data-ttu-id="ac57b-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac57b-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ac57b-391">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-392">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-392">Office on Windows</span></span><br><span data-ttu-id="ac57b-393">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-394">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-394">- Mail Read</span></span><br><span data-ttu-id="ac57b-395">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-395">
      - Mail Compose</span></span><br><span data-ttu-id="ac57b-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ac57b-397">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="ac57b-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ac57b-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac57b-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac57b-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac57b-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ac57b-406">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-407">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-407">Office 2019 on Windows</span></span><br><span data-ttu-id="ac57b-408">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-409">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-409">- Mail Read</span></span><br><span data-ttu-id="ac57b-410">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-410">
      - Mail Compose</span></span><br><span data-ttu-id="ac57b-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ac57b-412">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="ac57b-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ac57b-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac57b-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac57b-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ac57b-420">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-421">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-421">Office 2016 on Windows</span></span><br><span data-ttu-id="ac57b-422">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-423">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-423">- Mail Read</span></span><br><span data-ttu-id="ac57b-424">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-424">
      - Mail Compose</span></span><br><span data-ttu-id="ac57b-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ac57b-426">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="ac57b-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ac57b-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ac57b-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ac57b-431">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-432">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ac57b-432">Office 2013 on Windows</span></span><br><span data-ttu-id="ac57b-433">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-434">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-434">- Mail Read</span></span><br><span data-ttu-id="ac57b-435">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="ac57b-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="ac57b-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="ac57b-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ac57b-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ac57b-440">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-441">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-441">Office on iOS</span></span><br><span data-ttu-id="ac57b-442">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-443">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-443">- Mail Read</span></span><br><span data-ttu-id="ac57b-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ac57b-450">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-451">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-451">Office on Mac</span></span><br><span data-ttu-id="ac57b-452">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-453">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-453">- Mail Read</span></span><br><span data-ttu-id="ac57b-454">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-454">
      - Mail Compose</span></span><br><span data-ttu-id="ac57b-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac57b-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac57b-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac57b-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ac57b-464">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-465">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-465">Office 2019 on Mac</span></span><br><span data-ttu-id="ac57b-466">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-467">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-467">- Mail Read</span></span><br><span data-ttu-id="ac57b-468">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-468">
      - Mail Compose</span></span><br><span data-ttu-id="ac57b-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac57b-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ac57b-476">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-477">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-477">Office 2016 on Mac</span></span><br><span data-ttu-id="ac57b-478">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-479">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-479">- Mail Read</span></span><br><span data-ttu-id="ac57b-480">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ac57b-480">
      - Mail Compose</span></span><br><span data-ttu-id="ac57b-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac57b-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ac57b-488">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-489">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-489">Office on Android</span></span><br><span data-ttu-id="ac57b-490">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-491">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ac57b-491">- Mail Read</span></span><br><span data-ttu-id="ac57b-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac57b-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac57b-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac57b-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac57b-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ac57b-498">不可用</span><span class="sxs-lookup"><span data-stu-id="ac57b-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="ac57b-499">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="ac57b-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ac57b-500">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="ac57b-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="ac57b-501">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="ac57b-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="ac57b-502">Word</span><span class="sxs-lookup"><span data-stu-id="ac57b-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac57b-503">平台</span><span class="sxs-lookup"><span data-stu-id="ac57b-503">Platform</span></span></th>
    <th><span data-ttu-id="ac57b-504">扩展点</span><span class="sxs-lookup"><span data-stu-id="ac57b-504">Extension points</span></span></th>
    <th><span data-ttu-id="ac57b-505">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac57b-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ac57b-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-507">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ac57b-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="ac57b-508">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-508">- TaskPane</span></span><br><span data-ttu-id="ac57b-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac57b-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac57b-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac57b-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac57b-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac57b-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-516">- BindingEvents</span></span><br><span data-ttu-id="ac57b-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-518">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-519">
         - File</span></span><br><span data-ttu-id="ac57b-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-521">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-524">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-525">
         - Selection</span></span><br><span data-ttu-id="ac57b-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-526">
         - Settings</span></span><br><span data-ttu-id="ac57b-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-527">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-528">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-529">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-530">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-532">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-532">Office on Windows</span></span><br><span data-ttu-id="ac57b-533">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-534">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-534">- TaskPane</span></span><br><span data-ttu-id="ac57b-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac57b-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac57b-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac57b-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac57b-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac57b-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-542">- BindingEvents</span></span><br><span data-ttu-id="ac57b-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-543">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-545">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-546">
         - File</span></span><br><span data-ttu-id="ac57b-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-548">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-551">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-552">
         - Selection</span></span><br><span data-ttu-id="ac57b-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-553">
         - Settings</span></span><br><span data-ttu-id="ac57b-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-554">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-555">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-556">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-557">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-559">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-559">Office 2019 on Windows</span></span><br><span data-ttu-id="ac57b-560">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-561">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-561">- TaskPane</span></span><br><span data-ttu-id="ac57b-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac57b-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac57b-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac57b-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-568">- BindingEvents</span></span><br><span data-ttu-id="ac57b-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-569">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-571">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-572">
         - File</span></span><br><span data-ttu-id="ac57b-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-574">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-577">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-578">
         - Selection</span></span><br><span data-ttu-id="ac57b-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-579">
         - Settings</span></span><br><span data-ttu-id="ac57b-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-580">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-581">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-582">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-583">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-585">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-585">Office 2016 on Windows</span></span><br><span data-ttu-id="ac57b-586">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-587">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac57b-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac57b-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac57b-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-591">- BindingEvents</span></span><br><span data-ttu-id="ac57b-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-592">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-594">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-595">
         - File</span></span><br><span data-ttu-id="ac57b-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-597">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-600">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-601">
         - Selection</span></span><br><span data-ttu-id="ac57b-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-602">
         - Settings</span></span><br><span data-ttu-id="ac57b-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-603">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-604">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-605">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-606">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-608">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ac57b-608">Office 2013 on Windows</span></span><br><span data-ttu-id="ac57b-609">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-610">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac57b-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac57b-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-613">- BindingEvents</span></span><br><span data-ttu-id="ac57b-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-614">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-616">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-617">
         - File</span></span><br><span data-ttu-id="ac57b-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-619">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-622">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-623">
         - Selection</span></span><br><span data-ttu-id="ac57b-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-624">
         - Settings</span></span><br><span data-ttu-id="ac57b-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-625">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-626">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-627">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-628">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-630">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-630">Office on iPad</span></span><br><span data-ttu-id="ac57b-631">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-632">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac57b-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac57b-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac57b-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ac57b-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-638">- BindingEvents</span></span><br><span data-ttu-id="ac57b-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-639">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-641">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-642">
         - File</span></span><br><span data-ttu-id="ac57b-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-644">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-647">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-648">
         - Selection</span></span><br><span data-ttu-id="ac57b-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-649">
         - Settings</span></span><br><span data-ttu-id="ac57b-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-650">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-651">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-652">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-653">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-655">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-655">Office on Mac</span></span><br><span data-ttu-id="ac57b-656">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-657">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-657">- TaskPane</span></span><br><span data-ttu-id="ac57b-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac57b-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac57b-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac57b-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac57b-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="ac57b-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-665">- BindingEvents</span></span><br><span data-ttu-id="ac57b-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-666">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-668">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-669">
         - File</span></span><br><span data-ttu-id="ac57b-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-671">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-674">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-675">
         - Selection</span></span><br><span data-ttu-id="ac57b-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-676">
         - Settings</span></span><br><span data-ttu-id="ac57b-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-677">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-678">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-679">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-680">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-682">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-682">Office 2019 on Mac</span></span><br><span data-ttu-id="ac57b-683">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-684">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-684">- TaskPane</span></span><br><span data-ttu-id="ac57b-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac57b-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac57b-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac57b-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ac57b-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-691">- BindingEvents</span></span><br><span data-ttu-id="ac57b-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-692">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-694">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-695">
         - File</span></span><br><span data-ttu-id="ac57b-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-697">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-700">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-701">
         - Selection</span></span><br><span data-ttu-id="ac57b-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-702">
         - Settings</span></span><br><span data-ttu-id="ac57b-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-703">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-704">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-705">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-706">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-708">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-708">Office 2016 on Mac</span></span><br><span data-ttu-id="ac57b-709">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-710">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac57b-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac57b-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac57b-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-714">- BindingEvents</span></span><br><span data-ttu-id="ac57b-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-715">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac57b-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac57b-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-717">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-718">
         - File</span></span><br><span data-ttu-id="ac57b-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-720">
         - MatrixBindings</span></span><br><span data-ttu-id="ac57b-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac57b-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac57b-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-723">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-724">
         - Selection</span></span><br><span data-ttu-id="ac57b-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-725">
         - Settings</span></span><br><span data-ttu-id="ac57b-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-726">
         - TableBindings</span></span><br><span data-ttu-id="ac57b-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-727">
         - TableCoercion</span></span><br><span data-ttu-id="ac57b-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac57b-728">
         - TextBindings</span></span><br><span data-ttu-id="ac57b-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-729">
         - TextCoercion</span></span><br><span data-ttu-id="ac57b-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="ac57b-731">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="ac57b-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ac57b-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ac57b-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac57b-733">平台</span><span class="sxs-lookup"><span data-stu-id="ac57b-733">Platform</span></span></th>
    <th><span data-ttu-id="ac57b-734">扩展点</span><span class="sxs-lookup"><span data-stu-id="ac57b-734">Extension points</span></span></th>
    <th><span data-ttu-id="ac57b-735">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac57b-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ac57b-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-737">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ac57b-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="ac57b-738">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-738">- Content</span></span><br><span data-ttu-id="ac57b-739">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-739">
         - TaskPane</span></span><br><span data-ttu-id="ac57b-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac57b-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac57b-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac57b-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-745">- ActiveView</span></span><br><span data-ttu-id="ac57b-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-746">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-747">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-748">
         - File</span></span><br><span data-ttu-id="ac57b-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-749">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-750">
         - Selection</span></span><br><span data-ttu-id="ac57b-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-751">
         - Settings</span></span><br><span data-ttu-id="ac57b-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-753">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-753">Office on Windows</span></span><br><span data-ttu-id="ac57b-754">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-755">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-755">- Content</span></span><br><span data-ttu-id="ac57b-756">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-756">
         - TaskPane</span></span><br><span data-ttu-id="ac57b-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac57b-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac57b-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac57b-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-762">- ActiveView</span></span><br><span data-ttu-id="ac57b-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-763">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-764">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-765">
         - File</span></span><br><span data-ttu-id="ac57b-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-766">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-767">
         - Selection</span></span><br><span data-ttu-id="ac57b-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-768">
         - Settings</span></span><br><span data-ttu-id="ac57b-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-770">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-770">Office 2019 on Windows</span></span><br><span data-ttu-id="ac57b-771">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-772">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-772">- Content</span></span><br><span data-ttu-id="ac57b-773">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-773">
         - TaskPane</span></span><br><span data-ttu-id="ac57b-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-777">- ActiveView</span></span><br><span data-ttu-id="ac57b-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-778">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-779">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-780">
         - File</span></span><br><span data-ttu-id="ac57b-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-781">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-782">
         - Selection</span></span><br><span data-ttu-id="ac57b-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-783">
         - Settings</span></span><br><span data-ttu-id="ac57b-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-785">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-785">Office 2016 on Windows</span></span><br><span data-ttu-id="ac57b-786">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-787">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-787">- Content</span></span><br><span data-ttu-id="ac57b-788">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac57b-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac57b-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-791">- ActiveView</span></span><br><span data-ttu-id="ac57b-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-792">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-793">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-794">
         - File</span></span><br><span data-ttu-id="ac57b-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-795">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-796">
         - Selection</span></span><br><span data-ttu-id="ac57b-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-797">
         - Settings</span></span><br><span data-ttu-id="ac57b-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-799">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ac57b-799">Office 2013 on Windows</span></span><br><span data-ttu-id="ac57b-800">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-801">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-801">- Content</span></span><br><span data-ttu-id="ac57b-802">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="ac57b-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac57b-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac57b-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-805">- ActiveView</span></span><br><span data-ttu-id="ac57b-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-806">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-807">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-808">
         - File</span></span><br><span data-ttu-id="ac57b-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-809">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-810">
         - Selection</span></span><br><span data-ttu-id="ac57b-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-811">
         - Settings</span></span><br><span data-ttu-id="ac57b-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-813">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-813">Office on iPad</span></span><br><span data-ttu-id="ac57b-814">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-815">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-815">- Content</span></span><br><span data-ttu-id="ac57b-816">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac57b-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-820">- ActiveView</span></span><br><span data-ttu-id="ac57b-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-821">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-822">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-823">
         - File</span></span><br><span data-ttu-id="ac57b-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-824">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-825">
         - Selection</span></span><br><span data-ttu-id="ac57b-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-826">
         - Settings</span></span><br><span data-ttu-id="ac57b-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-828">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ac57b-828">Office on Mac</span></span><br><span data-ttu-id="ac57b-829">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ac57b-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac57b-830">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-830">- Content</span></span><br><span data-ttu-id="ac57b-831">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-831">
         - TaskPane</span></span><br><span data-ttu-id="ac57b-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac57b-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac57b-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac57b-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-837">- ActiveView</span></span><br><span data-ttu-id="ac57b-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-838">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-839">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-840">
         - File</span></span><br><span data-ttu-id="ac57b-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-841">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-842">
         - Selection</span></span><br><span data-ttu-id="ac57b-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-843">
         - Settings</span></span><br><span data-ttu-id="ac57b-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-845">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-845">Office 2019 on Mac</span></span><br><span data-ttu-id="ac57b-846">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-847">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-847">- Content</span></span><br><span data-ttu-id="ac57b-848">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-848">
         - TaskPane</span></span><br><span data-ttu-id="ac57b-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-852">- ActiveView</span></span><br><span data-ttu-id="ac57b-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-853">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-854">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-855">
         - File</span></span><br><span data-ttu-id="ac57b-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-856">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-857">
         - Selection</span></span><br><span data-ttu-id="ac57b-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-858">
         - Settings</span></span><br><span data-ttu-id="ac57b-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-860">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-860">Office 2016 on Mac</span></span><br><span data-ttu-id="ac57b-861">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-862">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-862">- Content</span></span><br><span data-ttu-id="ac57b-863">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac57b-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac57b-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac57b-866">- ActiveView</span></span><br><span data-ttu-id="ac57b-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-867">
         - CompressedFile</span></span><br><span data-ttu-id="ac57b-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-868">
         - DocumentEvents</span></span><br><span data-ttu-id="ac57b-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="ac57b-869">
         - File</span></span><br><span data-ttu-id="ac57b-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac57b-870">
         - PdfFile</span></span><br><span data-ttu-id="ac57b-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-871">
         - Selection</span></span><br><span data-ttu-id="ac57b-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-872">
         - Settings</span></span><br><span data-ttu-id="ac57b-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ac57b-874">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="ac57b-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="ac57b-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="ac57b-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac57b-876">平台</span><span class="sxs-lookup"><span data-stu-id="ac57b-876">Platform</span></span></th>
    <th><span data-ttu-id="ac57b-877">扩展点</span><span class="sxs-lookup"><span data-stu-id="ac57b-877">Extension points</span></span></th>
    <th><span data-ttu-id="ac57b-878">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac57b-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ac57b-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-880">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ac57b-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="ac57b-881">- 内容</span><span class="sxs-lookup"><span data-stu-id="ac57b-881">- Content</span></span><br><span data-ttu-id="ac57b-882">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-882">
         - TaskPane</span></span><br><span data-ttu-id="ac57b-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac57b-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ac57b-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac57b-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac57b-887">- DocumentEvents</span></span><br><span data-ttu-id="ac57b-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac57b-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ac57b-889">
         - Settings</span></span><br><span data-ttu-id="ac57b-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="ac57b-891">项目</span><span class="sxs-lookup"><span data-stu-id="ac57b-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac57b-892">平台</span><span class="sxs-lookup"><span data-stu-id="ac57b-892">Platform</span></span></th>
    <th><span data-ttu-id="ac57b-893">扩展点</span><span class="sxs-lookup"><span data-stu-id="ac57b-893">Extension points</span></span></th>
    <th><span data-ttu-id="ac57b-894">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac57b-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ac57b-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-896">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ac57b-896">Office 2019 on Windows</span></span><br><span data-ttu-id="ac57b-897">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-898">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-900">- Selection</span></span><br><span data-ttu-id="ac57b-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-902">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ac57b-902">Office 2016 on Windows</span></span><br><span data-ttu-id="ac57b-903">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-904">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-906">- Selection</span></span><br><span data-ttu-id="ac57b-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac57b-908">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ac57b-908">Office 2013 on Windows</span></span><br><span data-ttu-id="ac57b-909">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ac57b-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac57b-910">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ac57b-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac57b-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac57b-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac57b-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="ac57b-912">- Selection</span></span><br><span data-ttu-id="ac57b-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac57b-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ac57b-914">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ac57b-914">See also</span></span>

- [<span data-ttu-id="ac57b-915">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="ac57b-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ac57b-916">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ac57b-917">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="ac57b-918">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="ac57b-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="ac57b-919">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="ac57b-919">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="ac57b-920">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="ac57b-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="ac57b-921">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="ac57b-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="ac57b-922">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="ac57b-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="ac57b-923">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ac57b-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="ac57b-924">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ac57b-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="ac57b-925">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="ac57b-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="ac57b-926">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="ac57b-926">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)